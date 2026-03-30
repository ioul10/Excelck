"""
core/pdf_parser.py  v5 — Parseur universel MCN
Supporte tous les formats PDF rencontrés:
  - DGI officiel (7 pages): labels fragmentés, 4 colonnes x≈190-455
  - Etats/BORJ (5 pages): séparateur †, 3-4 colonnes
  - SAPST Liasse (5 pages): points comme milliers, digit séparé
  - SGTM/Bilan-2017 (5 pages): format standard espace-milliers

Algorithme:
  1. extract_words() → mots avec (x0, x1, top)
  2. Regrouper par Y (tolérance 6pt)
  3. Détecter seuil label/nombre + zones de colonnes
  4. Fusionner tokens numériques adjacents (gap < 6pt)
  5. Parser nombres (formats: espaces, points, †)
  6. Assigner aux colonnes par ordre X
"""

import re
from collections import defaultdict, Counter
import pdfplumber
from utils.logger import get_logger

logger = get_logger(__name__)

# ── Patterns de skip ──────────────────────────────────────────────────────────
SKIP_LABELS_RE = re.compile(
    r'^(tableau|bilan\s*\(|compte de produits|b\s+i\s+l\s+a\s+n|'
    r'modèle|exercice du|identifiant|raison sociale|'
    r'1\)variation|2\)achats|nb\s*:|cadre réservé|'
    r'numéro d.enregistrement|signature|'
    r'\(1\)capital|\(2\)bénéf|pages|'
    r'a\.i\.|a\.c\.|f\.p\.|p\.c\.|t\.\s*:|n\.c\.\s*:|e\.\s*:|f\.\s*:|'
    r'conforme à la déclaration|etat sous référence)', re.I
)

SKIP_SINGLE = {
    'brut', 'net', 'amortissements', 'provisions', 'exercice',
    'precedent', 'précédent', 'designation', 'désignation',
    'operations', 'opérations', 'propres', 'concernant',
    'totaux', 'de', 'du', 'et', 'en', 'la', 'le', 'les',
    'nature', 'exploitation', 'financier', 'courant',
}

RESULT_SKIP_RE = re.compile(
    r'^(résultat d.exploitation|résultat financier|résultat courant(?!\s*\()|'
    r'résultat non courant|résultat avant impôts|'
    r'total i\b|total ii\b|total iii\b|total des |total général|'
    r'produits d.exploitation$|charges d.exploitation$|'
    r'charges financières?$|produits financiers?$|'
    r'produits non courants?$|charges non courants?$)', re.I
)

# Labels passif uniquement (pour filtrer contaminations actif)
PASSIF_ONLY_MARKERS = {
    'capitaux propres', 'capital social', 'capital appele',
    'prime emission', 'ecart reevaluation', 'reserve legale', 'reserves legales',
    'report nouveau', 'reports nouveau', 'resultat instance',
    'resultat net exercice', 'resultat net de l exercice',
    'capitaux propres assimil', 'subventions investissement passif',
    'subvention d investissement', 'subventions d invertissement',
    'provisions reglementees', 'dettes financement',
    'emprunts obligataires', 'provisions durables',
    'fournisseurs comptes rattaches passif', 'clients crediteurs avances passif',
    'autres creanciers', 'comptes regularisation passif',
    'autres provisions risques charges',
    'credits escompte', 'credits tresorerie', 'banques soldes crediteurs',
    'ecarts conversion passif',
    'fournisseurs et comptes rattaches',
}


class PDFParser:

    def __init__(self, pdf_path: str):
        self.path    = pdf_path
        self.pdf     = pdfplumber.open(pdf_path)
        self.pages   = self.pdf.pages
        self.n_pages = len(self.pages)
        logger.info(f"PDF chargé : {pdf_path} — {self.n_pages} pages")

    def parse(self) -> dict:
        result = {
            "info":          self._parse_info(),
            "actif_values":  self._extract_section(pages=[1, 2],    mode='actif'),
            "passif_values": self._extract_section(pages=[2, 3],    mode='passif'),
            "cpc_values":    self._extract_section(pages=[3, 4, 5], mode='cpc'),
        }
        self._enrich_passif(result)
        return result

    # ── Infos générales ────────────────────────────────────────────────────────

    def _parse_info(self) -> dict:
        info = {}
        text0 = self._page_text(0)

        for key, pat in [
            ("raison_sociale",       r"Raison\s+[Ss]ociale\s*[:\-]?\s*(.+)"),
            ("taxe_professionnelle",  r"(?:Taxe|Art\.)\s+[Pp]rof[a-z.]*\s*[:\-]?\s*([\d\s]+)"),
            ("identifiant_fiscal",    r"[Ii]dentifiant\s+[Ff]iscal\s*[:\-]?\s*(\d+)"),
            ("adresse",               r"[Aa]dresse\s*[:\-]?\s*(.+)"),
        ]:
            m = re.search(pat, text0, re.IGNORECASE)
            if m:
                info[key] = m.group(1).strip()

        # Fallback raison sociale
        if not info.get("raison_sociale"):
            for i in range(min(3, self.n_pages)):
                t = self._page_text(i)
                # Format DGI: "Raison Sociale : BEST BISCUITS MAROC"
                m = re.search(r"Raison Sociale\s*[:\-]\s*([A-Z][^\n]{3,80})", t)
                if m:
                    info["raison_sociale"] = m.group(1).strip()
                    break
                # Format standard
                m = re.search(r"((?:SOCIETE|SOCIÉTÉ|AGENCE|OFFICE|DIRECTION|S\.A\.|SARL|BEST|BORJ)[^\n]{3,80})", t)
                if m:
                    info["raison_sociale"] = m.group(1).strip()
                    break

        # Identifiant fiscal fallback
        if not info.get("identifiant_fiscal"):
            for line in text0.split('\n'):
                if re.match(r'^\d{6,10}$', line.strip()):
                    info["identifiant_fiscal"] = line.strip()
                    break

        # Exercice
        for i in range(min(3, self.n_pages)):
            t = self._page_text(i)
            m = re.search(r"(\d{2}/\d{2}/\d{4})\s+au\s+(\d{2}/\d{2}/\d{4})", t)
            if not m:
                m = re.search(r"(?:période du|au titre de la période du|du\s*:?\s*)(\d{2}/\d{2}/\d{4})\s+(?:au|AU)\s*:?\s*(\d{2}/\d{2}/\d{4})", t, re.I)
            if m:
                info["exercice"]       = f"Du {m.group(1)} au {m.group(2)}"
                info["exercice_debut"] = m.group(1)
                info["exercice_fin"]   = m.group(2)
                break

        # Date déclaration
        for i in range(min(2, self.n_pages)):
            t = self._page_text(i)
            m = re.search(r"(?:le\s+)?(\d{2}/\d{2}/\d{4})\s+\d{2}:\d{2}", t)
            if not m:
                m = re.search(r"[Ll]e\s+(\d{2}/\d{2}/\d{4})", t)
            if m:
                info["date_declaration"] = m.group(1)
                break

        for k in ["exercice", "exercice_fin", "exercice_debut", "date_declaration",
                  "raison_sociale", "identifiant_fiscal", "adresse", "taxe_professionnelle"]:
            info.setdefault(k, "")
        info["pages"] = self.n_pages
        return info

    # ── Extraction universelle ─────────────────────────────────────────────────

    def _extract_section(self, pages: list, mode: str) -> dict:
        values = {}
        for page_idx in pages:
            if page_idx >= self.n_pages:
                continue
            text = self._page_text(page_idx)
            if not self._page_matches_mode(text, mode):
                continue
            page_values = self._extract_page(page_idx, mode)
            for label, vals in page_values.items():
                if label not in values:
                    values[label] = vals
                else:
                    # Enrichir avec valeurs manquantes
                    ex = values[label]
                    ml = max(len(ex), len(vals))
                    merged = []
                    for i in range(ml):
                        ev = ex[i]   if i < len(ex)   else None
                        nv = vals[i] if i < len(vals) else None
                        merged.append(ev if ev is not None else nv)
                    values[label] = merged

        logger.info(f"{mode}: {len(values)} postes extraits")
        return values

    def _page_matches_mode(self, text: str, mode: str) -> bool:
        t = text.lower()
        if mode == 'actif':
            return any(k in t for k in ['actif', 'immobilisations', 'stocks', 'créances', 'bilan'])
        if mode == 'passif':
            return any(k in t for k in ['passif', 'capitaux propres', 'capital', 'dettes'])
        if mode == 'cpc':
            return any(k in t for k in ['produits', 'charges', 'exploitation', 'résultat', 'compte de'])
        return True

    def _extract_page(self, page_idx: int, mode: str) -> dict:
        page  = self.pages[page_idx]
        words = page.extract_words(x_tolerance=3, y_tolerance=3)
        if not words:
            return {}

        # Regrouper par Y (tolérance 6pt)
        lines = defaultdict(list)
        for w in words:
            lines[round(w['top'] / 6) * 6].append(w)

        # Détecter le seuil label/nombre
        num_xs = [w['x0'] for w in words
                  if self._is_num_token(w['text']) and w['x0'] > 150]
        if not num_xs:
            return {}

        label_thresh = min(num_xs) - 10

        # Détecter les colonnes
        col_xs = self._detect_columns(num_xs)

        result = {}
        prev_label = None  # Pour les labels sur ligne sans valeur (DGI)

        for y in sorted(lines.keys()):
            row = sorted(lines[y], key=lambda w: w['x0'])

            # Reconstituer le label (gérer fragments DGI)
            label_words = [w for w in row if w['x0'] < label_thresh]
            label = self._reconstruct_label(label_words)

            # Tokens numériques
            num_words = [(w['x0'], w['x1'], w['text']) for w in row
                        if w['x0'] >= label_thresh and self._is_num_token(w['text'])]

            if not label and not num_words:
                continue

            if label and not num_words:
                # Ligne avec label seulement → mémoriser pour la ligne suivante (DGI)
                prev_label = label
                continue

            if not label and num_words and prev_label:
                # Ligne avec valeurs sans label → utiliser le label précédent
                label = prev_label
                prev_label = None
            elif label and num_words:
                prev_label = None

            if not label:
                continue

            if not self._is_valid_label(label, mode):
                continue

            # Fusionner et parser les nombres
            merged = self._merge_and_parse_nums(num_words)
            if not merged:
                continue

            # Assigner aux colonnes
            vals = self._assign_to_cols(merged, col_xs, mode)

            if any(v is not None for v in vals):
                result[label] = vals

        return result

    # ── Détection des colonnes ─────────────────────────────────────────────────

    def _detect_columns(self, num_xs: list) -> list:
        if not num_xs:
            return []
        hist = Counter([round(x / 5) * 5 for x in num_xs])
        significant = sorted([(k, v) for k, v in hist.items() if v >= 2])
        if not significant:
            return []

        cols = []
        grp = [significant[0]]
        for x, c in significant[1:]:
            if x - grp[-1][0] < 25:
                grp.append((x, c))
            else:
                total_x = sum(xi * ci for xi, ci in grp)
                total_c = sum(ci for _, ci in grp)
                cols.append(total_x / total_c)
                grp = [(x, c)]
        total_x = sum(xi * ci for xi, ci in grp)
        total_c = sum(ci for _, ci in grp)
        cols.append(total_x / total_c)
        return sorted(cols)

    # ── Assignation des colonnes ───────────────────────────────────────────────

    def _assign_to_cols(self, merged: list, col_xs: list, mode: str) -> list:
        """
        Assigne les valeurs aux colonnes attendues.
        Actif  : 4 cols idéalement → retourne [brut, amort, net_n1]
        Passif : 2 cols → retourne [val_n, val_n1]
        CPC    : 4 cols idéalement → retourne [propre_n, prec_n, total_n1]
        """
        if not col_xs or not merged:
            return [v for _, v in merged]

        n_cols = len(col_xs)

        # Créer le mapping col_index → valeur
        col_map = {}
        for x, v in merged:
            closest = min(range(n_cols), key=lambda i: abs(col_xs[i] - x))
            col_map[closest] = v

        filled = {k: v for k, v in col_map.items() if v is not None}

        if mode == 'actif':
            if n_cols >= 4:
                # Vérifier si les 2 premières cols sont vides (PDF sans amort)
                if not filled.get(0) and not filled.get(1):
                    return [filled.get(2), None, filled.get(3)]
                # Déterminer l'ordre : Brut|Amort|Net|NetN1 (standard)
                #                   ou Brut|Net|Amort|NetN1 (SGTM)
                # Règle mathématique : Brut = col1 + col2 toujours
                # Si col1 == Brut → Net=Brut, Amort=0=col2
                # Sinon Amort = max(col1, col2) sauf si égaux
                c0 = col_map.get(0)   # Brut
                c1 = col_map.get(1)   # Amort ou Net selon format
                c2 = col_map.get(2)   # Net ou Amort selon format
                c3 = col_map.get(3)   # Net N-1
                if c1 is not None and c2 is not None:
                    if c1 == c0:      # Net = Brut → Amort = 0 = c2
                        amort = c2
                    elif c2 == c0:    # rare
                        amort = c1
                    else:
                        amort = max(c1, c2) if c1 != c2 else min(c1, c2)
                else:
                    amort = c1
                return [c0, amort, c3]
            elif n_cols == 3:
                c0 = col_map.get(0)
                c1 = col_map.get(1)
                c2 = col_map.get(2)
                # Si c0==c1 → Net=Brut → Amort=0, NetN1=c2
                if c0 is not None and c1 is not None and c0 == c1:
                    return [c0, 0.0, c2]
                # Calculer les gaps entre colonnes
                gaps = [col_xs[i+1] - col_xs[i] for i in range(n_cols-1)]
                if len(gaps) >= 2 and gaps[0] > gaps[1] * 1.5:
                    return [c0, None, c2]
                else:
                    return [c0, c1, c2]
            elif n_cols == 2:
                return [col_map.get(0), None, col_map.get(1)]
            else:
                return [v for _, v in merged]

        elif mode == 'passif':
            if n_cols >= 4:
                # Vérifier si les 2 premières cols sont vides
                if not filled.get(0) and not filled.get(1):
                    return [filled.get(2), filled.get(3)]
                # Passif n'a que 2 valeurs: prendre les 2 dernières
                sorted_filled = sorted(filled.items())
                if len(sorted_filled) >= 2:
                    return [sorted_filled[-2][1], sorted_filled[-1][1]]
                return [col_map.get(n_cols-2), col_map.get(n_cols-1)]
            elif n_cols >= 2:
                sorted_filled = sorted(filled.items())
                if len(sorted_filled) >= 2:
                    return [sorted_filled[-2][1], sorted_filled[-1][1]]
                elif len(sorted_filled) == 1:
                    return [sorted_filled[0][1], None]
                return [col_map.get(0), col_map.get(1)]
            else:
                return [v for _, v in merged]

        elif mode == 'cpc':
            if n_cols >= 4:
                if not filled.get(0) and not filled.get(1) and (filled.get(2) or filled.get(3)):
                    return [filled.get(2), None, filled.get(3)]
                return [col_map.get(0), col_map.get(1), col_map.get(3)]
            elif n_cols >= 3:
                return [col_map.get(0), col_map.get(1), col_map.get(2)]
            elif n_cols == 2:
                return [col_map.get(0), None, col_map.get(1)]
            else:
                return [v for _, v in merged]

        return [v for _, v in merged]

    # ── Reconstruction du label ────────────────────────────────────────────────

    def _reconstruct_label(self, label_words: list) -> str:
        """
        Reconstruction robuste du label pour tous les formats PDF.

        Règles:
        - Filtrer les préfixes parasites (*, numéros de compte, lettres rotatives)
        - Si deux groupes séparés par gap > 8pt → prendre le DERNIER (closest to nums)
        - gap ≤ 1pt → coller directement (même mot fragmenté: 'I'+'m'+'mobilisations')
        - gap 1-8pt → ajouter espace (mots distincts)
        """
        if not label_words:
            return ""

        # Filtrer les préfixes parasites
        filtered = []
        for w in label_words:
            t = w['text']
            # Ignorer * et numéros de compte (211, 212, 311...)
            if t == '*' or re.match(r'^\d{3}$', t):
                continue
            # Ignorer les lettres isolées rotatives uniquement si très à gauche (x<50)
            if len(t) <= 2 and re.match(r'^[A-Z.]+$', t) and w['x0'] < 50:
                continue
            filtered.append(w)

        if not filtered:
            return ""

        # Détecter les blocs séparés par grand gap (>8pt)
        # Prendre le DERNIER bloc (le plus proche des colonnes numériques)
        blocks = [[filtered[0]]]
        for i in range(1, len(filtered)):
            gap = filtered[i]['x0'] - filtered[i-1]['x1']
            if gap > 8:
                blocks.append([])
            blocks[-1].append(filtered[i])

        last_block = next((b for b in reversed(blocks) if b), [])
        if not last_block:
            return ""

        # Reconstruction du texte du bloc
        # gap ≤ 0.2pt → vrai fragment du même mot → coller sans espace
        # gap >  0.2pt → mots séparés → ajouter espace
        # Exemples: DGI 'I'(0pt)'m'(0pt)'mobilisations' → 'Immobilisations'
        #           SAPST 'Frais'(0.95pt)'Préliminaires' → 'Frais Préliminaires'
        result = last_block[0]['text']
        for i in range(1, len(last_block)):
            gap = last_block[i]['x0'] - last_block[i-1]['x1']
            if gap <= 0.2:
                result += last_block[i]['text']    # coller: vrais fragments
            else:
                result += ' ' + last_block[i]['text']  # espace: mots distincts

        # Nettoyages finaux
        result = re.sub(r'\s+', ' ', result).strip()
        result = result.strip('[]()- \t*')
        result = re.sub(r'^[A-Z]\s*\.\s*[A-Z]?\s*\.?\s*', '', result)  # A.I., F., etc.
        result = re.sub(r'^\d{3}\s*', '', result)  # codes de compte résiduels

        return result.strip()

    # ── Parsing numérique universel ────────────────────────────────────────────

    @staticmethod
    def _is_num_token(t: str) -> bool:
        """Reconnaît tous les tokens numériques: standard, points-milliers SAPST, †."""
        t2 = t.replace('†', '').replace(' ', '')
        return bool(re.match(
            r'^-?\d{1,3}$'
            r'|^-?\d+[,\.]\d{2}$'
            r'|^-?\d+(\.\d+)+[,\.]\d{2}$'
            r'|^\.(\d+\.)*\d+[,\.]\d{2}$'
            r'|^-?0,00$', t2
        ))
    @staticmethod
    def _parse_num_str(s: str) -> float | None:
        """Parse un nombre depuis une chaîne reconstituée."""
        if not s:
            return None
        neg = s.startswith('-')
        s = s.lstrip('-').replace('†', '').replace(' ', '')

        # Format SAPST avec points comme milliers: 1.500.000.000,00 ou 7.092.363,69
        if re.match(r'^\d+(\.\d+)+[,\.]\d{2}$', s):
            # Supprimer tous les points (séparateurs milliers), garder la virgule décimale
            parts = re.split(r'[,\.]', s)
            integer_part = ''.join(parts[:-1])
            decimal_part = parts[-1]
            s = f"{integer_part}.{decimal_part}"
        elif re.match(r'^(\.\d+)+[,\.]\d{2}$', s):
            # Partiel SAPST: .500.000.000,00 (premier digit séparé)
            parts = re.split(r'[,\.]', s.lstrip('.'))
            integer_part = ''.join(p for p in parts[:-1] if p)
            decimal_part = parts[-1]
            s = f"{integer_part}.{decimal_part}" if integer_part else f"0.{decimal_part}"
        else:
            s = s.replace(',', '.')

        try:
            v = float(s)
            return -v if neg else v
        except ValueError:
            return None

    @staticmethod
    def _merge_and_parse_nums(words_with_bounds: list) -> list:
        """
        Fusionne les tokens numériques adjacents (gap < 6pt) et les parse.
        Gère tous les formats: espace-milliers, points-milliers (SAPST), †.
        Retourne [(x0, float), ...]
        """
        if not words_with_bounds:
            return []

        groups, grp = [], [words_with_bounds[0]]
        for i in range(1, len(words_with_bounds)):
            px0, px1, pt = grp[-1]
            cx0, cx1, ct = words_with_bounds[i]
            gap = cx0 - px1

            pt_clean = pt.replace('†', '')
            ct_clean = ct.replace('†', '')
            both_num = PDFParser._is_num_token(pt) and PDFParser._is_num_token(ct)

            # Fusionner si gap < 6pt ET les deux sont numériques
            # OU si c'est un format SAPST: digit seul suivi de '7.092.363,69'
            is_sapst = (re.match(r'^\d$', pt_clean) and
                        re.match(r'^\d+\.\d+[,\.]\d{2}$|^\.\d+[,\.]\d{2}$', ct_clean) and
                        gap < 8)

            if (both_num and gap < 6) or is_sapst:
                grp.append((cx0, cx1, ct))
            else:
                groups.append(grp)
                grp = [(cx0, cx1, ct)]
        groups.append(grp)

        result = []
        for g in groups:
            neg = any(t.replace('†', '').startswith('-') for _, _, t in g)
            raw = ''.join(t for _, _, t in g)
            if neg:
                raw = '-' + raw.lstrip('-')
            v = PDFParser._parse_num_str(raw)
            if v is not None:
                result.append((g[0][0], v))

        return result

    # ── Validité des labels ────────────────────────────────────────────────────

    def _is_valid_label(self, label: str, mode: str) -> bool:
        if not label or len(label) < 3:
            return False
        l = label.lower().strip()
        if l in SKIP_SINGLE:
            return False
        if SKIP_LABELS_RE.match(l):
            return False
        if mode == 'cpc' and RESULT_SKIP_RE.match(l):
            return False
        if re.match(r'^\d+$', label):
            return False
        return True

    # ── Enrichissement passif ──────────────────────────────────────────────────

    def _enrich_passif(self, data: dict):
        pv  = data["passif_values"]
        cpc = data["cpc_values"]

        # 1. Résultat net depuis CPC
        rn_key = "Résultat net de l'exercice"
        if rn_key not in pv:
            for k, v in cpc.items():
                if re.search(r'resultat net.*xi|resultat net.*xi-xii', k, re.I):
                    propre   = v[0] or 0
                    prec     = v[1] or 0
                    total_n  = round(propre + prec, 2)
                    total_n1 = v[2] if len(v) > 2 else None
                    pv[rn_key] = [total_n, total_n1]
                    logger.info(f"Résultat net inféré : N={total_n} N-1={total_n1}")
                    break

        # 2. Capital social: chercher sous libellés alternatifs
        capital_keys = ["Capital social ou personnel", "Capital social ou personnel (1)"]
        if not any(k in pv for k in capital_keys):
            for k, v in list(pv.items()):
                kl = k.lower()
                if any(x in kl for x in ["capital appelé", "dont versé", "capital appele"]):
                    if v and v[0] and v[0] > 0:
                        pv["Capital social ou personnel"] = v
                        logger.info(f"Capital social inféré : {v}")
                        break

    # ── Utilitaires ────────────────────────────────────────────────────────────

    def _page_text(self, idx: int) -> str:
        if idx >= self.n_pages:
            return ""
        return self.pages[idx].extract_text() or ""

    def __del__(self):
        try:
            self.pdf.close()
        except Exception:
            pass
