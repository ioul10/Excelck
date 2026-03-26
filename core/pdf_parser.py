"""
core/pdf_parser.py  v4 — Parseur universel par coordonnées X/Y

Algorithme :
  1. extract_words() → liste de mots avec position (x0, x1, top)
  2. Regrouper par ligne Y (tolérance 4pt)
  3. Détecter automatiquement le seuil label/nombre
  4. Fusionner les tokens numériques adjacents (gap < 6pt)
  5. Assigner les colonnes par ordre X

Supporte n'importe quel PDF MCN (Modèle Comptable Normal) :
  - Mise en page dense ou espacée
  - Grandes ou petites valeurs
  - Tout générateur de PDF (comptabilité, ERP...)
"""

import re
from collections import defaultdict
import pdfplumber
from utils.logger import get_logger

logger = get_logger(__name__)

# Labels à ignorer (en-têtes, métadonnées, totaux calculés)
SKIP_RE = re.compile(r'^(tableau|bilan\s*\(|compte de produits|b\s+i\s+l\s+a\s+n|'
                     r'modèle|exercice du|identifiant|raison sociale|'
                     r'1\)variation|2\)achats|nb\s*:|cadre réservé|'
                     r'numéro d.enregistrement|signature|'
                     r'\(1\)capital|\(2\)bénéf|pages|total général i.*ii.*iii)', re.I)

SKIP_SINGLE = {'brut', 'net', 'amortissements', 'provisions', 'exercice',
               'precedent', 'précédent', 'designation', 'désignation',
               'operations', 'opérations', 'propres', 'concernant',
               'totaux', 'l\'exercice', 'de', 'du', 'et', 'en'}

RESULT_SKIP = re.compile(r'^(résultat d.exploitation|résultat financier|résultat courant|'
                         r'résultat non courant|résultat avant impôts|'
                         r'resultat d.exploitation|resultat financier|resultat courant|'
                         r'resultat non courant|resultat avant impots|'
                         r'total i |total ii|total iii|total des |total général|'
                         r'produits d.exploitation$|charges d.exploitation$|'
                         r'charges financières$|produits financiers$|'
                         r'produits non courants$|charges non courants$)', re.I)


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
            ("raison_sociale",      r"Raison sociale\s*[:\-]?\s*(.+)"),
            ("taxe_professionnelle", r"Taxe\s+[Pp]rofessionnelle\s*[:\-]?\s*([\d\s]+)"),
            ("identifiant_fiscal",   r"[Ii]dentifiant\s+[Ff]iscal\s*[:\-]?\s*(\d+)"),
            ("adresse",              r"Adresse\s*[:\-]?\s*(.+)"),
        ]:
            m = re.search(pat, text0, re.IGNORECASE)
            if m:
                info[key] = m.group(1).strip()

        # Fallback raison sociale depuis pages 1-2
        if not info.get("raison_sociale"):
            for i in range(1, min(3, self.n_pages)):
                t = self._page_text(i)
                m = re.search(r"((?:SOCIETE|SOCIÉTÉ|AGENCE|OFFICE|DIRECTION|S\.A\.)[^\n]{3,80})", t)
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
            if m:
                info["exercice"]       = f"Du {m.group(1)} au {m.group(2)}"
                info["exercice_debut"] = m.group(1)
                info["exercice_fin"]   = m.group(2)
                break

        for k in ["exercice", "exercice_fin", "exercice_debut", "date_declaration",
                  "raison_sociale", "identifiant_fiscal", "adresse", "taxe_professionnelle"]:
            info.setdefault(k, "")

        # Date déclaration
        if not info.get("date_declaration"):
            for i in range(min(2, self.n_pages)):
                t = self._page_text(i)
                m = re.search(r"[Ll]e\s+(\d{2}/\d{2}/\d{4})", t)
                if m:
                    info["date_declaration"] = m.group(1)
                    break

        info["pages"] = self.n_pages
        return info

    # ── Extraction universelle ─────────────────────────────────────────────────

    def _extract_section(self, pages: list, mode: str) -> dict:
        """
        Extrait les postes d'une ou plusieurs pages.
        mode = 'actif' | 'passif' | 'cpc'
        Retourne {label: [val1, val2, ...]} avec len(vals) = nb colonnes du tableau.
        """
        values = {}

        for page_idx in pages:
            if page_idx >= self.n_pages:
                continue

            # Vérifier que la page correspond bien à la section attendue
            text = self._page_text(page_idx)
            if not self._page_matches_mode(text, mode):
                continue

            page_values = self._extract_page(page_idx, mode)

            for label, vals in page_values.items():
                if label not in values:
                    values[label] = vals
                else:
                    # Enrichir avec les valeurs manquantes
                    existing = values[label]
                    merged = []
                    max_len = max(len(existing), len(vals))
                    for i in range(max_len):
                        ev = existing[i] if i < len(existing) else None
                        nv = vals[i]     if i < len(vals)     else None
                        merged.append(ev if ev is not None else nv)
                    values[label] = merged

        logger.info(f"{mode}: {len(values)} postes extraits")
        return values

    def _page_matches_mode(self, text: str, mode: str) -> bool:
        t = text.lower()
        if mode == 'actif':
            return any(k in t for k in ['actif', 'immobilisations', 'stocks', 'créances'])
        if mode == 'passif':
            return any(k in t for k in ['passif', 'capitaux propres', 'capital', 'dettes'])
        if mode == 'cpc':
            return any(k in t for k in ['produits', 'charges', 'exploitation', 'résultat'])
        return True

    def _extract_page(self, page_idx: int, mode: str) -> dict:
        """
        Extraction par coordonnées X/Y sur une page.
        """
        page  = self.pages[page_idx]
        words = page.extract_words(x_tolerance=3, y_tolerance=3)
        if not words:
            return {}

        # Regrouper par ligne Y (tolérance 6pt pour couvrir les PDFs
        # où label et valeurs sont sur des Y légèrement différents)
        lines = defaultdict(list)
        for w in words:
            lines[round(w['top'] / 6) * 6].append(w)

        # Détecter le seuil label/nombre
        num_xs = [w['x0'] for w in words
                  if self._is_num_token(w['text']) and w['x0'] > 200]
        if not num_xs:
            return {}

        label_thresh = min(num_xs) - 10

        # Détecter les positions X des colonnes (clustering)
        col_xs = self._detect_columns(num_xs)

        result = {}
        for y in sorted(lines.keys()):
            row = sorted(lines[y], key=lambda w: w['x0'])

            # Labels : tokens à gauche du seuil, sans lettres isolées rotatives
            label_words = [w['text'] for w in row
                          if w['x0'] < label_thresh
                          and not re.match(r'^[A-Z\s]{1,2}$', w['text'])]
            if not label_words:
                continue

            label = ' '.join(label_words).strip()
            label = re.sub(r'\s+', ' ', label)

            if not self._is_valid_label(label, mode):
                continue

            # Tokens numériques à droite
            num_words = [(w['x0'], w['x1'], w['text']) for w in row
                        if w['x0'] >= label_thresh
                        and self._is_num_token(w['text'])]
            if not num_words:
                continue

            # Fusionner les tokens du même nombre
            merged = self._merge_nums(num_words)
            if not merged:
                continue

            # Assigner aux colonnes
            vals = self._assign_to_cols(merged, col_xs, mode)

            if any(v is not None for v in vals):
                result[label] = vals

        return result

    def _detect_columns(self, num_xs: list) -> list:
        """
        Détecte les positions X des colonnes numériques par clustering.
        Retourne une liste triée de positions X représentatives.
        """
        if not num_xs:
            return []

        sorted_xs = sorted(set(round(x / 5) * 5 for x in num_xs))
        clusters = []
        grp = [sorted_xs[0]]

        for x in sorted_xs[1:]:
            if x - grp[-1] < 30:
                grp.append(x)
            else:
                clusters.append(sum(grp) / len(grp))
                grp = [x]
        clusters.append(sum(grp) / len(grp))

        return sorted(clusters)

    def _assign_to_cols(self, merged: list, col_xs: list, mode: str) -> list:
        """
        Assigne les valeurs numériques aux colonnes attendues selon le mode.

        Pour l'actif  : [brut, amort, net_n, net_n1]   → on garde brut, amort, net_n1
        Pour le passif: [val_n, val_n1]
        Pour le CPC   : [propre_n, prec_n, total_n, total_n1] → on garde propre_n, prec_n, total_n1
        """
        if not col_xs or not merged:
            return [v for _, v in merged]

        # Créer un dict col_position → valeur
        col_map = {}
        for x, v in merged:
            # Trouver la colonne la plus proche
            closest = min(col_xs, key=lambda c: abs(c - x))
            col_idx = col_xs.index(closest)
            col_map[col_idx] = v

        # Nombre de colonnes attendues
        n_cols = len(col_xs)

        if mode == 'actif':
            # 4 colonnes : brut(0), amort(1), net_n(2), net_n1(3)
            # On retourne [brut, amort, net_n1]
            if n_cols >= 4:
                return [col_map.get(0), col_map.get(1), col_map.get(3)]
            elif n_cols == 3:
                return [col_map.get(0), col_map.get(1), col_map.get(2)]
            else:
                return [v for _, v in merged]

        elif mode == 'passif':
            # 2 colonnes : val_n(0), val_n1(1)
            if n_cols >= 2:
                return [col_map.get(0), col_map.get(1)]
            else:
                return [v for _, v in merged]

        elif mode == 'cpc':
            # Idéalement 4 colonnes : propre_n(0), prec_n(1), total_n(2), total_n1(3)
            # Mais certains PDFs n'ont que 2 colonnes numériques (total_n et total_n1)
            # Détecter : si les valeurs sont toutes dans les 2 dernières colonnes
            # → les mapper comme propre_n=col_max-1, total_n1=col_max
            filled = {k: v for k, v in col_map.items() if v is not None}
            if n_cols >= 4:
                # Vérifier si les 2 premières colonnes sont vides
                if not filled.get(0) and not filled.get(1) and (filled.get(2) or filled.get(3)):
                    # Seulement total_n et total_n1 → propre_n = total_n, total_n1 = total_n1
                    return [filled.get(2), None, filled.get(3)]
                return [col_map.get(0), col_map.get(1), col_map.get(3)]
            elif n_cols >= 3:
                return [col_map.get(0), col_map.get(1), col_map.get(2)]
            elif n_cols == 2:
                # 2 colonnes = propre_n et total_n1
                return [col_map.get(0), None, col_map.get(1)]
            else:
                return [v for _, v in merged]

        return [v for _, v in merged]

    # ── Enrichissement passif ──────────────────────────────────────────────────

    def _enrich_passif(self, data: dict):
        """Enrichit passif_values avec des valeurs inférées."""
        pv  = data["passif_values"]
        cpc = data["cpc_values"]

        # 1. Résultat net : depuis RESULTAT NET (XI-XII) du CPC
        rn_key = "Résultat net de l'exercice"
        if rn_key not in pv:
            for k, v in cpc.items():
                if re.search(r'resultat net.*xi|resultat net.*xi-xii', k, re.I):
                    propre = v[0] or 0
                    prec   = v[1] or 0
                    total_n  = round(propre + prec, 2)
                    total_n1 = v[2] if len(v) > 2 else None
                    pv[rn_key] = [total_n, total_n1]
                    logger.info(f"Résultat net inféré : N={total_n} N-1={total_n1}")
                    break

        # 2. Capital social : si absent, le chercher sous d'autres libellés courants
        capital_keys = ["Capital social ou personnel", "Capital social ou personnel (1)"]
        has_capital = any(k in pv for k in capital_keys)
        if not has_capital:
            # Chercher "capital appelé" ou "dont versé" comme proxy
            for k, v in pv.items():
                kl = k.lower()
                if ("capital appelé" in kl or "dont versé" in kl or "capital appele" in kl):
                    if v and v[0] and v[0] > 0:
                        pv["Capital social ou personnel"] = v
                        logger.info(f"Capital social inféré depuis '{k}': {v}")
                        break

    # ── Utilitaires ────────────────────────────────────────────────────────────

    @staticmethod
    def _is_num_token(t: str) -> bool:
        return bool(re.match(r'^-?\d{1,3}$|^-?\d+[,\.]\d{2}$|^0,00$', t))

    @staticmethod
    def _merge_nums(words_with_bounds: list) -> list:
        """
        Fusionne les tokens numériques adjacents (gap < 6pt = même nombre).
        [(x0, x1, text), ...] → [(x0, float), ...]
        """
        if not words_with_bounds:
            return []

        groups, grp = [], [words_with_bounds[0]]
        for i in range(1, len(words_with_bounds)):
            px0, px1, pt = grp[-1]
            cx0, cx1, ct = words_with_bounds[i]
            if PDFParser._is_num_token(pt) and PDFParser._is_num_token(ct) and cx0 - px1 < 6:
                grp.append((cx0, cx1, ct))
            else:
                groups.append(grp)
                grp = [(cx0, cx1, ct)]
        groups.append(grp)

        result = []
        for g in groups:
            neg = any(t.startswith('-') for _, _, t in g)
            s = ''.join(t for _, _, t in g).lstrip('-').replace(',', '.').replace(' ', '')
            if re.match(r'^\d+\.?\d*$', s):
                try:
                    result.append((g[0][0], (-1 if neg else 1) * float(s)))
                except ValueError:
                    pass
        return result

    def _is_valid_label(self, label: str, mode: str) -> bool:
        """Filtre les lignes qui ne sont pas des postes à injecter."""
        if len(label) < 3:
            return False
        l = label.lower().strip()

        if l in SKIP_SINGLE:
            return False
        if SKIP_RE.match(l):
            return False

        # Pour le CPC : bloquer les résultats calculés par formule
        if mode == 'cpc' and RESULT_SKIP.match(l):
            return False

        # Labels purement numériques
        if re.match(r'^\d+$', label):
            return False

        return True

    def _page_text(self, idx: int) -> str:
        if idx >= self.n_pages:
            return ""
        return self.pages[idx].extract_text() or ""

    def __del__(self):
        try:
            self.pdf.close()
        except Exception:
            pass
