"""
core/table_parser.py — Parseur universel hybride
Méthode 1 : extract_tables() (pdfplumber) — quand le PDF a des bordures claires
Méthode 2 : X/Y (fallback) — quand les tableaux ne sont pas détectés

Structure des colonnes détectée automatiquement.
"""

import re
import pdfplumber
from collections import defaultdict
from utils.logger import get_logger

logger = get_logger(__name__)


# ── Utilitaires ───────────────────────────────────────────────────────────────

def parse_num(s) -> float | None:
    """Parse un nombre depuis une chaîne (formats FR et multi-points)."""
    if not s: return None
    s = str(s).strip().replace(' ', '').replace('\xa0', '').replace('†', '')
    if not s or s in ['-', '—', '']: return None
    neg = s.startswith('-')
    s = s.lstrip('-')
    # Multi-points SAPST : 1.500.000,00
    if re.match(r'^\d+(\.\d+)+,\d{2}$', s):
        parts = re.split(r'[,\.]', s)
        s = ''.join(parts[:-1]) + '.' + parts[-1]
    else:
        s = s.replace(',', '.')
    try:
        v = float(s)
        return -v if neg else v
    except ValueError:
        return None


def soft(s) -> str:
    """Normalisation légère du label : minuscules, retire *, (1), [A], romains."""
    if not s: return ''
    s = str(s).strip()
    s = re.sub(r'\n', ' ', s)
    s = re.sub(r'^\*\s*', '', s)                          # retirer *
    s = re.sub(r'\(\d+\)', '', s)                         # retirer (1)(2)
    s = re.sub(r'\s*\(\s*[A-Z]\s*\)\s*', ' ', s)         # retirer (A)(B)
    s = re.sub(r'\[\w\]', '', s)                          # retirer [A][B]
    s = re.sub(r'^(XVI|XV|XIV|XIII|XII|XI|X|IX|VIII|VII|VI|IV|III|II|I)\s+',
               '', s, flags=re.IGNORECASE)                # retirer numéros romains
    s = re.sub(r'\s+', ' ', s)
    return s.lower().strip()


def find_label(row) -> str:
    """Trouve le label dans une ligne (première cellule non-numérique de taille > 3)."""
    for v in row:
        if v and parse_num(str(v)) is None and len(str(v).strip()) > 3:
            return soft(str(v))
    return ''


def nums_by_col(row) -> dict:
    """Retourne {index_col: valeur} pour toutes les cellules numériques."""
    result = {}
    for i, v in enumerate(row):
        p = parse_num(str(v))
        if p is not None:
            result[i] = p
    return result


# ── Détection de la structure du tableau ──────────────────────────────────────

def detect_structure(table, section: str) -> dict:
    """
    Détecte les indices de colonnes pour chaque type de valeur.
    Retourne ex: {'label': 1, 'brut': 2, 'amort': 4, 'net_n1': 6}
    ou {'label': 1, 'val_n': 3, 'val_n1': 4}
    """
    if not table or not table[0]:
        return {}

    n_cols = len(table[0])

    # Trouver les lignes headers (les 6 premières)
    headers = []
    for row in table[:6]:
        headers.append([str(v).strip().upper() if v else '' for v in row])

    # Détecter col label (col avec le plus de texte non-numérique)
    label_col = 1  # par défaut col B

    if section == 'actif':
        # Chercher les mots-clés dans les headers
        brut_col = amort_col = net_col = net1_col = None
        for row in headers:
            for i, v in enumerate(row):
                if 'BRUT' in v: brut_col = i
                if 'AMORT' in v: amort_col = i
                if 'NET' in v and net_col is None: net_col = i
                if 'PRECEDENT' in v or 'PREC' in v: net1_col = i

        # Si non trouvés → deviner par position
        if brut_col is None:
            # Bilan2017: col2=brut, col4=amort, col5=net_n, col6=net_n1
            # BORJ:      col3=brut, col4=amort, col5=net_n, col6=net_n1
            # Chercher sur une ligne de données
            for row in table[6:]:
                nbc = nums_by_col(row)
                lbl = find_label(row)
                if lbl and len(nbc) >= 3:
                    cols = sorted(nbc.keys())
                    brut_col  = cols[0]
                    amort_col = cols[1] if len(cols) > 2 else None
                    net_col   = cols[-2] if len(cols) >= 2 else cols[-1]
                    net1_col  = cols[-1]
                    break

        return {
            'label': label_col,
            'brut':  brut_col,
            'amort': amort_col,
            'net_n': net_col,
            'net_n1': net1_col,
        }

    elif section == 'passif':
        # n_cols=5 → [sect, label, vide, N, N-1]
        # n_cols=4 → [sect, label, N, N-1]
        if n_cols >= 5:
            return {'label': 1, 'val_n': n_cols - 2, 'val_n1': n_cols - 1}
        else:
            return {'label': 1, 'val_n': n_cols - 2, 'val_n1': n_cols - 1}

    elif section == 'cpc':
        # Bilan2017 7 cols: [sect, sect, label, propre(3), prec(4), total_n(5), total_n1(6)]
        # BORJ 6 cols:      [sect, label, propre(2), prec(3), total_n(4), total_n1(5)]
        label_col = 2 if n_cols >= 7 else 1
        propre_col = 3 if n_cols >= 7 else 2
        prec_col   = 4 if n_cols >= 7 else 3
        total_n1   = n_cols - 1
        return {
            'label':   label_col,
            'propre':  propre_col,
            'prec':    prec_col,
            'total_n1': total_n1,
        }

    return {}


# ── Extraction depuis tableau ─────────────────────────────────────────────────

def extract_from_table(table, section: str) -> dict:
    """
    Extrait {label: [val1, val2, ...]} depuis un tableau pdfplumber.
    """
    if not table:
        return {}

    struct = detect_structure(table, section)
    if not struct:
        return {}

    result = {}
    label_col = struct.get('label', 1)

    # Lignes de données — le passif a 1 seule ligne de header
    skip_rows = 1 if section in ('passif', 'cpc') else 5
    for row in table[skip_rows:]:
        if len(row) <= label_col:
            continue

        label = soft(str(row[label_col])) if row[label_col] else ''

        # Chercher label dans d'autres colonnes si col 1 vide
        if not label:
            label = find_label(row)
        if not label or len(label) < 3:
            continue

        # Ignorer les totaux et sections
        if any(t in label for t in ['total i', 'total ii', 'total iii',
                                     'total général', 'total general',
                                     'total i+ii', 'a c t i f', 'p a s s i f']):
            continue

        if section == 'actif':
            brut  = parse_num(str(row[struct['brut']]))  if struct.get('brut')  is not None and struct['brut']  < len(row) else None
            amort = parse_num(str(row[struct['amort']])) if struct.get('amort') is not None and struct['amort'] < len(row) else None
            net1  = parse_num(str(row[struct['net_n1']])) if struct.get('net_n1') is not None and struct['net_n1'] < len(row) else None
            if any(v is not None for v in [brut, amort, net1]):
                result[label] = [brut, amort, net1]

        elif section == 'passif':
            vn  = parse_num(str(row[struct['val_n']]))  if struct.get('val_n')  is not None and struct['val_n']  < len(row) else None
            vn1 = parse_num(str(row[struct['val_n1']])) if struct.get('val_n1') is not None and struct['val_n1'] < len(row) else None
            if any(v is not None for v in [vn, vn1]):
                result[label] = [vn, vn1]

        elif section == 'cpc':
            propre = parse_num(str(row[struct['propre']]))   if struct.get('propre')   is not None and struct['propre']   < len(row) else None
            prec   = parse_num(str(row[struct['prec']]))     if struct.get('prec')     is not None and struct['prec']     < len(row) else None
            tot_n1 = parse_num(str(row[struct['total_n1']])) if struct.get('total_n1') is not None and struct['total_n1'] < len(row) else None
            if any(v is not None for v in [propre, prec, tot_n1]):
                result[label] = [propre, prec, tot_n1]

    logger.info(f"Table [{section}] : {len(result)} postes extraits")
    return result


# ── Fallback X/Y ──────────────────────────────────────────────────────────────

def is_num_tok(t: str) -> bool:
    t2 = t.replace('†', '').replace(' ', '')
    return bool(re.match(
        r'^-?\d{1,3}$|^-?\d+[,\.]\d{2}$|^-?\d+(\.\d+)+[,\.]\d{2}$'
        r'|^\.(\d+\.)*\d+[,\.]\d{2}$|^-?0,00$', t2
    ))


def extract_xy(page, section: str) -> dict:
    """Extraction X/Y (méthode v5) en fallback."""
    words = page.extract_words(x_tolerance=3, y_tolerance=3)
    if not words: return {}

    lines = defaultdict(list)
    for w in words:
        lines[round(w['top'] / 4) * 4].append(w)

    num_xs = [w['x0'] for w in words if is_num_tok(w['text']) and w['x0'] > 150]
    if not num_xs: return {}
    thresh = min(num_xs) - 10

    result = {}
    prev_label = None

    for y in sorted(lines.keys()):
        row = sorted(lines[y], key=lambda w: w['x0'])
        lw = [w for w in row if w['x0'] < thresh]
        nw = [w for w in row if w['x0'] >= thresh and is_num_tok(w['text'])]

        # Reconstruire label
        filtered = [w for w in lw if not (len(w['text']) <= 2
                    and re.match(r'^[A-Z.]+$', w['text']) and w['x0'] < 50)]
        if filtered:
            label_raw = filtered[0]['text']
            for i in range(1, len(filtered)):
                gap = filtered[i]['x0'] - filtered[i-1]['x1']
                label_raw += filtered[i]['text'] if gap <= 0.2 else ' ' + filtered[i]['text']
            label = soft(label_raw.strip())
        else:
            label = ''

        # Fusionner les tokens numériques
        vals = []
        if nw:
            grp = [nw[0]]
            for i in range(1, len(nw)):
                gap = nw[i]['x0'] - grp[-1]['x1']
                is_sapst = (len(grp[-1]['text']) <= 2 and grp[-1]['text'].isdigit()
                            and nw[i]['text'].startswith('.') and gap < 8)
                if gap < 5 or is_sapst:
                    grp.append(nw[i])
                else:
                    raw = ''.join(w['text'] for w in grp)
                    v = parse_num(raw)
                    if v is not None: vals.append(v)
                    grp = [nw[i]]
            raw = ''.join(w['text'] for w in grp)
            v = parse_num(raw)
            if v is not None: vals.append(v)

        if label and not vals:
            prev_label = label
        elif vals:
            effective = label if label else (prev_label or '')
            if not label: prev_label = None
            if effective and len(effective) > 3:
                if section == 'actif':
                    brut  = vals[0] if len(vals) > 0 else None
                    amort = vals[1] if len(vals) > 1 else None
                    net_n1 = vals[-1] if len(vals) >= 3 else None
                    result[effective] = [brut, amort, net_n1]
                elif section == 'passif':
                    result[effective] = [vals[0] if vals else None,
                                         vals[-1] if len(vals) > 1 else None]
                elif section == 'cpc':
                    result[effective] = [vals[0] if vals else None,
                                         vals[1] if len(vals) > 1 else None,
                                         vals[-1] if len(vals) >= 3 else None]

    logger.info(f"X/Y [{section}] : {len(result)} postes extraits")
    return result


# ── Parseur principal ─────────────────────────────────────────────────────────

class TableParser:
    """
    Parseur universel : essaie extract_tables() d'abord, puis X/Y en fallback.
    """

    def __init__(self, pdf_path: str):
        self.path = pdf_path
        self.pdf  = pdfplumber.open(pdf_path)
        self.n    = len(self.pdf.pages)
        logger.info(f"PDF chargé : {pdf_path} — {self.n} pages")

    def _is_table_usable(self, table, section: str, min_rows: int = 5) -> bool:
        """Vérifie qu'un tableau a assez de lignes avec données."""
        if not table: return False
        rows_with_data = sum(
            1 for row in table[4:]
            if sum(1 for v in row if v and parse_num(str(v)) is not None) >= 1
        )
        return rows_with_data >= min_rows

    def _extract_section(self, page_indices: list, section: str) -> dict:
        """
        Extrait une section depuis les pages indiquées.
        Essaie extract_tables() sur chaque page, sinon X/Y.
        """
        result = {}
        for idx in page_indices:
            if idx >= self.n:
                continue
            page = self.pdf.pages[idx]
            tables = page.extract_tables()
            used_table = False

            for table in tables:
                if self._is_table_usable(table, section):
                    extracted = extract_from_table(table, section)
                    if len(extracted) >= 3:
                        result.update(extracted)
                        used_table = True
                        logger.info(f"Page {idx+1} [{section}] → extract_tables() ✓")
                        break

            if not used_table:
                extracted = extract_xy(page, section)
                if extracted:
                    result.update(extracted)
                    logger.info(f"Page {idx+1} [{section}] → X/Y fallback")

        return result

    def _parse_info(self) -> dict:
        """Extrait les infos générales depuis la page 1."""
        import re as re2
        info = {}
        try:
            text = self.pdf.pages[0].extract_text() or ''
        except Exception:
            return {}

        for key, pat in [
            ('raison_sociale',      r'Raison\s+[Ss]ociale\s*[:\-]?\s*([A-Z][^\n]{3,60})'),
            ('identifiant_fiscal',  r'[Ii]dentifiant\s+[Ff]iscal\s*[:\-]?\s*(\d+)'),
            ('taxe_professionnelle',r'(?:Taxe|Art\.)\s+[Pp]rof\w*\.?\s*[:\-]?\s*([\d\s]+)'),
            ('adresse',             r'[Aa]dresse\s*[:\-]?\s*([^\n]{5,60})'),
        ]:
            m = re2.search(pat, text, re2.IGNORECASE)
            if m:
                info[key] = m.group(1).strip()

        m = re2.search(r'(\d{2}/\d{2}/\d{4})\s+au\s+(\d{2}/\d{2}/\d{4})', text)
        if m:
            info['exercice']       = f"Du {m.group(1)} au {m.group(2)}"
            info['exercice_debut'] = m.group(1)
            info['exercice_fin']   = m.group(2)

        for k in ['raison_sociale','identifiant_fiscal','exercice','exercice_fin']:
            info.setdefault(k, '')
        info['pages'] = self.n
        return info

    def parse(self) -> dict:
        """Parse tout le PDF et retourne les données structurées."""
        # Détecter format DGI (7 pages) ou AMMC (5 pages)
        is_dgi = self.n == 7

        if is_dgi:
            pages_actif  = [1, 2]
            pages_passif = [3]
            pages_cpc    = [4, 5, 6]
        else:
            pages_actif  = [1, 2]
            pages_passif = [2, 3]
            pages_cpc    = [3, 4, 5]

        actif  = self._extract_section(pages_actif,  'actif')
        passif = self._extract_section(pages_passif, 'passif')
        cpc    = self._extract_section(pages_cpc,    'cpc')

        # Inférer résultat net depuis CPC si manquant dans passif
        rn_key = "résultat net de l'exercice"
        if rn_key not in passif:
            for k, v in cpc.items():
                if 'résultat net' in k and 'xi' in k.lower() or k == 'résultat net':
                    propre = v[0] or 0
                    prec   = v[1] or 0
                    total  = round(propre + prec, 2)
                    passif[rn_key] = [total, v[2] if len(v) > 2 else None]
                    logger.info(f"Résultat net inféré : {total}")
                    break

        self.pdf.close()
        return {
            'info':          self._parse_info(),
            'actif_values':  actif,
            'passif_values': passif,
            'cpc_values':    cpc,
        }
