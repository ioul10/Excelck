"""
core/direct_injector.py — Injection directe par label exact

Approche : 
  1. Construire un index Excel : soft_label → (feuille, row)
  2. Pour chaque ligne du PDF : soft_label → chercher dans l'index → injecter
  3. Pas de fuzzy matching agressif — uniquement correspondance exacte + aliases

Avantages :
  - "Autres immobilisations incorporelles" ne confond plus avec "corporelles"
  - "Comptes de régularisation- Actif" matche malgré l'espace autour du tiret
  - Pas de dérive par fuzzy matching
"""

import re, pdfplumber
from collections import defaultdict
from utils.logger import get_logger

logger = get_logger(__name__)

# ── Alias PDF → label normalisé Excel ─────────────────────────────────────────
# Quand le PDF utilise une abréviation ou variante du label Excel
ALIASES = {
    # Actif
    'immobilisations en recherche et dev.':     'immobilisations en recherche et développement',
    'immobilisations en recherche et dev':      'immobilisations en recherche et développement',
    'mobilier, mat. de bureau, aménagement divers': 'mobilier, mat. de bureau, aménagements divers',
    'mobilier matériel de bureau et aménagement divers': 'mobilier, mat. de bureau, aménagements divers',
    'mobilier, matériel de bureau et aménagement divers': 'mobilier, mat. de bureau, aménagements divers',
    'fournis. débiteurs, avances et acomptes':  'fournisseurs débiteurs, avances et acomptes',
    'fournisseurs débiteurs, avances et acomptes': 'fournisseurs débiteurs, avances et acomptes',
    'comptes de régularisation-actif':          'comptes de régularisation - actif',
    'comptes de régularisation- actif':         'comptes de régularisation - actif',
    # Passif
    'capital social ou personnel':              'capital appelé',  # → injecté en R7
    'subvention d\'investissement':             'subventions d\'investissement',
    'comptes de régularisation passif':         'comptes de régularisation passif',
    # Etat : résolu par contexte section
}

# Labels qui sont des titres de section → ne pas injecter
SECTION_TITLES = {
    'immobilisations en non-valeurs',
    'immobilisations incorporelles',
    'immobilisations corporelles',
    'immobilisations financières',
    'écarts de conversion actif',
    'total i',
    'stocks',
    'créances de l\'actif circulant',
    'titres et valeurs de placement',
    'écarts de conversion actif - éléments circulants',
    'total ii',
    'trésorerie-actif',
    'trésorerie actif',
    'total iii',
    'total général',
    'total général i+ii+iii',
    'capitaux propres',
    'capitaux propres assimilés',
    'dettes de financement',
    'provisions durables pour risques et charges',
    'écarts de conversion - passif',
    'total i passif',
    'dettes du passif circulant',
    'trésorerie passif',
    'total général passif',
}


def soft_normalize(s: str) -> str:
    """
    Normalisation légère — garde les accents, uniformise espaces et tirets.
    Retire les suffixes [A], [B], (1), (2), etc.
    """
    s = str(s).strip()
    # Retirer suffixes de section : [A], [B], (1)...
    s = re.sub(r'\s*[\[\(]\w[\]\)]\s*$', '', s)
    # Uniformiser tirets
    s = re.sub(r'\s*[-–—]\s*', ' - ', s)
    # Espaces multiples
    s = re.sub(r'\s+', ' ', s)
    return s.lower().strip()


def build_excel_index(wb) -> dict:
    """
    Construit l'index : soft_label → (sheet_name, row)
    Uniquement les cellules VIDES (pas de formule).
    """
    index = {}

    for sheet_name in ['2 - Bilan Actif', '3 - Bilan Passif', '4 - CPC']:
        if sheet_name not in wb.sheetnames:
            continue
        ws = wb[sheet_name]

        for row in ws.iter_rows(min_row=5, max_row=ws.max_row):
            lbl = row[1].value   # col B = label
            if not lbl:
                continue
            lbl_str = str(lbl)
            if lbl_str.startswith(('▶', ' ▶', '=')):
                continue

            # Vérifier que la cellule valeur (col C) est vide ou non-formule
            c_val = row[2].value
            if isinstance(c_val, str) and c_val.startswith('='):
                continue  # Cellule avec formule → calculée auto

            n = soft_normalize(lbl_str)
            if n and n not in SECTION_TITLES:
                index[n] = (sheet_name, row[0].row)

    logger.info(f"Index Excel : {len(index)} cellules indexées")
    return index


def _is_num_token(t: str) -> bool:
    """Reconnaît les tokens numériques."""
    t2 = t.replace('†', '').replace(' ', '')
    return bool(re.match(
        r'^-?\d{1,3}$'
        r'|^-?\d+[,\.]\d{2}$'
        r'|^-?\d+(\.\d+)+[,\.]\d{2}$'
        r'|^\.(\d+\.)*\d+[,\.]\d{2}$'
        r'|^-?0,00$', t2
    ))


def _parse_num_str(s: str):
    """Parse un nombre depuis une chaîne."""
    if not s: return None
    neg = s.startswith('-')
    s = s.lstrip('-').replace('†', '').replace(' ', '')
    if re.match(r'^\d+(\.\d+)+[,\.]\d{2}$', s):
        parts = re.split(r'[,\.]', s)
        s = ''.join(parts[:-1]) + '.' + parts[-1]
    elif re.match(r'^(\.\d+)+[,\.]\d{2}$', s):
        parts = re.split(r'[,\.]', s.lstrip('.'))
        s = ''.join(p for p in parts[:-1] if p) + '.' + parts[-1]
    else:
        s = s.replace(',', '.')
    try:
        v = float(s)
        return -v if neg else v
    except ValueError:
        return None


def extract_page_rows(page) -> list:
    """
    Extrait toutes les lignes d'une page avec leurs labels et valeurs.
    Retourne : [(label, [val1, val2, ...]), ...]
    """
    words = page.extract_words(x_tolerance=3, y_tolerance=3)
    if not words:
        return []

    # Grouper par Y
    lines = defaultdict(list)
    for w in words:
        lines[round(w['top'] / 4) * 4].append(w)

    # Trouver le seuil label/valeurs
    num_xs = [w['x0'] for w in words if _is_num_token(w['text']) and w['x0'] > 150]
    if not num_xs:
        return []
    thresh = min(num_xs) - 10

    result = []
    for y in sorted(lines.keys()):
        row = sorted(lines[y], key=lambda w: w['x0'])
        lw = [w for w in row if w['x0'] < thresh]
        nw = [w for w in row if w['x0'] >= thresh and _is_num_token(w['text'])]

        if not lw and not nw:
            continue

        # Reconstruire label (gap ≤ 0.2pt → coller, sinon espace)
        # Filtrer les lettres rotatives (x0 < 50, longueur ≤ 2)
        filtered = [w for w in lw
                    if not (len(w['text']) <= 2
                            and re.match(r'^[A-Z.]+$', w['text'])
                            and w['x0'] < 50)]

        if filtered:
            label = filtered[0]['text']
            for i in range(1, len(filtered)):
                gap = filtered[i]['x0'] - filtered[i-1]['x1']
                label += filtered[i]['text'] if gap <= 0.2 else ' ' + filtered[i]['text']
            label = re.sub(r'\s+', ' ', label).strip()
        else:
            label = ''

        # Fusionner les tokens numériques adjacents
        vals = []
        if nw:
            grp = [nw[0]]
            for i in range(1, len(nw)):
                gap = nw[i]['x0'] - grp[-1]['x1']
                # Fusion SAPST : '1' + '.500.000,00'
                is_sapst = (len(grp[-1]['text']) <= 2
                            and grp[-1]['text'].isdigit()
                            and (nw[i]['text'].startswith('.')
                                 or '.' in nw[i]['text'])
                            and gap < 8)
                if gap < 5 or is_sapst:
                    grp.append(nw[i])
                else:
                    raw = ''.join(w['text'] for w in grp)
                    v = _parse_num_str(raw)
                    if v is not None:
                        vals.append((grp[0]['x0'], v))
                    grp = [nw[i]]
            raw = ''.join(w['text'] for w in grp)
            v = _parse_num_str(raw)
            if v is not None:
                vals.append((grp[0]['x0'], v))

        if label or vals:
            result.append((label, vals))

    return result


def detect_columns(page_rows: list, section: str) -> list:
    """
    Détecte les positions X des colonnes et assigne les valeurs.
    Retourne [(label, [brut_ou_n, amort_ou_n1, net_n1_ou_None]), ...]
    """
    # Collecter toutes les positions X de valeurs
    all_xs = []
    for _, vals in page_rows:
        for x, _ in vals:
            all_xs.append(x)

    if not all_xs:
        return [(lbl, [v for _, v in vals]) for lbl, vals in page_rows]

    # Grouper les X en colonnes (gap > 15pt → nouvelle colonne)
    all_xs_sorted = sorted(set(round(x / 5) * 5 for x in all_xs))
    col_centers = [all_xs_sorted[0]]
    for x in all_xs_sorted[1:]:
        if x - col_centers[-1] > 15:
            col_centers.append(x)
        else:
            col_centers[-1] = (col_centers[-1] + x) // 2

    def assign_col(x):
        return min(range(len(col_centers)),
                   key=lambda i: abs(col_centers[i] - x))

    result = []
    n_cols = len(col_centers)

    for label, vals in page_rows:
        if not vals:
            result.append((label, []))
            continue

        # Assigner chaque valeur à sa colonne
        assigned = {}
        for x, v in vals:
            c = assign_col(x)
            assigned[c] = v

        if section == 'actif':
            # Colonnes : brut, amort, net_n, net_n1
            brut  = assigned.get(0)
            amort = assigned.get(1)
            net_n1 = assigned.get(n_cols - 1) if n_cols >= 3 else None
            result.append((label, [brut, amort, net_n1]))

        elif section in ('passif', 'cpc'):
            # Colonnes : val_n, val_n1
            val_n  = assigned.get(0)
            val_n1 = assigned.get(n_cols - 1) if n_cols >= 2 else None
            result.append((label, [val_n, val_n1]))

        else:
            result.append((label, [v for _, v in vals]))

    return result


class DirectInjector:
    """
    Injecteur direct : label PDF → cellule Excel par correspondance exacte.
    """

    def __init__(self, wb):
        self.wb    = wb
        self.index = build_excel_index(wb)
        # Index inversé pour debug
        self._injected = []
        self._skipped  = []

    def _find_cell(self, label: str, context: str = '') -> tuple:
        """
        Trouve la cellule Excel pour un label PDF.
        context : 'actif' | 'passif' | 'cpc' pour lever l'ambiguïté
        """
        n = soft_normalize(label)

        # Appliquer les alias
        n = ALIASES.get(n, n)

        # Cherche directe
        found = self.index.get(n)
        if found:
            # Vérifier que le contexte correspond
            if context and context not in found[0].lower():
                # "Etat" existe en actif ET passif → choisir selon contexte
                pass
            return found

        # Cherche avec suppression des accents si pas trouvé
        import unicodedata
        def no_accent(s):
            return ''.join(c for c in unicodedata.normalize('NFD', s)
                           if unicodedata.category(c) != 'Mn')

        n_na = no_accent(n)
        for key, val in self.index.items():
            if no_accent(key) == n_na:
                if not context or context in val[0].lower():
                    return val

        return None

    def inject_page(self, page_rows: list, section: str) -> int:
        """
        Injecte les valeurs d'une page dans l'Excel.
        Retourne le nombre de valeurs injectées.
        """
        rows_with_cols = detect_columns(page_rows, section)
        ws_actif  = self.wb['2 - Bilan Actif']
        ws_passif = self.wb['3 - Bilan Passif']
        ws_cpc    = self.wb['4 - CPC'] if '4 - CPC' in self.wb.sheetnames else None
        count = 0

        for label, vals in rows_with_cols:
            if not label or not vals:
                continue

            # Ignorer les titres de sections
            n = soft_normalize(label)
            if n in SECTION_TITLES:
                continue
            if any(t in n for t in ['total', 'sous-total']):
                continue

            cell_info = self._find_cell(label, section)
            if not cell_info:
                self._skipped.append(label)
                continue

            sheet_name, row = cell_info
            ws = (ws_actif if 'Actif' in sheet_name
                  else ws_passif if 'Passif' in sheet_name
                  else ws_cpc)
            if ws is None:
                continue

            if section == 'actif':
                brut, amort, net_n1 = (vals + [None, None, None])[:3]
                if brut is not None:
                    ws.cell(row, 3).value = brut   # col C = Brut
                    count += 1
                if amort is not None:
                    ws.cell(row, 4).value = amort  # col D = Amort
                    count += 1
                if net_n1 is not None:
                    ws.cell(row, 6).value = net_n1 # col F = Net N-1
                    count += 1
            else:  # passif ou cpc
                val_n, val_n1 = (vals + [None, None])[:2]
                if val_n is not None:
                    ws.cell(row, 3).value = val_n   # col C = N
                    count += 1
                if val_n1 is not None:
                    ws.cell(row, 4).value = val_n1  # col D = N-1
                    count += 1

            self._injected.append((label, sheet_name, row))

        return count

    def inject_pdf(self, pdf_path: str) -> dict:
        """
        Injecte tout un PDF dans l'Excel.
        """
        import pdfplumber as plumber
        pdf = plumber.open(pdf_path)
        n   = len(pdf.pages)
        total = 0

        # Détecter le format
        is_dgi = n == 7

        if is_dgi:
            pages_actif  = [1, 2]   # pages 2-3
            pages_passif = [3]       # page 4
            pages_cpc    = [4, 5, 6] # pages 5-7
        else:
            pages_actif  = [1, 2]   # pages 2-3 (AMMC peut avoir actif sur 2 pages aussi)
            pages_passif = [2]       # page 3
            pages_cpc    = [3, 4]    # pages 4-5

        for pg_idx in pages_actif:
            if pg_idx < n:
                rows = extract_page_rows(pdf.pages[pg_idx])
                total += self.inject_page(rows, 'actif')

        for pg_idx in pages_passif:
            if pg_idx < n:
                rows = extract_page_rows(pdf.pages[pg_idx])
                total += self.inject_page(rows, 'passif')

        for pg_idx in pages_cpc:
            if pg_idx < n:
                rows = extract_page_rows(pdf.pages[pg_idx])
                total += self.inject_page(rows, 'cpc')

        pdf.close()
        logger.info(f"DirectInjector : {total} valeurs injectées | {len(self._skipped)} skippées")
        return {
            'injected': total,
            'skipped':  list(set(self._skipped)),
        }
