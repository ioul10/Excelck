"""
core/injector_v7.py — Injecteur v7
Construit l'index depuis EX_template.xlsx et injecte les valeurs extraites.
Correspondance par label exact (soft) avec fallback sans accents.
"""

import unicodedata, re, openpyxl
from utils.logger import get_logger

logger = get_logger(__name__)


def soft(s: str) -> str:
    """Normalisation légère : minuscules, sans *, (1), [A], romains."""
    if not s: return ''
    s = str(s).strip()
    s = re.sub(r'\n', ' ', s)
    s = re.sub(r'^\*\s*', '', s)
    s = re.sub(r'\(\d+\)', '', s)
    s = re.sub(r'\s*\(\s*[A-Z]\s*\)\s*', ' ', s)
    s = re.sub(r'\[\w\]', '', s)
    s = re.sub(r'^(XVI|XV|XIV|XIII|XII|XI|X|IX|VIII|VII|VI|IV|III|II|I)\s+',
               '', s, flags=re.IGNORECASE)
    s = re.sub(r'\s+', ' ', s)
    return s.lower().strip()


def no_accent(s: str) -> str:
    return ''.join(c for c in unicodedata.normalize('NFD', s)
                   if unicodedata.category(c) != 'Mn')


# Aliases : label PDF → label normalisé Excel
ALIASES = {
    # Passif
    'capital social ou personnel':                    'capital appelé',
    'subvention d\'investissement':                   'subventions d\'investissement',
    'moins : capital appelé':                         'capital appelé',
    'capital appelé':                                 'capital appelé',
    # CPC variantes
    'achats consommés de matières et':                'achats consommés de matières et fournitures',
    'achats consommés de matières et fournitures':    'achats consommés de matières et fournitures',
    'achats revendus de marchandises':                'achats revendus de marchandises',
    'charges d\'intérêts':                            'charges d\'intérêts',
    'charges d\'interêts':                            'charges d\'intérêts',
    'impots sur les benefices':                       'impôts sur les bénéfices',
    'reprises financières : transferts de charges':   'reprises financières : transferts de charges',
    'autres produits non courants':                   'autres produits non courants',
    # Actif
    'immobilisations en recherche et dev.':           'immobilisations en recherche et développement',
    'mobilier, mat. de bureau, aménagement divers':   'mobilier, mat. de bureau, aménagements divers',
    'mobilier matériel de bureau et aménagement divers': 'mobilier, mat. de bureau, aménagements divers',
    'fournis. débiteurs, avances et acomptes':        'fournisseurs débiteurs, avances et acomptes',
}

# Labels à ignorer (totaux calculés par formule, sections)
SKIP = {
    'immobilisations en non-valeurs', 'immobilisations incorporelles',
    'immobilisations corporelles', 'immobilisations financières',
    'écarts de conversion actif', 'total i', 'total ii', 'total iii',
    'total général', 'total général i+ii+iii', 'total a+b+c+d+e',
    'total (a+b+c+d+e)', 'stocks', 'créances de l\'actif circulant',
    'titres et valeurs de placement', 'trésorerie-actif', 'trésorerie actif',
    'capitaux propres', 'total des capitaux propres',
    'capitaux propres assimilés', 'dettes de financement',
    'provisions durables pour risques et charges',
    'écarts de conversion - passif', 'total i passif',
    'dettes du passif circulant', 'trésorerie passif',
    'total général passif', 'chiffres d\'affaires',
    'produits d\'exploitation', 'charges d\'exploitation',
    'produits financiers', 'charges financières', 'charges financieres',
    'produits non courants', 'charges non courants', 'charges non-courants',
    'résultat d\'exploitation', 'résultat financier', 'résultat courant',
    'résultat non courant', 'résultat avant impôts', 'résultat net',
    'total des produits', 'total des charges',
    'n a n c i e r', 'e x p l o i t a t i o n',
    '3 = 2 + 1', 'iii', 'vii', 'vi', 'ix', 'x', 'xi', 'xii',
}


def build_index(wb) -> dict:
    """
    Construit {label_norm: (sheet, row)} pour toutes les cellules VIDES
    (pas de formule) dans Actif, Passif, CPC.
    """
    index = {}
    for sheet_name in ['2 - Bilan Actif', '3 - Bilan Passif', '4 - CPC']:
        if sheet_name not in wb.sheetnames:
            continue
        ws = wb[sheet_name]
        for row in ws.iter_rows(min_row=5, max_row=ws.max_row):
            lbl = row[1].value  # col B
            if not lbl or str(lbl).startswith(('▶', ' ▶', '=')):
                continue
            # Ignorer si col C est une formule
            c_val = row[2].value
            if isinstance(c_val, str) and c_val.startswith('='):
                continue
            n = soft(str(lbl))
            if n and n not in SKIP and len(n) > 2:
                index[n] = (sheet_name, row[0].row)

    logger.info(f"Index : {len(index)} cellules indexées")
    return index


def find_cell(label: str, index: dict) -> tuple | None:
    """Cherche la cellule pour un label (avec fallback sans accents)."""
    n = ALIASES.get(label, label)

    # 1. Cherche directe
    found = index.get(n)
    if found:
        return found

    # 2. Fallback sans accents
    n_na = no_accent(n)
    for k, v in index.items():
        if no_accent(k) == n_na:
            return v

    return None


def inject(extracted: dict, template_path: str, output_path: str) -> dict:
    """
    Injecte les données extraites dans le template Excel.
    Retourne les stats d'injection.
    """
    import shutil
    shutil.copy(template_path, output_path)
    wb = openpyxl.load_workbook(output_path)
    idx = build_index(wb)

    injected = 0
    skipped  = []

    for section, data in [
        ('actif',  extracted.get('actif_values',  {})),
        ('passif', extracted.get('passif_values', {})),
        ('cpc',    extracted.get('cpc_values',    {})),
    ]:
        for label, vals in data.items():
            # Ignorer les totaux et labels parasites
            n = soft(label)
            if n in SKIP or len(n) < 3:
                continue

            ci = find_cell(label, idx)
            if not ci:
                skipped.append(f"[{section}] {label}")
                continue

            sheet, row = ci
            ws = wb[sheet]

            if section == 'actif':
                # [brut, amort, net_n1]
                brut, amort, net_n1 = (vals + [None, None, None])[:3]
                if brut   is not None: ws.cell(row, 3).value = brut;   injected += 1
                if amort  is not None: ws.cell(row, 4).value = amort;  injected += 1
                if net_n1 is not None: ws.cell(row, 6).value = net_n1; injected += 1

            elif section == 'passif':
                # [val_n, val_n1]
                val_n, val_n1 = (vals + [None, None])[:2]
                if val_n  is not None: ws.cell(row, 3).value = val_n;  injected += 1
                if val_n1 is not None: ws.cell(row, 4).value = val_n1; injected += 1

            elif section == 'cpc':
                # [propre_n, prec_n, total_n1]
                propre, prec, total_n1 = (vals + [None, None, None])[:3]
                if propre   is not None: ws.cell(row, 3).value = propre;   injected += 1
                if prec     is not None: ws.cell(row, 4).value = prec;     injected += 1
                if total_n1 is not None: ws.cell(row, 6).value = total_n1; injected += 1

    # Mettre à jour les headers
    _update_headers(wb, extracted.get('info', {}))

    wb.save(output_path)
    logger.info(f"Injection : {injected} valeurs · {len(skipped)} non mappés")
    return {'injected': injected, 'skipped': skipped}


def _update_headers(wb, info: dict):
    raison   = info.get('raison_sociale') or '—'
    id_fisc  = info.get('identifiant_fiscal') or ''
    exercice = info.get('exercice') or ''
    sub = f"{raison}  —  IF: {id_fisc}" if id_fisc else raison

    for sheet, title in [
        ('2 - Bilan Actif',   f'BILAN — ACTIF  |  {exercice}'),
        ('3 - Bilan Passif',  f'BILAN — PASSIF  |  {exercice}'),
        ('4 - CPC',           f'COMPTE DE PRODUITS ET CHARGES  |  {exercice}'),
    ]:
        if sheet not in wb.sheetnames:
            continue
        ws = wb[sheet]
        # Ligne 1 col B = titre, ligne 2 col A = sous-titre
        if ws.cell(1, 2).value:
            ws.cell(1, 2).value = title
        ws.cell(2, 1).value = sub
