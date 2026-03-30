"""
core/pdf_to_excel.py — Extraction PDF fiscal → Excel
Approche hybride :
  1. extract_tables() si le PDF a des bordures claires (Bilan2017, BORJ)
  2. extract_words() + X/Y si pas de tableau structuré (SGTM, SAPST)
Aucun mapping, aucun template — miroir fidèle du PDF.
"""

import re, unicodedata
from collections import defaultdict
import pdfplumber
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from utils.logger import get_logger

logger = get_logger(__name__)

# ── Couleurs ──────────────────────────────────────────────────────────────────
C_DARK   = "1F3864"
C_MED    = "2E75B6"
C_LIGHT  = "D6E4F0"
C_TOTAL  = "1F3864"
C_RESULT = "2E4057"
C_WHITE  = "FFFFFF"
C_GRAY   = "F5F7FA"
C_BORDER = "B8CCE4"
NUM_FMT  = '#,##0.00;(#,##0.00);"-"'

# ── Mots-clés sections ────────────────────────────────────────────────────────
SKIP_LABELS = {
    'brut', 'net', 'amortissements', 'provisions', 'exercice',
    'exercice precedent', 'a c t i f', 'p a s s i f',
    'tableau n 1', 'tableau n 2', 'bilan actif', 'bilan passif',
    'compte de produits', 'designation', 'operations',
    'propres a l exercice', 'concernant les exercices',
    'totaux de l exercice',
}
SKIP_RE = re.compile(
    r'^(tableau\s*n|bilan\s*\(|compte\s*de\s*produits|agence\s*du|'
    r'identifiant\s*fiscal|exercice\s*du|fes\s*le|casablanca\s*le|'
    r'signature|cadre\s*reserve|\(1\)|\(2\)|nb\s*:|\d+$)',
    re.IGNORECASE
)

TOTAL_RE   = re.compile(r'^(total\s+[ivi0-9]|total\s+g[eé]n|total\s+a\+)', re.I)
RESULT_RE  = re.compile(r'^(r[eé]sultat|impots?\s+sur|impôts?\s+sur)', re.I)
SECTION_RE = re.compile(r'^[A-ZÀÂÉÈÊÎÔÙÛÇ\s]{6,}$')


# ═══════════════════════════════════════════════════════════════════════════════
# UTILITAIRES
# ═══════════════════════════════════════════════════════════════════════════════

def _norm(s):
    s = unicodedata.normalize('NFD', s).encode('ascii','ignore').decode().lower()
    return re.sub(r'\s+', ' ', re.sub(r'[^\w\s]', ' ', s)).strip()

def _parse_fr(s):
    """Parse un nombre FR : '1 234 567,89' ou '1.234.567,89' → float."""
    if not s: return None
    s = str(s).strip().replace('\xa0','').replace(' ','')
    if not s or s in ['-','—','/']: return None
    neg = s.startswith('-'); s = s.lstrip('-')
    # Multi-points : 1.234.567,89
    m = re.match(r'^(\d{1,3}(?:\.\d{3})*),(\d{2})$', s)
    if m: s = m.group(1).replace('.','') + '.' + m.group(2)
    elif re.match(r'^\d+,\d{2}$', s): s = s.replace(',','.')
    elif re.match(r'^\d+$', s): pass
    else: return None
    try: return -(float(s)) if neg else float(s)
    except: return None

def _is_num_tok(t):
    return bool(re.match(r'^-?\d+$', t.replace(',','').replace('.',''))) \
           and len(t.replace(',','').replace('.','')) >= 1

def _row_type(label):
    if TOTAL_RE.match(label): return 'total'
    if RESULT_RE.match(label): return 'result'
    if SECTION_RE.match(label.strip()) and len(label.strip()) > 5: return 'section'
    return 'normal'

def _should_skip(label):
    if not label or len(label.strip()) < 2: return True
    n = _norm(label)
    if n in SKIP_LABELS: return True
    if SKIP_RE.match(label.strip()): return True
    if re.match(r'^\d+$', n): return True
    return False


# ═══════════════════════════════════════════════════════════════════════════════
# EXTRACTION : MÉTHODE 1 — extract_tables()
# ═══════════════════════════════════════════════════════════════════════════════

def _table_usable(table):
    """Vérifie qu'un tableau a assez de lignes avec des données numériques."""
    if not table or len(table) < 4: return False
    n_cols = len(table[0]) if table[0] else 0
    if n_cols < 2: return False
    # Compter les lignes avec au moins 1 nombre
    data_rows = sum(
        1 for row in table[3:]
        if any(_parse_fr(str(c)) is not None for c in row if c)
    )
    return data_rows >= 3


def _extract_via_tables(page):
    """Extrait les lignes depuis extract_tables()."""
    tables = page.extract_tables()
    rows = []
    for t in tables:
        if not _table_usable(t): continue
        n_cols = len(t[0])

        for row in t:
            if not row: continue
            cells = [str(c).strip().replace('\n',' ') if c else '' for c in row]

            # Trouver le label (première cellule non-numérique de longueur > 2)
            label = ''
            for c in cells:
                if c and _parse_fr(c) is None and len(c) > 2:
                    label = c
                    break
            if not label: continue

            # Valeurs numériques dans les autres colonnes
            vals = [_parse_fr(c) for c in cells if _parse_fr(c) is not None]

            if label and (vals or _row_type(label) in ('total','result','section')):
                rows.append((label.strip(), vals))

    return rows


# ═══════════════════════════════════════════════════════════════════════════════
# EXTRACTION : MÉTHODE 2 — X/Y (fallback)
# ═══════════════════════════════════════════════════════════════════════════════

def _extract_via_xy(page):
    """Extrait les lignes depuis extract_words() avec positionnement X/Y."""
    words = page.extract_words(x_tolerance=3, y_tolerance=3)
    if not words: return []

    # Seuil X entre labels et valeurs numériques
    num_words = [w for w in words if _is_num_tok(w['text']) and w['x0'] > 100]
    if not num_words: return []
    thresh = min(w['x0'] for w in num_words) - 5

    # Grouper par Y
    lines = defaultdict(list)
    for w in words:
        lines[round(w['top'] / 3) * 3].append(w)

    rows = []
    prev_label = None

    for y in sorted(lines.keys()):
        row = sorted(lines[y], key=lambda w: w['x0'])
        lw  = [w for w in row if w['x0'] < thresh]
        nw  = [w for w in row if w['x0'] >= thresh and _is_num_tok(w['text'])]

        # Reconstruire le label (filtrer lettres rotatives x<50 longueur<=1)
        filtered = [w for w in lw
                    if not (len(w['text']) <= 1
                            and re.match(r'^[A-Z.]$', w['text'])
                            and w['x0'] < 50)]
        label = ''
        if filtered:
            label = filtered[0]['text']
            for i in range(1, len(filtered)):
                gap = filtered[i]['x0'] - filtered[i-1]['x1']
                label += filtered[i]['text'] if gap <= 1 else ' ' + filtered[i]['text']
            label = re.sub(r'\s+', ' ', label).strip()

        # Fusionner les tokens numériques adjacents → valeurs
        vals = []
        if nw:
            nw_s = sorted(nw, key=lambda w: w['x0'])
            grp  = [nw_s[0]]
            for w in nw_s[1:]:
                if w['x0'] - grp[-1]['x1'] < 18:
                    grp.append(w)
                else:
                    raw = ''.join(x['text'] for x in grp)
                    v   = _parse_fr(raw)
                    if v is not None: vals.append(v)
                    grp = [w]
            raw = ''.join(x['text'] for x in grp)
            v   = _parse_fr(raw)
            if v is not None: vals.append(v)

        # Ligne orpheline de valeurs → rattacher au label précédent
        if not label and vals and prev_label:
            # Chercher la dernière ligne du résultat et ajouter les valeurs
            if rows and rows[-1][0] == prev_label:
                existing = rows[-1][1]
                if not existing:
                    rows[-1] = (prev_label, vals)
                continue

        if label:
            prev_label = label

        if label and (vals or _row_type(label) in ('total','result','section')):
            rows.append((label, vals))

    return rows


# ═══════════════════════════════════════════════════════════════════════════════
# SÉLECTION AUTOMATIQUE DE LA MÉTHODE
# ═══════════════════════════════════════════════════════════════════════════════

def _extract_page(page):
    """Choisit la meilleure méthode d'extraction pour une page."""
    # Essayer extract_tables() d'abord
    rows_table = _extract_via_tables(page)
    if len(rows_table) >= 5:
        logger.info(f"  → extract_tables() : {len(rows_table)} lignes")
        return rows_table, 'tables'

    # Fallback X/Y
    rows_xy = _extract_via_xy(page)
    logger.info(f"  → X/Y fallback : {len(rows_xy)} lignes")
    return rows_xy, 'xy'


# ═══════════════════════════════════════════════════════════════════════════════
# INFOS GÉNÉRALES
# ═══════════════════════════════════════════════════════════════════════════════

def _extract_info(pdf):
    info = {}
    for i in range(min(2, len(pdf.pages))):
        text = pdf.pages[i].extract_text() or ''
        for key, pat in [
            ('raison_sociale',      r'[Rr]aison\s+[Ss]ociale\s*[:\-]?\s*([A-Z][^\n]{3,60})'),
            ('identifiant_fiscal',  r'[Ii]dentifiant\s+[Ff]iscal\s*[:\-]?\s*(\d+)'),
            ('taxe_pro',            r'[Tt]axe\s+[Pp]rof\w*\.?\s*[:\-]?\s*([\d\s]+)'),
            ('adresse',             r'[Aa]dresse\s*[:\-]?\s*([^\n]{5,60})'),
        ]:
            if key not in info:
                m = re.search(pat, text, re.IGNORECASE)
                if m: info[key] = m.group(1).strip()

        if 'exercice' not in info:
            m = re.search(r'(\d{2}/\d{2}/\d{4})\s+au\s+(\d{2}/\d{2}/\d{4})', text)
            if m:
                info['exercice']     = f"Du {m.group(1)} au {m.group(2)}"
                info['exercice_fin'] = m.group(2)

    for k in ('raison_sociale','identifiant_fiscal','exercice','exercice_fin','taxe_pro','adresse'):
        info.setdefault(k, '')
    return info


# ═══════════════════════════════════════════════════════════════════════════════
# STYLES EXCEL
# ═══════════════════════════════════════════════════════════════════════════════

def _border():
    s = Side(style='thin', color=C_BORDER)
    return Border(top=s, bottom=s, left=s, right=s)

def _cell(ws, r, c, value=None, bg=C_WHITE, fg='222222', bold=False,
          align='left', num_fmt=None, size=9, wrap=True, indent=0):
    cell = ws.cell(row=r, column=c, value=value)
    cell.font      = Font(name='Arial', size=size, bold=bold, color=fg)
    cell.fill      = PatternFill('solid', fgColor=bg)
    cell.alignment = Alignment(horizontal=align, vertical='center',
                               wrap_text=wrap, indent=indent)
    cell.border    = _border()
    if num_fmt: cell.number_format = num_fmt
    return cell

def _row_style(row_type):
    """Retourne (bg, fg, bold) selon le type de ligne."""
    if row_type == 'total':   return C_TOTAL,  C_WHITE, True
    if row_type == 'result':  return C_RESULT, C_WHITE, True
    if row_type == 'section': return C_LIGHT,  C_DARK,  True
    return C_WHITE, '222222', False


# ═══════════════════════════════════════════════════════════════════════════════
# ÉCRITURE DES FEUILLES
# ═══════════════════════════════════════════════════════════════════════════════

def _write_identification(wb, info):
    ws = wb.create_sheet('1 — Identification')
    ws.sheet_view.showGridLines = False
    ws.column_dimensions['A'].width = 28
    ws.column_dimensions['B'].width = 50

    ws.merge_cells('A1:B1')
    _cell(ws, 1, 1, 'PIÈCES ANNEXES À LA DÉCLARATION FISCALE',
          bg=C_DARK, fg=C_WHITE, bold=True, align='center', size=13)
    ws.row_dimensions[1].height = 26

    ws.merge_cells('A2:B2')
    _cell(ws, 2, 1, 'IMPÔTS SUR LES SOCIÉTÉS — Modèle Comptable Normal (loi 9-88)',
          bg=C_MED, fg=C_WHITE, align='center', size=10)
    ws.row_dimensions[2].height = 18

    fields = [
        ('Raison sociale',        info['raison_sociale']),
        ('Identifiant fiscal',    info['identifiant_fiscal']),
        ('Taxe professionnelle',  info['taxe_pro']),
        ('Adresse',               info['adresse']),
        ('Exercice',              info['exercice']),
    ]
    for i, (lbl, val) in enumerate(fields, 4):
        ws.row_dimensions[i].height = 18
        _cell(ws, i, 1, lbl, bg=C_LIGHT, fg=C_DARK, bold=True, size=9, indent=1)
        _cell(ws, i, 2, val, bg=C_WHITE,  fg='222222', size=9, indent=1)


def _write_data_sheet(wb, sheet_name, headers, rows):
    """
    Écrit une feuille de données.
    headers : liste de strings pour les colonnes valeurs
    rows    : [(label, [val1, val2, ...]), ...]
    """
    ws = wb.create_sheet(sheet_name)
    ws.sheet_view.showGridLines = False

    # Déterminer le nombre de colonnes valeurs
    max_vals = max((len(vals) for _, vals in rows if vals), default=len(headers))
    n_headers = max(len(headers), max_vals)
    total_cols = 1 + n_headers  # col 1 = label, cols 2..n = valeurs

    # Largeurs
    ws.column_dimensions['A'].width = 50
    for c in range(2, total_cols + 1):
        ws.column_dimensions[get_column_letter(c)].width = 18

    # Titre
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=total_cols)
    _cell(ws, 1, 1, sheet_name.split('—')[-1].strip(),
          bg=C_DARK, fg=C_WHITE, bold=True, align='center', size=12)
    ws.row_dimensions[1].height = 22

    # Headers colonnes
    _cell(ws, 2, 1, 'DÉSIGNATION', bg=C_MED, fg=C_WHITE, bold=True,
          align='center', size=9)
    for ci, h in enumerate(headers[:n_headers], 2):
        _cell(ws, 2, ci, h, bg=C_MED, fg=C_WHITE, bold=True,
              align='center', size=9, wrap=True)
    ws.row_dimensions[2].height = 28
    ws.freeze_panes = 'A3'

    # Données
    r = 3
    for label, vals in rows:
        rtype = _row_type(label)
        bg, fg, bold = _row_style(rtype)
        ws.row_dimensions[r].height = 15 if rtype == 'normal' else 17

        # Label
        indent = 1 if rtype == 'normal' else 0
        _cell(ws, r, 1, label, bg=bg, fg=fg, bold=bold,
              align='left', size=9, indent=indent)

        # Valeurs
        for ci, v in enumerate(vals[:n_headers], 2):
            _cell(ws, r, ci, v, bg=bg, fg=fg, bold=bold,
                  align='right', size=9,
                  num_fmt=NUM_FMT if isinstance(v, (int, float)) else None)

        # Cellules vides restantes
        for ci in range(len(vals) + 2, total_cols + 1):
            _cell(ws, r, ci, bg=bg, fg=fg)

        r += 1

    return r - 3  # nb lignes écrites


# ═══════════════════════════════════════════════════════════════════════════════
# DÉTECTION DES HEADERS PDF
# ═══════════════════════════════════════════════════════════════════════════════

def _detect_headers(page, thresh_y=80):
    """
    Extrait les labels des colonnes valeurs depuis les premières lignes de la page.
    Retourne une liste de strings.
    """
    text = page.extract_text() or ''
    lines = text.split('\n')

    # Chercher les lignes qui ressemblent à des headers de colonnes
    header_kw = {
        'brut': 'Brut',
        'amortissements': 'Amort. & Prov.',
        'net': 'Net (N)',
        'exercice precedent': 'Net (N-1)',
        'precedent': 'Net (N-1)',
        'exercice n': 'Exercice N',
        'propres': 'Propres N',
        'exercices prec': 'Exerc. Préc.',
        'totaux': 'Total N',
    }
    found = []
    for line in lines[:10]:
        n = _norm(line)
        for kw, label in header_kw.items():
            if kw in n and label not in found:
                found.append(label)

    # Fallback : noms génériques selon nb de colonnes
    return found or ['Valeur 1', 'Valeur 2', 'Valeur 3', 'Valeur 4']


# ═══════════════════════════════════════════════════════════════════════════════
# POINT D'ENTRÉE
# ═══════════════════════════════════════════════════════════════════════════════

# Noms des feuilles selon la page
PAGE_SHEETS = {
    0: None,          # Page 1 → ignorée (infos, gérée séparément)
    1: '2 — Bilan Actif',
    2: '3 — Bilan Passif',
    3: '4 — CPC (1)',
    4: '5 — CPC (2)',
    5: '6 — CPC (3)',
    6: '7 — Annexes',
}

# DGI : 7 pages → décalage
PAGE_SHEETS_DGI = {
    0: None,
    1: '2 — Bilan Actif (1)',
    2: '2 — Bilan Actif (2)',
    3: '3 — Bilan Passif',
    4: '4 — CPC (1)',
    5: '5 — CPC (2)',
    6: '6 — CPC (3)',
}


def convert(pdf_path: str, output_path: str) -> dict:
    """
    Convertit un PDF fiscal en Excel structuré.
    Retourne dict avec info, tables, rows, pages.
    """
    pdf    = pdfplumber.open(pdf_path)
    n      = len(pdf.pages)
    info   = _extract_info(pdf)
    is_dgi = (n == 7)

    page_map = PAGE_SHEETS_DGI if is_dgi else PAGE_SHEETS

    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    _write_identification(wb, info)

    total_tables = 0
    total_rows   = 0
    used_names   = set()

    for page_idx, page in enumerate(pdf.pages):
        sheet_name = page_map.get(page_idx)
        if sheet_name is None:
            continue

        # Éviter les doublons de nom de feuille
        base_name = sheet_name
        cnt = 2
        while sheet_name in used_names:
            sheet_name = f"{base_name} ({cnt})"
            cnt += 1
        used_names.add(sheet_name)

        logger.info(f"Page {page_idx+1} → '{sheet_name}'")
        rows, method = _extract_page(page)

        # Filtrer les lignes parasites
        rows = [(lbl, vals) for lbl, vals in rows
                if not _should_skip(lbl)]

        if not rows:
            logger.info(f"  → page vide, ignorée")
            continue

        # Détecter les headers colonnes depuis la page
        headers = _detect_headers(page)

        n_written = _write_data_sheet(wb, sheet_name, headers, rows)
        total_tables += 1
        total_rows   += n_written
        logger.info(f"  → {n_written} lignes écrites (méthode: {method})")

    pdf.close()
    wb.save(output_path)

    logger.info(f"Excel sauvegardé : {total_tables} feuilles, {total_rows} lignes")
    return {
        'info':   info,
        'tables': total_tables,
        'rows':   total_rows,
        'pages':  n,
    }
