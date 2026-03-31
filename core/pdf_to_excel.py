"""
core/pdf_to_excel.py
Extraction directe PDF → Excel.
Principe : extract_tables() → une feuille Excel par page, données brutes.
Pas de mapping, pas de fuzzy — ce qu'on voit dans le PDF, on le met dans l'Excel.
"""

import re
import pdfplumber
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from utils.logger import get_logger

logger = get_logger(__name__)


# ── Nettoyage ─────────────────────────────────────────────────────────────────

def is_rotated(v: str) -> bool:
    """Détecte les lettres rotatives (A\nC\nT\nI\nF → colonne décorative)."""
    if not v: return False
    parts = [p.strip() for p in str(v).split('\n') if p.strip()]
    return len(parts) >= 3 and all(len(p) <= 2 and p.isalpha() for p in parts)


def clean_cell(v) -> str:
    """Nettoie une cellule PDF."""
    if v is None: return ''
    s = str(v).strip()
    if is_rotated(s): return ''
    # Uniformiser les sauts de ligne
    s = re.sub(r'\n+', ' ', s).strip()
    return s


def parse_number(s: str):
    """
    Convertit une chaîne en nombre.
    Gère : '1 234 567,89' | '1.234.567,89' | '-26 682 992,16'
    Retourne float ou None.
    """
    if not s: return None
    s = s.strip().replace('\xa0', '').replace(' ', '')
    neg = s.startswith('-')
    s = s.lstrip('-').lstrip('+')

    # Format FR : séparateurs points ou espaces, virgule décimale
    if re.match(r'^\d{1,3}(\.\d{3})*,\d{2}$', s):
        s = s.replace('.', '').replace(',', '.')
    elif re.match(r'^\d+,\d{2}$', s):
        s = s.replace(',', '.')
    elif re.match(r'^\d+$', s):
        pass  # entier
    else:
        return None

    try:
        v = float(s)
        return -v if neg else v
    except ValueError:
        return None


# ── Styles ────────────────────────────────────────────────────────────────────

HEADER_FILL  = PatternFill("solid", fgColor="1F3864")
HEADER_FONT  = Font(color="FFFFFF", bold=True, size=10)
SECTION_FILL = PatternFill("solid", fgColor="D9E1F2")
SECTION_FONT = Font(bold=True, size=9)
DATA_FONT    = Font(size=9)
NUMBER_FORMAT = '#,##0.00'

THIN = Side(style='thin', color="CCCCCC")
BORDER = Border(left=THIN, right=THIN, bottom=THIN, top=THIN)

def style_header_row(ws, row_num, max_col):
    for c in range(1, max_col + 1):
        cell = ws.cell(row_num, c)
        if cell.value:
            cell.fill = HEADER_FILL
            cell.font = HEADER_FONT
            cell.alignment = Alignment(horizontal='center', wrap_text=True)

def style_data_row(ws, row_num, max_col, is_total=False):
    for c in range(1, max_col + 1):
        cell = ws.cell(row_num, c)
        if cell.value is not None:
            if is_total:
                cell.font = Font(bold=True, size=9)
                cell.fill = SECTION_FILL
            else:
                cell.font = DATA_FONT
            if isinstance(cell.value, (int, float)):
                cell.number_format = NUMBER_FORMAT
                cell.alignment = Alignment(horizontal='right')
            cell.border = BORDER

def auto_width(ws):
    for col in ws.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            try:
                val = str(cell.value) if cell.value else ''
                max_len = max(max_len, len(val))
            except: pass
        ws.column_dimensions[col_letter].width = min(max(max_len + 2, 8), 45)


# ── Extraction principale ─────────────────────────────────────────────────────

PAGE_NAMES = {
    1: 'Identification',
    2: 'Bilan Actif',
    3: 'Bilan Passif',
    4: 'CPC (1)',
    5: 'CPC (2)',
    6: 'CPC (3)',
    7: 'Annexes',
}

TOTAL_KEYWORDS = {
    'total', 'sous-total', 'résultat', 'resultat',
    'chiffre d\'affaires', 'total i', 'total ii', 'total iii',
}


def is_total_row(row_vals: list) -> bool:
    first = next((v for v in row_vals if v and len(str(v)) > 3), '')
    return any(kw in str(first).lower() for kw in TOTAL_KEYWORDS)


def write_table_to_sheet(ws, table, start_row: int = 1) -> int:
    """
    Écrit un tableau dans une feuille Excel.
    Retourne le nombre de lignes écrites.
    """
    header_rows = {0, 1, 2, 3}  # Les premières lignes sont des headers
    written = 0

    for rel_r, row in enumerate(table):
        abs_r = start_row + rel_r
        row_vals = []

        for c, cell in enumerate(row, 1):
            raw = clean_cell(cell)
            if not raw:
                continue
            # Essayer de convertir en nombre
            num = parse_number(raw)
            val = num if num is not None else raw
            ws.cell(abs_r, c).value = val
            row_vals.append(val)

        # Appliquer les styles
        max_col = max((c for c in range(1, ws.max_column + 1)
                       if ws.cell(abs_r, c).value is not None), default=1)

        if rel_r in header_rows:
            style_header_row(ws, abs_r, max_col)
        else:
            style_data_row(ws, abs_r, max_col, is_total=is_total_row(row_vals))

        written += 1

    return written


def extract_info(pdf) -> dict:
    """Extrait les infos générales depuis la page 1."""
    info = {}
    try:
        text = pdf.pages[0].extract_text() or ''
    except Exception:
        return {}

    for key, pat in [
        ('raison_sociale',     r'[Rr]aison\s+[Ss]ociale\s*[:\-]?\s*([A-Z][^\n]{3,60})'),
        ('identifiant_fiscal', r'[Ii]dentifiant\s+[Ff]iscal\s*[:\-]?\s*(\d+)'),
        ('taxe_pro',           r'[Tt]axe\s+[Pp]rof\w*\.?\s*[:\-]?\s*([\d\s]+)'),
        ('adresse',            r'[Aa]dresse\s*[:\-]?\s*([^\n]{5,60})'),
    ]:
        m = re.search(pat, text, re.IGNORECASE)
        if m:
            info[key] = m.group(1).strip()

    m = re.search(r'(\d{2}/\d{2}/\d{4})\s+au\s+(\d{2}/\d{2}/\d{4})', text)
    if m:
        info['exercice']   = f"Du {m.group(1)} au {m.group(2)}"
        info['exercice_fin'] = m.group(2)

    return info


def write_info_sheet(wb, info: dict):
    """Crée la feuille d'identification."""
    ws = wb.create_sheet('Identification', 0)

    ws.cell(1, 1).value = "PIÈCES ANNEXES À LA DÉCLARATION FISCALE"
    ws.cell(1, 1).font  = Font(bold=True, size=14, color="1F3864")
    ws.cell(2, 1).value = "IMPÔTS SUR LES SOCIÉTÉS — Modèle Comptable Normal"
    ws.cell(2, 1).font  = Font(italic=True, size=11)

    ws.cell(4, 1).value = "Champ"
    ws.cell(4, 2).value = "Valeur"
    style_header_row(ws, 4, 2)

    rows = [
        ("Raison sociale",     info.get('raison_sociale', '')),
        ("Identifiant fiscal", info.get('identifiant_fiscal', '')),
        ("Taxe professionnelle",info.get('taxe_pro', '')),
        ("Adresse",            info.get('adresse', '')),
        ("Exercice",           info.get('exercice', '')),
    ]
    for i, (k, v) in enumerate(rows, 5):
        ws.cell(i, 1).value = k
        ws.cell(i, 2).value = v
        ws.cell(i, 1).font  = Font(bold=True, size=10)
        ws.cell(i, 2).font  = Font(size=10)

    ws.column_dimensions['A'].width = 25
    ws.column_dimensions['B'].width = 50


def convert(pdf_path: str, output_path: str) -> dict:
    """
    Convertit un PDF fiscal en Excel.
    Retourne les stats de conversion.
    """
    pdf   = pdfplumber.open(pdf_path)
    n     = len(pdf.pages)
    wb    = openpyxl.Workbook()
    wb.remove(wb.active)

    # Infos générales → feuille 1
    info = extract_info(pdf)
    write_info_sheet(wb, info)

    total_tables = 0
    total_rows   = 0

    for i, page in enumerate(pdf.pages, 1):
        tables = page.extract_tables()
        if not tables:
            continue

        # Filtrer les micro-tableaux (< 3 lignes ou < 2 colonnes)
        tables = [t for t in tables
                  if len(t) >= 3 and len(t[0]) >= 2]
        if not tables:
            continue

        page_name = PAGE_NAMES.get(i, f'Page {i}')
        ws = wb.create_sheet(page_name)
        cur_row = 1

        for t in tables:
            rows_written = write_table_to_sheet(ws, t, cur_row)
            cur_row += rows_written + 2  # 2 lignes vides entre tableaux
            total_tables += 1
            total_rows   += rows_written

        auto_width(ws)
        logger.info(f"Page {i} ({page_name}) : {len(tables)} tableau(x) → {cur_row-1} lignes")

    pdf.close()

    # Mettre à jour les headers des feuilles avec la raison sociale
    raison   = info.get('raison_sociale', '')
    exercice = info.get('exercice', '')
    for sheet_name in ['Bilan Actif', 'Bilan Passif', 'CPC (1)', 'CPC (2)']:
        if sheet_name in wb.sheetnames:
            ws2 = wb[sheet_name]
            # Insérer une ligne titre en haut
            ws2.insert_rows(1)
            ws2.cell(1, 1).value = f"{sheet_name.upper()}  —  {raison}  |  {exercice}"
            ws2.cell(1, 1).font  = Font(bold=True, size=11, color="1F3864")

    wb.save(output_path)
    logger.info(f"Excel généré : {total_tables} tableaux, {total_rows} lignes")

    return {
        'tables':    total_tables,
        'rows':      total_rows,
        'pages':     n,
        'info':      info,
    }
