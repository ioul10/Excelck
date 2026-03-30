"""
core/pdf_to_excel.py
Extraction PDF fiscal → Excel structuré et formaté professionnellement.
Détecte automatiquement les sections (Actif / Passif / CPC) par mots-clés.
Compatible avec tous les formats MCN loi 9-88 Maroc.
"""

import re
import unicodedata
import pdfplumber
import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ── Palette couleurs ──────────────────────────────────────────────────────────
C_DARK_BLUE  = "1F3864"
C_MED_BLUE   = "2E75B6"
C_LIGHT_BLUE = "BDD7EE"
C_SECTION    = "D6E4F0"
C_SUBTOTAL   = "EBF3FB"
C_RESULT     = "2E4057"
C_WHITE      = "FFFFFF"
C_GRAY_BG    = "F5F7FA"
C_BORDER     = "B8CCE4"

NUM_FMT = '#,##0.00;(#,##0.00);"-"'

# ── Mots-clés de détection des sections ──────────────────────────────────────
ACTIF_KW  = ["actif immobilise", "immobilisations incorporelles", "bilan (actif",
              "bilan actif", "actif immobilisé", "frais preliminaires",
              "frais préliminaires", "immobilisations en non valeur"]
PASSIF_KW = ["capitaux propres", "bilan (passif", "bilan passif",
              "capital social", "passif circulant", "dettes de financement",
              "p a s s i f"]
CPC_KW    = ["produits exploitation", "charges exploitation",
             "compte de produits", "ventes de marchandises",
             "chiffre d affaires", "chiffres d affaires",
             "produits d exploitation", "charges d exploitation"]

# ── Labels à ignorer (en-têtes parasites) ────────────────────────────────────
SKIP_EXACT = {
    "brut", "net", "designation", "operations", "exercice", "exercice precedent",
    "a c t i f", "p a s s i f", "nb:", "1", "2", "3 = 2 + 1", "4",
    "propres a l exercice", "concernant les exercices precedents",
    "totaux de l exercice", "totaux de l exercice precedent",
    "tableau n 1(1/2)", "tableau n 1(2/2)", "tableau n 2(1/2)", "tableau n 2(2/2)",
    "bilan (actif) (modele normal)", "bilan (passif) (modele normal)",
    "compte de produits et charges", "amortissements et provi",
    "amortissements et provisions",
}
SKIP_PREFIX = (
    "tableau n", "bilan (", "compte de produits", "agence du",
    "identifiant fiscal", "exercice du", "(1)capital", "(2)benefic",
    "1)variation", "2)achats revendus", "fes le", "signature",
    "cadre reserve",
)

# ── Patterns totaux / résultats (pas injectés comme données simples) ──────────
TOTAL_KW   = ("total i ", "total ii", "total iii", "total general",
              "total (a+b", "total i+ii", "total iv", "total v ",
              "total vi", "total vii", "total viii", "total ix",
              "total xiv", "total xv")
RESULT_KW  = ("resultat d exploitation", "resultat financier",
              "resultat courant", "resultat non courant",
              "resultat avant impot", "resultat net",
              "impots sur les benefices", "impots sur les bénéfices")
SECTION_KW = ("produits d exploitation", "charges d exploitation",
              "produits financiers", "charges financieres",
              "produits non courant", "charges non courant",
              "capitaux propres", "dettes de financement",
              "passif circulant", "tresorerie")


# ═══════════════════════════════════════════════════════════════════════════════
# UTILITAIRES
# ═══════════════════════════════════════════════════════════════════════════════

def _norm(s: str) -> str:
    s = unicodedata.normalize('NFD', s)
    s = s.encode('ascii', 'ignore').decode('utf-8').lower()
    s = re.sub(r"[^\w\s]", " ", s)
    return re.sub(r"\s+", " ", s).strip()

def _parse_num(s) -> float | None:
    if s is None:
        return None
    s = str(s).strip().replace("\n", "").replace("\xa0", "")
    if not s or s in ["-", "—", "", "None", "/"]:
        return None
    neg = False
    if s.startswith("(") and s.endswith(")"):
        neg, s = True, s[1:-1]
    if s.startswith("-"):
        neg, s = True, s[1:]
    s = s.replace(" ", "").replace(",", ".")
    try:
        v = float(s)
        return -v if neg else v
    except ValueError:
        return None

def _clean_label(s) -> str:
    if not s:
        return ""
    s = str(s).replace("\n", " ").strip()
    s = re.sub(r"\s{2,}", " ", s)
    s = re.sub(r"^(I{1,3}|IV|V|VI{1,3}|IX|X{1,3})\s+", "", s)
    return s.strip()

def _should_skip(label: str) -> bool:
    n = _norm(label)
    if not n or len(n) < 2:
        return True
    if n in SKIP_EXACT:
        return True
    if any(n.startswith(p) for p in SKIP_PREFIX):
        return True
    if re.match(r"^\d+$", n):
        return True
    return False

def _row_type(label: str) -> str:
    """Détermine le type visuel d'une ligne selon son label."""
    n = _norm(label)
    if any(n.startswith(k) for k in TOTAL_KW) or "total general" in n:
        return "total"
    if any(n.startswith(k) for k in RESULT_KW):
        return "result"
    if any(n.startswith(k) for k in SECTION_KW):
        return "section"
    # Lignes en majuscules = section
    stripped = label.strip()
    if stripped and stripped == stripped.upper() and len(stripped) > 4:
        return "section"
    return "normal"


# ═══════════════════════════════════════════════════════════════════════════════
# EXTRACTION PDF
# ═══════════════════════════════════════════════════════════════════════════════

class PDFExtractor:

    def __init__(self, pdf_path: str):
        self.pdf    = pdfplumber.open(pdf_path)
        self.pages  = self.pdf.pages
        self.n      = len(self.pages)

    def extract(self) -> dict:
        ranges  = self._detect_ranges()
        info    = self._extract_info()
        actif   = self._extract_section(ranges["actif"],  mode="actif")
        passif  = self._extract_section(ranges["passif"], mode="passif")
        cpc     = self._extract_section(ranges["cpc"],    mode="cpc")
        self.pdf.close()
        return {"info": info, "actif": actif, "passif": passif, "cpc": cpc,
                "pages": self.n}

    def _detect_ranges(self) -> dict:
        ranges = {"actif": [], "passif": [], "cpc": []}
        for i, page in enumerate(self.pages):
            t = _norm(page.extract_text() or "")
            if any(k in t for k in ACTIF_KW):
                ranges["actif"].append(i)
            if any(k in t for k in PASSIF_KW):
                ranges["passif"].append(i)
            if any(k in t for k in CPC_KW):
                ranges["cpc"].append(i)
        # Fallback
        if not ranges["actif"]:
            ranges["actif"] = list(range(1, min(3, self.n)))
        if not ranges["passif"]:
            ranges["passif"] = list(range(2, min(5, self.n)))
        if not ranges["cpc"]:
            ranges["cpc"] = list(range(3, min(7, self.n)))
        return ranges

    def _extract_info(self) -> dict:
        info = {}
        for i in range(min(2, self.n)):
            tables = self.pages[i].extract_tables()
            for table in tables:
                for row in table:
                    cells = [str(c).strip() if c else "" for c in row]
                    joined = " ".join(cells).lower()
                    if "raison sociale" in joined:
                        info.setdefault("raison_sociale", self._last_val(row))
                    elif "taxe professionnelle" in joined:
                        info.setdefault("taxe_professionnelle", self._last_val(row))
                    elif "identifiant fiscal" in joined:
                        info.setdefault("identifiant_fiscal", self._last_val(row))
                    elif "adresse" in joined:
                        info.setdefault("adresse", self._last_val(row))
        # Exercice depuis texte
        for i in range(min(5, self.n)):
            t = self.pages[i].extract_text() or ""
            m = re.search(r"(\d{2}/\d{2}/\d{4})\s+au\s+(\d{2}/\d{2}/\d{4})", t)
            if m:
                info["exercice"]       = f"Du {m.group(1)} au {m.group(2)}"
                info["exercice_debut"] = m.group(1)
                info["exercice_fin"]   = m.group(2)
                break
        info.setdefault("raison_sociale", "")
        info.setdefault("exercice", "")
        info.setdefault("exercice_fin", "")
        return info

    def _last_val(self, row) -> str:
        cells = [str(c).strip() for c in row if c and str(c).strip()]
        for c in reversed(cells):
            if len(c) > 2 and not any(k in c.lower() for k in
                    ["raison", "taxe", "identifiant", "adresse", ":"]):
                return c
        return cells[-1] if cells else ""

    def _extract_section(self, page_indices: list, mode: str) -> list:
        """
        Retourne une liste de dicts :
          actif  : {ref, label, brut, amort, net_n, net_n1, type}
          passif : {ref, label, val_n, val_n1, type}
          cpc    : {num, label, propre_n, prec_n, total_n, total_n1, type}
        """
        rows = []
        seen = set()

        for idx in page_indices:
            if idx >= self.n:
                continue
            tables = self.pages[idx].extract_tables()
            for table in tables:
                for row in table:
                    if not row:
                        continue
                    parsed = self._parse_row(row, mode)
                    if parsed is None:
                        continue
                    label = parsed.get("label", "")
                    if not label or _should_skip(label):
                        continue
                    key = _norm(label)
                    if key in seen:
                        continue
                    seen.add(key)
                    parsed["type"] = _row_type(label)
                    rows.append(parsed)

        return rows

    def _parse_row(self, row: list, mode: str) -> dict | None:
        row = [c for c in row]  # keep None
        n = len(row)

        if mode == "actif":
            # Structure : [ref?, label, brut, ?, amort, net_n, net_n1]
            if n >= 7:
                ref = str(row[0]).strip() if row[0] else ""
                label = _clean_label(row[1])
                brut   = _parse_num(row[2])
                amort  = _parse_num(row[4])
                net_n  = _parse_num(row[5])
                net_n1 = _parse_num(row[6])
            elif n == 6:
                ref = str(row[0]).strip() if row[0] else ""
                label = _clean_label(row[1])
                brut   = _parse_num(row[2])
                amort  = _parse_num(row[3])
                net_n  = _parse_num(row[4])
                net_n1 = _parse_num(row[5])
            elif n == 5:
                ref = ""
                label = _clean_label(row[0])
                brut   = _parse_num(row[1])
                amort  = _parse_num(row[2])
                net_n  = _parse_num(row[3])
                net_n1 = _parse_num(row[4])
            elif n >= 3:
                ref = ""
                label = _clean_label(row[0])
                brut   = _parse_num(row[1]) if n > 1 else None
                amort  = None
                net_n  = _parse_num(row[2]) if n > 2 else None
                net_n1 = _parse_num(row[3]) if n > 3 else None
            else:
                return None
            if not label:
                return None
            return {"ref": ref, "label": label,
                    "brut": brut, "amort": amort,
                    "net_n": net_n, "net_n1": net_n1}

        elif mode == "passif":
            if n >= 5:
                ref    = str(row[0]).strip() if row[0] else ""
                label  = _clean_label(row[1])
                val_n  = _parse_num(row[3])
                val_n1 = _parse_num(row[4])
            elif n == 4:
                ref    = str(row[0]).strip() if row[0] else ""
                label  = _clean_label(row[1])
                val_n  = _parse_num(row[2])
                val_n1 = _parse_num(row[3])
            elif n == 3:
                ref    = ""
                label  = _clean_label(row[0])
                val_n  = _parse_num(row[1])
                val_n1 = _parse_num(row[2])
            else:
                return None
            if not label:
                return None
            return {"ref": ref, "label": label, "val_n": val_n, "val_n1": val_n1}

        elif mode == "cpc":
            if n >= 8:
                num      = str(row[1]).strip() if row[1] else ""
                label    = _clean_label(row[2])
                propre_n = _parse_num(row[3])
                prec_n   = _parse_num(row[5])
                total_n  = _parse_num(row[6])
                total_n1 = _parse_num(row[7])
            elif n == 7:
                num      = str(row[1]).strip() if row[1] else ""
                label    = _clean_label(row[2])
                propre_n = _parse_num(row[3])
                prec_n   = _parse_num(row[4])
                total_n  = _parse_num(row[5])
                total_n1 = _parse_num(row[6])
            elif n == 6:
                num      = str(row[0]).strip() if row[0] else ""
                label    = _clean_label(row[1])
                propre_n = _parse_num(row[2])
                prec_n   = _parse_num(row[3])
                total_n  = _parse_num(row[4])
                total_n1 = _parse_num(row[5])
            elif n == 5:
                num      = ""
                label    = _clean_label(row[0])
                propre_n = _parse_num(row[1])
                prec_n   = _parse_num(row[2])
                total_n  = _parse_num(row[3])
                total_n1 = _parse_num(row[4])
            else:
                return None
            if not label:
                return None
            return {"num": num, "label": label,
                    "propre_n": propre_n, "prec_n": prec_n,
                    "total_n": total_n, "total_n1": total_n1}

        return None


# ═══════════════════════════════════════════════════════════════════════════════
# FORMATAGE EXCEL
# ═══════════════════════════════════════════════════════════════════════════════

def _border(sides="all"):
    s = Side(style='thin', color=C_BORDER)
    n = Side(style=None)
    t = s if sides == "all" or 't' in sides else n
    b = s if sides == "all" or 'b' in sides else n
    l = s if sides == "all" or 'l' in sides else n
    r = s if sides == "all" or 'r' in sides else n
    return Border(top=t, bottom=b, left=l, right=r)

def _fills(typ):
    """Retourne (bg_hex, fg_hex, bold) selon le type de ligne."""
    if typ == "total":
        return C_DARK_BLUE, C_WHITE, True
    if typ == "result":
        return C_RESULT, C_WHITE, True
    if typ == "section":
        return C_SECTION, C_DARK_BLUE, True
    if typ == "subtotal":
        return C_SUBTOTAL, C_DARK_BLUE, False
    return C_WHITE, "222222", False

def _set(ws, row, col, value=None, bg=C_WHITE, fg="222222", bold=False,
         align="left", num_fmt=None, indent=0, wrap=True, size=9):
    cell = ws.cell(row=row, column=col, value=value)
    cell.font      = Font(name="Arial", size=size, bold=bold, color=fg)
    cell.fill      = PatternFill("solid", fgColor=bg)
    cell.alignment = Alignment(horizontal=align, vertical="center",
                               indent=indent, wrap_text=wrap)
    cell.border    = _border()
    if num_fmt:
        cell.number_format = num_fmt
    return cell

def _header_row(ws, r, labels_cols: list, height=28):
    """labels_cols : [(col_start, col_end, text, bg)]"""
    ws.row_dimensions[r].height = height
    for col_s, col_e, text, bg in labels_cols:
        if col_s != col_e:
            ws.merge_cells(start_row=r, start_column=col_s,
                           end_row=r, end_column=col_e)
        _set(ws, r, col_s, text, bg=bg, fg=C_WHITE, bold=True,
             align="center", size=10, wrap=True)

def _cover(ws, r, raison, exercice, if_num, title, n_cols):
    """Bloc titre + infos entreprise. Retourne la prochaine ligne."""
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=n_cols)
    _set(ws, r, 1, title, bg=C_DARK_BLUE, fg=C_WHITE, bold=True,
         align="center", size=12)
    ws.row_dimensions[r].height = 26
    r += 1

    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=n_cols-2 or 1)
    _set(ws, r, 1, f"Raison sociale : {raison}",
         bg=C_LIGHT_BLUE, fg=C_DARK_BLUE, bold=True, size=9, indent=1)
    if n_cols >= 3:
        ws.merge_cells(start_row=r, start_column=n_cols-1,
                       end_row=r, end_column=n_cols)
        _set(ws, r, n_cols-1, f"IF : {if_num}",
             bg=C_LIGHT_BLUE, fg=C_DARK_BLUE, align="right", size=9)
    ws.row_dimensions[r].height = 16
    r += 1

    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=n_cols)
    _set(ws, r, 1, f"Exercice : {exercice}",
         bg=C_GRAY_BG, fg="555555", size=9, indent=1)
    ws.row_dimensions[r].height = 14
    r += 1

    # Ligne séparatrice vide
    for c in range(1, n_cols+1):
        _set(ws, r, c, bg=C_WHITE)
    ws.row_dimensions[r].height = 4
    r += 1
    return r

def _data_row(ws, r, row_type, cells_data):
    """
    cells_data : [(col, value, align, is_num)]
    """
    bg, fg, bold = _fills(row_type)
    h = 15 if row_type == "normal" else 17
    ws.row_dimensions[r].height = h
    for col, value, align, is_num in cells_data:
        _set(ws, r, col, value, bg=bg, fg=fg, bold=bold,
             align=align, num_fmt=NUM_FMT if is_num else None,
             indent=1 if (align == "left" and row_type == "normal") else 0)


# ═══════════════════════════════════════════════════════════════════════════════
# CONSTRUCTION DES FEUILLES
# ═══════════════════════════════════════════════════════════════════════════════

def _sheet_identification(wb, info):
    ws = wb.create_sheet("1 — Identification")
    ws.sheet_view.showGridLines = False
    ws.column_dimensions['A'].width = 30
    ws.column_dimensions['B'].width = 46

    ws.merge_cells('A1:B1')
    _set(ws, 1, 1, "PIÈCES ANNEXES À LA DÉCLARATION FISCALE",
         bg=C_DARK_BLUE, fg=C_WHITE, bold=True, align="center", size=13)
    ws.row_dimensions[1].height = 28

    ws.merge_cells('A2:B2')
    _set(ws, 2, 1, "IMPÔTS SUR LES SOCIÉTÉS — Modèle Comptable Normal (loi 9-88)",
         bg=C_MED_BLUE, fg=C_WHITE, align="center", size=10)
    ws.row_dimensions[2].height = 18

    fields = [
        ("Raison sociale",       info.get("raison_sociale", "—")),
        ("Identifiant fiscal",   info.get("identifiant_fiscal", "—")),
        ("Taxe professionnelle", info.get("taxe_professionnelle", "—")),
        ("Adresse",              info.get("adresse", "—")),
        ("Exercice",             info.get("exercice", "—")),
    ]
    for i, (lbl, val) in enumerate(fields, 4):
        ws.row_dimensions[i].height = 18
        _set(ws, i, 1, lbl, bg=C_LIGHT_BLUE, fg=C_DARK_BLUE, bold=True,
             size=9, indent=1)
        _set(ws, i, 2, val, bg=C_WHITE, fg="222222", size=9, indent=1)


def _sheet_actif(wb, info, rows):
    ws = wb.create_sheet("2 — Bilan Actif")
    ws.sheet_view.showGridLines = False
    ws.column_dimensions['A'].width = 4
    ws.column_dimensions['B'].width = 46
    ws.column_dimensions['C'].width = 18
    ws.column_dimensions['D'].width = 18
    ws.column_dimensions['E'].width = 18
    ws.column_dimensions['F'].width = 18

    r = _cover(ws, 1, info.get("raison_sociale", ""), info.get("exercice", ""),
               info.get("identifiant_fiscal", ""), "BILAN ACTIF — Modèle Comptable Normal", 6)

    _header_row(ws, r, [
        (1, 2, "ACTIF",             C_DARK_BLUE),
        (3, 3, "BRUT",              C_MED_BLUE),
        (4, 4, "AMORT. & PROV.",    C_MED_BLUE),
        (5, 5, "NET — EXERCICE N",  C_DARK_BLUE),
        (6, 6, "NET — EXERCICE N-1",C_MED_BLUE),
    ])
    ws.freeze_panes = f"A{r+1}"
    r += 1

    total_rows = 0
    for d in rows:
        typ = d.get("type", "normal")
        ref = d.get("ref") or ""
        # Cellule référence
        _data_row(ws, r, typ, [
            (1, ref if ref else None, "center", False),
            (2, d["label"],  "left",  False),
            (3, d.get("brut"),  "right", True),
            (4, d.get("amort"), "right", True),
            (5, d.get("net_n"), "right", True),
            (6, d.get("net_n1"),"right", True),
        ])
        # Ref cell styling overrides
        bg, fg, bold = _fills(typ)
        c_ref = ws.cell(row=r, column=1)
        c_ref.font = Font(name="Arial", size=8, bold=True,
                          color=C_MED_BLUE if typ not in ("total","result") else C_WHITE)
        c_ref.fill = PatternFill("solid", fgColor=bg)
        r += 1
        total_rows += 1

    return total_rows


def _sheet_passif(wb, info, rows):
    ws = wb.create_sheet("3 — Bilan Passif")
    ws.sheet_view.showGridLines = False
    ws.column_dimensions['A'].width = 4
    ws.column_dimensions['B'].width = 48
    ws.column_dimensions['C'].width = 20
    ws.column_dimensions['D'].width = 20

    r = _cover(ws, 1, info.get("raison_sociale", ""), info.get("exercice", ""),
               info.get("identifiant_fiscal", ""), "BILAN PASSIF — Modèle Comptable Normal", 4)

    _header_row(ws, r, [
        (1, 2, "PASSIF",         C_DARK_BLUE),
        (3, 3, "EXERCICE N",     C_DARK_BLUE),
        (4, 4, "EXERCICE N-1",   C_MED_BLUE),
    ])
    ws.freeze_panes = f"A{r+1}"
    r += 1

    total_rows = 0
    for d in rows:
        typ = d.get("type", "normal")
        ref = d.get("ref") or ""
        _data_row(ws, r, typ, [
            (1, ref if ref else None, "center", False),
            (2, d["label"],  "left",  False),
            (3, d.get("val_n"),  "right", True),
            (4, d.get("val_n1"), "right", True),
        ])
        bg, fg, bold = _fills(typ)
        c_ref = ws.cell(row=r, column=1)
        c_ref.font = Font(name="Arial", size=8, bold=True,
                          color=C_MED_BLUE if typ not in ("total","result") else C_WHITE)
        c_ref.fill = PatternFill("solid", fgColor=bg)
        r += 1
        total_rows += 1

    # Note légale
    r += 1
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=4)
    c = ws.cell(row=r, column=1,
                value="(1) Capital personnel débiteur.  (2) Bénéficiaire (+) / Déficitaire (−).")
    c.font = Font(name="Arial", italic=True, size=8, color="888888")
    c.alignment = Alignment(horizontal="left", vertical="center", indent=1)

    return total_rows


def _sheet_cpc(wb, info, rows):
    ws = wb.create_sheet("4 — CPC")
    ws.sheet_view.showGridLines = False
    ws.column_dimensions['A'].width = 5
    ws.column_dimensions['B'].width = 48
    ws.column_dimensions['C'].width = 18
    ws.column_dimensions['D'].width = 18
    ws.column_dimensions['E'].width = 18
    ws.column_dimensions['F'].width = 18

    r = _cover(ws, 1, info.get("raison_sociale", ""), info.get("exercice", ""),
               info.get("identifiant_fiscal", ""),
               "COMPTE DE PRODUITS ET CHARGES (Hors Taxes)", 6)

    _header_row(ws, r, [
        (1, 2, "DÉSIGNATION",           C_DARK_BLUE),
        (3, 3, "PROPRES À\nL'EXERCICE", C_MED_BLUE),
        (4, 4, "EXERCICES\nPRÉCÉDENTS", C_MED_BLUE),
        (5, 5, "TOTAUX\nEXERCICE N",    C_DARK_BLUE),
        (6, 6, "TOTAUX\nEXERCICE N-1",  C_MED_BLUE),
    ])
    ws.freeze_panes = f"A{r+1}"
    r += 1

    total_rows = 0
    for d in rows:
        typ = d.get("type", "normal")
        num = d.get("num") or ""
        _data_row(ws, r, typ, [
            (1, num if num else None, "center", False),
            (2, d["label"],      "left",  False),
            (3, d.get("propre_n"), "right", True),
            (4, d.get("prec_n"),   "right", True),
            (5, d.get("total_n"),  "right", True),
            (6, d.get("total_n1"), "right", True),
        ])
        bg, fg, bold = _fills(typ)
        c_num = ws.cell(row=r, column=1)
        c_num.font = Font(name="Arial", size=8, bold=True,
                          color=C_MED_BLUE if typ not in ("total","result") else C_WHITE)
        c_num.fill = PatternFill("solid", fgColor=bg)
        r += 1
        total_rows += 1

    r += 1
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=6)
    c = ws.cell(row=r, column=1,
                value="(1) Stock final − Stock initial : Augmentation (+) / Diminution (−).   "
                      "(2) Achats revendus = Achats − Variation de stock.")
    c.font = Font(name="Arial", italic=True, size=8, color="888888")
    c.alignment = Alignment(horizontal="left", vertical="center", indent=1)

    return total_rows


# ═══════════════════════════════════════════════════════════════════════════════
# POINT D'ENTRÉE PUBLIC
# ═══════════════════════════════════════════════════════════════════════════════

def convert(pdf_path: str, output_path: str) -> dict:
    """
    Extrait le PDF fiscal et génère un Excel structuré et formaté.

    Retourne un dict :
      {info, tables, rows, pages}
    compatible avec app.py (stats['tables'], stats['rows'], stats['pages']).
    """
    # 1. Extraction
    extractor = PDFExtractor(pdf_path)
    data = extractor.extract()

    info   = data["info"]
    actif  = data["actif"]
    passif = data["passif"]
    cpc    = data["cpc"]

    # 2. Construction du workbook
    wb = Workbook()
    wb.remove(wb.active)

    _sheet_identification(wb, info)
    n_actif  = _sheet_actif(wb, info, actif)
    n_passif = _sheet_passif(wb, info, passif)
    n_cpc    = _sheet_cpc(wb, info, cpc)

    wb.save(output_path)

    total_rows   = n_actif + n_passif + n_cpc
    total_tables = sum([
        1 if actif  else 0,
        1 if passif else 0,
        1 if cpc    else 0,
    ])

    return {
        "info":   info,
        "tables": total_tables,
        "rows":   total_rows,
        "pages":  data["pages"],
    }
