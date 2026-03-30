"""
core/dgi_parser.py — Parseur PDF DGI (7 pages)
Structure DGI :
  Page 1    : Identification
  Pages 2-3 : Bilan Actif  (tableau 6 cols)
  Page 4    : Bilan Passif (tableau 4 cols)
  Pages 5-7 : CPC          (tableau 6 cols)
"""

import re, pdfplumber
from utils.logger import get_logger

logger = get_logger(__name__)


def _parse_fr(s) -> float | None:
    if not s: return None
    s = str(s).strip().replace(' ', '').replace('\xa0', '')
    if not s or s in ['-', '—', '']: return None
    neg = s.startswith('-'); s = s.lstrip('-')
    m = re.match(r'^(\d{1,3}(?:\.\d{3})*),(\d{2})$', s)
    if m: s = m.group(1).replace('.', '') + '.' + m.group(2)
    elif re.match(r'^\d+,\d{2}$', s): s = s.replace(',', '.')
    elif re.match(r'^\d+$', s): pass
    else: return None
    try: return -float(s) if neg else float(s)
    except: return None


def _clean_label(s) -> str:
    if not s: return ''
    s = str(s).replace('\n', ' ').strip()
    s = re.sub(r'^\*\s*', '', s)       # retirer *
    s = re.sub(r'^\-\s*', '', s)       # retirer -
    s = re.sub(r'\(\d+\)$', '', s)     # retirer (1)(2) en fin
    s = re.sub(r'\s+', ' ', s)
    return s.strip()


def _extract_info(pdf) -> dict:
    info = {}
    for i in range(min(2, len(pdf.pages))):
        text = pdf.pages[i].extract_text() or ''
        for key, pat in [
            ('raison_sociale',      r'[Rr]aison\s+[Ss]ociale\s*[:\-]?\s*([A-Z][^\n]{3,60})'),
            ('identifiant_fiscal',  r'[Ii]dentifiant\s+[Ff]iscal\s*[:\-]?\s*(\d+)'),
            ('taxe_professionnelle',r'[Tt]axe\s+[Pp]rof\w*\.?\s*[:\-]?\s*([\d\s]+)'),
            ('adresse',             r'[Aa]dresse\s*[:\-]?\s*([^\n]{5,60})'),
            ('article_is',          r'[Aa]rticle\s+[Ii][Ss]\s*[:\-]?\s*(\d+)'),
        ]:
            if key not in info:
                m = re.search(pat, text, re.IGNORECASE)
                if m: info[key] = m.group(1).strip()

        if 'exercice' not in info:
            m = re.search(r'p[eé]riode\s+du\s+(\d{2}/\d{2}/\d{4})\s+au\s+(\d{2}/\d{2}/\d{4})', text, re.I)
            if not m:
                m = re.search(r'(\d{2}/\d{2}/\d{4})\s+au\s+(\d{2}/\d{2}/\d{4})', text)
            if m:
                info['exercice']       = f"Du {m.group(1)} au {m.group(2)}"
                info['exercice_debut'] = m.group(1)
                info['exercice_fin']   = m.group(2)

        # Date de dépôt
        if 'date_declaration' not in info:
            m = re.search(r'le\s+(\d{2}/\d{2}/\d{4})', text, re.I)
            if m: info['date_declaration'] = m.group(1)

    for k in ('raison_sociale', 'identifiant_fiscal', 'taxe_professionnelle',
              'adresse', 'exercice', 'exercice_fin', 'article_is'):
        info.setdefault(k, '')
    info['pages'] = len(pdf.pages)
    return info


def _extract_actif(pdf) -> dict:
    """Pages 2-3 (index 1-2) → actif {label: [brut, amort, net_n1]}"""
    result = {}
    for pg_idx in [1, 2]:
        if pg_idx >= len(pdf.pages): continue
        tables = pdf.pages[pg_idx].extract_tables()
        t = next((t for t in tables if len(t) > 3 and len(t[0]) >= 5), None)
        if not t: continue
        for row in t[1:]:  # skip header
            if len(row) < 4: continue
            label = _clean_label(row[1])
            if not label or len(label) < 3: continue
            brut   = _parse_fr(row[2]) if len(row) > 2 else None
            amort  = _parse_fr(row[3]) if len(row) > 3 else None
            net_n1 = _parse_fr(row[5]) if len(row) > 5 else None
            if any(v is not None for v in [brut, amort, net_n1]):
                result[label] = [brut, amort, net_n1]
    logger.info(f"DGI actif: {len(result)} postes")
    return result


def _extract_passif(pdf) -> dict:
    """Page 4 (index 3) → passif {label: [val_n, val_n1]}"""
    result = {}
    if 3 >= len(pdf.pages): return result
    tables = pdf.pages[3].extract_tables()
    t = next((t for t in tables if len(t) > 3 and len(t[0]) >= 3), None)
    if not t: return result
    for row in t[1:]:
        if len(row) < 3: continue
        label  = _clean_label(row[1] if len(row) > 1 else row[0])
        if not label or len(label) < 3: continue
        val_n  = _parse_fr(row[2]) if len(row) > 2 else None
        val_n1 = _parse_fr(row[3]) if len(row) > 3 else None
        if any(v is not None for v in [val_n, val_n1]):
            result[label] = [val_n, val_n1]
    logger.info(f"DGI passif: {len(result)} postes")
    return result


def _extract_cpc(pdf) -> dict:
    """Pages 5-7 (index 4-6) → cpc {label: [propre_n, prec_n, total_n1]}"""
    result = {}
    for pg_idx in [4, 5, 6]:
        if pg_idx >= len(pdf.pages): continue
        tables = pdf.pages[pg_idx].extract_tables()
        t = next((t for t in tables if len(t) > 2 and len(t[0]) >= 4), None)
        if not t: continue
        for row in t:
            if len(row) < 3: continue
            label   = _clean_label(row[1] if len(row) > 1 else row[0])
            if not label or len(label) < 3: continue
            propre  = _parse_fr(row[2]) if len(row) > 2 else None
            prec    = _parse_fr(row[3]) if len(row) > 3 else None
            tot_n1  = _parse_fr(row[5]) if len(row) > 5 else None
            if any(v is not None for v in [propre, prec, tot_n1]):
                result[label] = [propre, prec, tot_n1]
    logger.info(f"DGI CPC: {len(result)} postes")
    return result


def parse(pdf_path: str) -> dict:
    """
    Parse un PDF DGI (7 pages).
    Retourne {'info', 'actif_values', 'passif_values', 'cpc_values'}.
    """
    pdf = pdfplumber.open(pdf_path)
    out = {
        'info':          _extract_info(pdf),
        'actif_values':  _extract_actif(pdf),
        'passif_values': _extract_passif(pdf),
        'cpc_values':    _extract_cpc(pdf),
    }
    pdf.close()
    return out
