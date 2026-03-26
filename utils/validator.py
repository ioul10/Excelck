"""utils/validator.py — v2"""
from utils.logger import get_logger
logger = get_logger(__name__)


def validate_pdf_structure_v2(parser) -> dict:
    if parser.n_pages < 2:
        return {"valid": False,
                "message": f"PDF trop court ({parser.n_pages} page). Minimum 4 pages.",
                "meta": {}}

    full_low = " ".join(parser._page_text(i) for i in range(min(4, parser.n_pages))).lower()

    missing = []
    if not any(k in full_low for k in ["immobilisations", "actif immobilisé"]):
        missing.append("Bilan Actif")
    if not any(k in full_low for k in ["capitaux propres", "passif"]):
        missing.append("Bilan Passif")
    if not any(k in full_low for k in ["produits d'exploitation", "charges d'exploitation"]):
        missing.append("CPC")

    if missing:
        return {"valid": False,
                "message": f"Sections manquantes : {', '.join(missing)}.",
                "meta": {}}

    meta = parser._parse_info()
    meta["pages"] = parser.n_pages
    return {"valid": True, "message": "OK", "meta": meta}
