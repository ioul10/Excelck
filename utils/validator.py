"""utils/validator.py — v3 — supporte DGI 7 pages et AMMC 5 pages"""
from utils.logger import get_logger
logger = get_logger(__name__)


def validate_pdf_structure_v2(parser, mode: str = "auto") -> dict:
    """
    Valide la structure du PDF.
    mode: "dgi" | "ammc" | "auto" (détecte automatiquement)
    """
    if parser.n_pages < 2:
        return {
            "valid": False,
            "message": f"PDF trop court ({parser.n_pages} page). Minimum 4 pages.",
            "meta": {}
        }

    # Scanner TOUTES les pages pour les mots-clés
    full_low = " ".join(
        parser._page_text(i) for i in range(parser.n_pages)
    ).lower()

    missing = []
    if not any(k in full_low for k in ["immobilisations", "actif immobilisé", "bilan (actif)"]):
        missing.append("Bilan Actif")
    if not any(k in full_low for k in ["capitaux propres", "passif", "bilan (passif)"]):
        missing.append("Bilan Passif")
    if not any(k in full_low for k in [
        "produits", "charges", "exploitation",
        "compte de produits", "ventes de"
    ]):
        missing.append("CPC")

    if missing:
        return {
            "valid": False,
            "message": f"Sections manquantes : {', '.join(missing)}.",
            "meta": {}
        }

    meta = parser._parse_info()
    meta["pages"] = parser.n_pages

    # Détection automatique du format si mode=auto
    if mode == "auto":
        if parser.n_pages == 7:
            meta["format_detected"] = "dgi"
        elif parser.n_pages in (5, 6):
            meta["format_detected"] = "ammc"
        else:
            meta["format_detected"] = "ammc"
    else:
        meta["format_detected"] = mode

    return {"valid": True, "message": "OK", "meta": meta}
