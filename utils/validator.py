"""utils/validator.py — v3 (robuste, sans accents)"""
import unicodedata
from utils.logger import get_logger
logger = get_logger(__name__)


def _normalize(text: str) -> str:
    """Enlève les accents et met en minuscules pour comparaison robuste."""
    text = unicodedata.normalize('NFD', text)
    return text.encode('ascii', 'ignore').decode('utf-8').lower()


def validate_pdf_structure_v2(parser) -> dict:
    if parser.n_pages < 2:
        return {
            "valid": False,
            "message": f"PDF trop court ({parser.n_pages} page). Minimum 2 pages.",
            "meta": {}
        }

    # Normaliser tout le texte des premières pages (sans accents)
    raw_text = " ".join(parser._page_text(i) for i in range(min(6, parser.n_pages)))
    full_low = _normalize(raw_text)

    missing = []

    # Actif : plusieurs variantes possibles
    if not any(k in full_low for k in [
        "immobilisations", "actif immobilise", "actif immobilisé",
        "bilan (actif", "bilan actif", "frais preliminaires",
    ]):
        missing.append("Bilan Actif")

    # Passif : variantes
    if not any(k in full_low for k in [
        "capitaux propres", "passif", "capital social",
        "bilan (passif", "bilan passif",
    ]):
        missing.append("Bilan Passif")

    # CPC : variantes
    if not any(k in full_low for k in [
        "produits exploitation", "charges exploitation",
        "produits d exploitation", "charges d exploitation",
        "compte de produits", "ventes de marchandises",
        "chiffre d affaires", "chiffres d affaires", "cpc",
    ]):
        missing.append("CPC")

    if missing:
        return {
            "valid": False,
            "message": f"Sections manquantes ou non détectées : {', '.join(missing)}. "
                       f"Vérifiez que le PDF contient bien le Bilan et le CPC au format MCN.",
            "meta": {}
        }

    meta = parser._parse_info()
    meta["pages"] = parser.n_pages
    return {"valid": True, "message": "OK", "meta": meta}
