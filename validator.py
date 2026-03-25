"""
utils/validator.py
Valide qu'un PDF est bien une pièce annexe IS Modèle Normal.
"""

import re
from utils.logger import get_logger

logger = get_logger(__name__)

REQUIRED_KEYWORDS = [
    # Page infos
    ["identifiant fiscal", "taxe professionnelle", "impôts"],
    # Bilan actif
    ["immobilisations", "actif", "stocks", "trésorerie"],
    # Bilan passif
    ["passif", "capitaux propres", "capital"],
    # CPC
    ["produits", "charges", "exploitation", "résultat"],
]


def validate_pdf_structure(extractor) -> dict:
    """
    Vérifie la structure du PDF.
    Retourne {"valid": bool, "message": str, "meta": dict}
    """
    if extractor.num_pages < 2:
        return {
            "valid":   False,
            "message": f"PDF trop court ({extractor.num_pages} page(s)). Minimum 4 pages requises.",
            "meta":    {},
        }

    full_text = extractor.get_all_text().lower()

    # Vérifier présence des sections clés
    missing = []
    checks = {
        "Bilan Actif":   ["immobilisations", "actif"],
        "Bilan Passif":  ["capitaux propres", "passif"],
        "CPC":           ["produits d'exploitation", "charges d'exploitation"],
    }
    for section, keywords in checks.items():
        if not any(k in full_text for k in keywords):
            missing.append(section)

    if missing:
        return {
            "valid":   False,
            "message": f"Sections manquantes : {', '.join(missing)}. "
                       f"Vérifiez que le PDF est bien une pièce annexe IS (Modèle Normal).",
            "meta":    {},
        }

    # Extraire métadonnées
    meta = extractor.extract_all().get("info", {})
    meta["pages"] = extractor.num_pages

    return {
        "valid":   True,
        "message": "Structure valide",
        "meta":    meta,
    }
