"""
core/transformer.py
Étape 3 : Transformation des données brutes → structure canonique fiscale
         + Reconstruction des relations entre postes (totaux, sous-totaux)
"""

import re
from utils.logger import get_logger
from core.fiscal_schema import (
    BILAN_ACTIF_SCHEMA,
    BILAN_PASSIF_SCHEMA,
    CPC_SCHEMA,
)

logger = get_logger(__name__)


class FiscalTransformer:
    """
    Prend les données brutes de PDFExtractor et les aligne sur le schéma
    officiel du Modèle Comptable Normal.

    Stratégie de matching :
      1. Correspondance exacte (après normalisation)
      2. Correspondance par mots-clés (fuzzy léger)
      3. Position ordinale si le PDF est bien structuré
    """

    def __init__(self, raw_data: dict):
        self.raw = raw_data
        self.info         = raw_data.get("info", {})
        self.raw_actif    = raw_data.get("bilan_actif", [])
        self.raw_passif   = raw_data.get("bilan_passif", [])
        self.raw_cpc      = raw_data.get("cpc", [])

    def transform(self) -> dict:
        return {
            "info":         self.info,
            "bilan_actif":  self._align(self.raw_actif,  BILAN_ACTIF_SCHEMA,  cols=5),
            "bilan_passif": self._align(self.raw_passif, BILAN_PASSIF_SCHEMA, cols=3),
            "cpc":          self._align(self.raw_cpc,    CPC_SCHEMA,          cols=5),
        }

    # ── Alignement sur schéma ─────────────────────────────────────────────────

    def _align(self, raw_rows: list, schema: list, cols: int) -> list:
        """
        Aligne les lignes extraites sur le schéma officiel.
        - Si PDF bien parsé → matching par label
        - Sinon → remplissage positionnel

        Retourne une liste de tuples alignés sur le schéma.
        """
        if not raw_rows:
            logger.warning("Pas de données brutes — utilisation du schéma vide")
            return [row[:cols] for row in schema]

        # Construire index de recherche dans les données brutes
        raw_index = self._build_index(raw_rows)

        result = []
        for schema_row in schema:
            canonical_label = schema_row[0]
            default_values  = list(schema_row[1:cols])

            # Chercher dans les données extraites
            matched = self._find_match(canonical_label, raw_index, raw_rows)

            if matched:
                # Fusionner : prendre les valeurs extraites, compléter par défaut
                extracted_vals = list(matched[1:cols])
                merged = []
                for i in range(cols - 1):
                    ev = extracted_vals[i] if i < len(extracted_vals) else None
                    dv = default_values[i]  if i < len(default_values) else None
                    # Préférer la valeur extraite si disponible
                    merged.append(ev if ev is not None else dv)
                result.append((canonical_label, *merged))
            else:
                # Pas trouvé → garder les valeurs du schéma (fallback)
                result.append(tuple(schema_row[:cols]))

        return result

    def _build_index(self, rows: list) -> dict:
        """Index normalisé label → row pour accès O(1)."""
        index = {}
        for row in rows:
            if row and row[0]:
                key = self._normalize_key(row[0])
                index[key] = row
        return index

    def _find_match(self, canonical: str, index: dict, rows: list):
        """
        Cherche une correspondance dans l'index extrait.
        1. Exact normalisé
        2. Contient les mots-clés principaux
        """
        key = self._normalize_key(canonical)

        # 1. Correspondance exacte
        if key in index:
            return index[key]

        # 2. Correspondance partielle : mots-clés
        canon_words = set(self._keywords(canonical))
        best_match  = None
        best_score  = 0

        for raw_key, row in index.items():
            raw_words = set(self._keywords(raw_key))
            common    = canon_words & raw_words
            # Score = nb mots communs / max(len)
            if not canon_words:
                continue
            score = len(common) / max(len(canon_words), len(raw_words))
            if score > best_score and score >= 0.6:
                best_score = score
                best_match = row

        return best_match

    # ── Helpers ───────────────────────────────────────────────────────────────

    @staticmethod
    def _normalize_key(s: str) -> str:
        """Normalise un libellé pour comparaison."""
        s = s.lower().strip()
        # Supprimer accents simplifiés
        replacements = {
            "é": "e", "è": "e", "ê": "e", "à": "a", "â": "a",
            "ô": "o", "û": "u", "î": "i", "ç": "c", "œ": "oe",
        }
        for k, v in replacements.items():
            s = s.replace(k, v)
        # Supprimer ponctuation sauf lettres/chiffres/espaces
        s = re.sub(r"[^\w\s]", " ", s)
        s = re.sub(r"\s+", " ", s).strip()
        return s

    @staticmethod
    def _keywords(s: str) -> list:
        """Extrait les mots significatifs (> 3 chars) d'un libellé."""
        stop = {
            "les", "des", "de", "du", "et", "en", "sur", "par", "pour",
            "aux", "une", "un", "la", "le", "au", "ou", "est", "sont",
            "a", "b", "c", "d", "e", "f", "g", "h", "i"
        }
        words = re.findall(r"[a-záàâéèêîïôùûüçœ]{3,}", s.lower())
        return [w for w in words if w not in stop]
