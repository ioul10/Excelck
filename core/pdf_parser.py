"""
core/pdf_parser.py  v3 — dynamique
Extraction PDF → valeurs structurées via pdfplumber tables.
Détection automatique des pages par mots-clés (plus de numéros hardcodés).
"""

import re
import unicodedata
import pdfplumber
from utils.logger import get_logger

logger = get_logger(__name__)

# Labels à ignorer (en-têtes, totaux intermédiaires non désirés)
SKIP_LABELS = {
    "tableau n° 1(1/2)", "tableau n° 1(2/2)", "tableau n° 2(1/2)", "tableau n° 2(2/2)",
    "bilan (actif) (modèle normal)", "bilan (passif) (modèle normal)",
    "compte de produits et charges", "agence du bassin", "identifiant fiscal",
    "exercice du", "brut", "amortissements et provi", "a c t i f", "p a s s i f",
    "exercice", "exercice precedent", "net", "designation", "operations",
    "propres à", "concernant les", "1", "2", "3 = 2 + 1", "4",
    "(1)capital personnel", "(2)bénéficiaire", "1)variation de stock",
    "2)achats revendus", "totaux de", "totaux de l'exercice", "nb:",
}

SKIP_PREFIXES = (
    "tableau", "bilan (", "compte de produits", "agence du",
    "(1)", "(2)", "1)variation", "2)achats",
)

TOTAL_SKIP_EXACT = {
    "total i", "total ii", "total iii", "total général", "total general",
    "total des produits", "total des charges",
    "trésorerie-actif", "trésorerie passif", "tresorerie passif",
    "capitaux propres", "dettes du passif circulant",
}

TOTAL_SKIP_PREFIXES = (
    "résultat d'exploitation", "résultat financier", "résultat courant",
    "résultat non courant", "résultat avant impôts",
    "resultat d'exploitation", "resultat financier", "resultat courant",
    "resultat non courant", "resultat avant impots",
    "total i ", "total ii", "total iii", "total des ", "total général",
    "produits d'exploitation", "charges d'exploitation",
    "charges financières", "charges financieres",
    "produits non courants", "charges non courants",
)
TOTAL_SKIP = TOTAL_SKIP_EXACT


def _normalize_text(text: str) -> str:
    """Normalise le texte : minuscules, sans accents."""
    text = unicodedata.normalize('NFD', text)
    text = text.encode('ascii', 'ignore').decode('utf-8')
    return text.lower()


class PDFParser:

    def __init__(self, pdf_path: str):
        self.path    = pdf_path
        self.pdf     = pdfplumber.open(pdf_path)
        self.pages   = self.pdf.pages
        self.n_pages = len(self.pages)
        self._ranges = {}
        logger.info(f"PDF chargé : {pdf_path} — {self.n_pages} pages")

    def parse(self) -> dict:
        self._ranges = self._detect_page_ranges()
        logger.info(f"Pages détectées : {self._ranges}")
        result = {
            "info":          self._parse_info(),
            "actif_values":  self._parse_actif(),
            "passif_values": self._parse_passif(),
            "cpc_values":    self._parse_cpc(),
        }
        self._enrich_passif(result)
        return result

    # ── Détection dynamique des pages ────────────────────────────────────────

    def _detect_page_ranges(self) -> dict:
        """
        Détecte dynamiquement les pages de chaque section par mots-clés.
        Robuste aux variations de mise en page entre PDFs.
        """
        ranges = {"actif": [], "passif": [], "cpc": []}

        ACTIF_KEYWORDS = [
            "actif immobilise", "immobilisations incorporelles",
            "bilan (actif", "bilan actif", "actif immobilisé",
            "frais preliminaires", "frais préliminaires",
        ]
        PASSIF_KEYWORDS = [
            "capitaux propres", "bilan (passif", "bilan passif",
            "capital social", "passif circulant", "dettes de financement",
        ]
        CPC_KEYWORDS = [
            "produits exploitation", "charges exploitation",
            "compte de produits", "ventes de marchandises",
            "chiffre d affaires", "chiffres d affaires",
        ]

        for i, page in enumerate(self.pages):
            raw_text = page.extract_text() or ""
            text = _normalize_text(raw_text)

            if any(k in text for k in ACTIF_KEYWORDS):
                ranges["actif"].append(i)
            if any(k in text for k in PASSIF_KEYWORDS):
                ranges["passif"].append(i)
            if any(k in text for k in CPC_KEYWORDS):
                ranges["cpc"].append(i)

        # Fallback : si aucune page détectée, utiliser les plages classiques
        if not ranges["actif"]:
            logger.warning("Actif : détection échouée, fallback pages 1-2")
            ranges["actif"] = list(range(1, min(3, self.n_pages)))
        if not ranges["passif"]:
            logger.warning("Passif : détection échouée, fallback pages 2-4")
            ranges["passif"] = list(range(2, min(5, self.n_pages)))
        if not ranges["cpc"]:
            logger.warning("CPC : détection échouée, fallback pages 3-6")
            ranges["cpc"] = list(range(3, min(7, self.n_pages)))

        return ranges

    # ── Enrichissement passif ─────────────────────────────────────────────────

    def _enrich_passif(self, data: dict):
        pv = data["passif_values"]
        cpc = data["cpc_values"]

        rn_key = "Résultat net de l'exercice"
        if rn_key not in pv:
            for k, v in cpc.items():
                if "RESULTAT NET (XI-XII)" in k or "RESULTAT NET (XI" in k:
                    propre = v[0] or 0
                    prec   = v[1] or 0
                    total_n   = round(propre + prec, 2)
                    total_n1  = v[2]
                    pv[rn_key] = [total_n, total_n1]
                    logger.info(f"Résultat net passif inféré depuis CPC : N={total_n} N-1={total_n1}")
                    break

    # ── Infos générales ───────────────────────────────────────────────────────

    def _parse_info(self) -> dict:
        info = {}
        tables = self.pages[0].extract_tables()
        for table in tables:
            for row in table:
                cells = [str(c).strip() if c else "" for c in row]
                joined = " ".join(cells).lower()
                if "raison sociale" in joined:
                    info["raison_sociale"] = self._find_value_in_row(row)
                elif "taxe professionnelle" in joined:
                    info["taxe_professionnelle"] = self._find_value_in_row(row)
                elif "identifiant fiscal" in joined:
                    info["identifiant_fiscal"] = self._find_value_in_row(row)
                elif "adresse" in joined:
                    info["adresse"] = self._find_value_in_row(row)
                elif re.search(r"\d{2}/\d{2}/\d{4}", " ".join(cells)):
                    for c in cells:
                        if re.match(r"\d{2}/\d{2}/\d{4}", c.strip()):
                            info["date_declaration"] = c.strip()

        for i in range(1, min(4, self.n_pages)):
            t = self._page_text(i)
            m = re.search(r"(\d{2}/\d{2}/\d{4})\s+au\s+(\d{2}/\d{2}/\d{4})", t)
            if m:
                info["exercice"]       = f"Du {m.group(1)} au {m.group(2)}"
                info["exercice_debut"] = m.group(1)
                info["exercice_fin"]   = m.group(2)
                break
        info.setdefault("exercice", "")
        info.setdefault("exercice_fin", "")
        info["pages"] = self.n_pages
        return info

    def _find_value_in_row(self, row) -> str:
        cells = [str(c).strip() for c in row if c and str(c).strip()]
        for c in reversed(cells):
            if len(c) > 2 and not any(k in c.lower() for k in ["raison", "taxe", "identifiant", "adresse", ":"]):
                return c
        return cells[-1] if cells else ""

    # ── Bilan Actif ───────────────────────────────────────────────────────────

    def _parse_actif(self) -> dict:
        """
        Colonnes attendues (modèle normal) :
          7+ cols : [latéral, label, brut, vide, amort, net_n, net_n1]
          5 cols  : [latéral, label, brut, amort, net_n1]
          4 cols  : [label, brut, amort, net_n1]
        """
        values = {}
        page_indices = self._ranges.get("actif", [])

        for page_idx in page_indices:
            if page_idx >= self.n_pages:
                continue
            tables = self.pages[page_idx].extract_tables()
            for table in tables:
                for row in table:
                    if not row or len(row) < 3:
                        continue

                    # Déterminer position du label selon nb de colonnes
                    n = len(row)
                    label_col = 1 if n >= 5 else 0
                    label = self._clean_label(row[label_col])
                    if not label or self._should_skip(label):
                        continue

                    # Adapter les indices de colonnes selon n
                    if n >= 7:
                        brut   = self._parse_num(row[2])
                        amort  = self._parse_num(row[4])
                        net_n1 = self._parse_num(row[6])
                    elif n == 6:
                        brut   = self._parse_num(row[2])
                        amort  = self._parse_num(row[3])
                        net_n1 = self._parse_num(row[5])
                    elif n == 5:
                        brut   = self._parse_num(row[2])
                        amort  = self._parse_num(row[3])
                        net_n1 = self._parse_num(row[4])
                    elif n == 4:
                        brut   = self._parse_num(row[1])
                        amort  = self._parse_num(row[2])
                        net_n1 = self._parse_num(row[3])
                    else:
                        continue

                    if any(v is not None for v in [brut, amort, net_n1]):
                        if label not in values:
                            values[label] = [brut, amort, net_n1]

        logger.info(f"Actif : {len(values)} postes (pages {page_indices})")
        return values

    # ── Bilan Passif ──────────────────────────────────────────────────────────

    def _parse_passif(self) -> dict:
        """
        Colonnes attendues :
          5+ cols : [latéral, label, vide, val_n, val_n1]
          4 cols  : [latéral, label, val_n, val_n1]
          3 cols  : [label, val_n, val_n1]
        """
        values = {}
        page_indices = self._ranges.get("passif", [])

        for page_idx in page_indices:
            if page_idx >= self.n_pages:
                continue
            tables = self.pages[page_idx].extract_tables()
            for table in tables:
                for row in table:
                    if not row or len(row) < 3:
                        continue

                    n = len(row)
                    label_col = 1 if n >= 4 else 0
                    label = self._clean_label(row[label_col])
                    if not label or self._should_skip(label):
                        continue

                    if n >= 5:
                        val_n  = self._parse_num(row[3])
                        val_n1 = self._parse_num(row[4])
                    elif n == 4:
                        val_n  = self._parse_num(row[2])
                        val_n1 = self._parse_num(row[3])
                    elif n == 3:
                        val_n  = self._parse_num(row[1])
                        val_n1 = self._parse_num(row[2])
                    else:
                        continue

                    if any(v is not None for v in [val_n, val_n1]):
                        if label not in values:
                            values[label] = [val_n, val_n1]

        logger.info(f"Passif : {len(values)} postes (pages {page_indices})")
        return values

    # ── CPC ───────────────────────────────────────────────────────────────────

    def _parse_cpc(self) -> dict:
        """
        Structure des tableaux CPC selon le nombre de colonnes :
          7 cols: [lat, num, label, propre_n, prec_n, total_n, total_n1]
          8 cols: [lat, num, label, propre_n, VIDE, prec_n, total_n, total_n1]
          6 cols: [num, label, propre_n, prec_n, total_n, total_n1]
          5 cols: [label, propre_n, prec_n, total_n, total_n1]
        """
        values = {}
        page_indices = self._ranges.get("cpc", [])

        for page_idx in page_indices:
            if page_idx >= self.n_pages:
                continue
            tables = self.pages[page_idx].extract_tables()
            for table in tables:
                for row in table:
                    if not row or len(row) < 4:
                        continue

                    n = len(row)

                    # Trouver la colonne label selon n
                    if n >= 7:
                        label_col = 2
                    elif n == 6:
                        label_col = 1
                    else:
                        label_col = 0

                    label = self._clean_label(row[label_col] if len(row) > label_col else None)
                    if not label or self._should_skip(label):
                        continue

                    if n >= 8:
                        propre_n = self._parse_num(row[3])
                        prec_n   = self._parse_num(row[5])
                        total_n1 = self._parse_num(row[7])
                    elif n == 7:
                        propre_n = self._parse_num(row[3])
                        prec_n   = self._parse_num(row[4])
                        total_n1 = self._parse_num(row[6])
                    elif n == 6:
                        propre_n = self._parse_num(row[2])
                        prec_n   = self._parse_num(row[3])
                        total_n1 = self._parse_num(row[5])
                    elif n == 5:
                        propre_n = self._parse_num(row[1])
                        prec_n   = self._parse_num(row[2])
                        total_n1 = self._parse_num(row[4])
                    else:
                        continue

                    if any(v is not None for v in [propre_n, prec_n, total_n1]):
                        if label not in values:
                            values[label] = [propre_n, prec_n, total_n1]
                        else:
                            existing = values[label]
                            if existing[1] is None and prec_n is not None:
                                values[label] = [
                                    existing[0] if existing[0] is not None else propre_n,
                                    prec_n,
                                    existing[2] if existing[2] is not None else total_n1,
                                ]

        logger.info(f"CPC : {len(values)} postes (pages {page_indices})")
        return values

    # ── Utilitaires ───────────────────────────────────────────────────────────

    def _should_skip(self, label: str) -> bool:
        l = label.lower()
        if l in SKIP_LABELS:
            return True
        if any(l.startswith(p) for p in SKIP_PREFIXES):
            return True
        if l in TOTAL_SKIP_EXACT:
            return True
        if any(l.startswith(p) for p in TOTAL_SKIP_PREFIXES):
            return True
        if len(label) < 3:
            return True
        if re.match(r"^\d+$", label):
            return True
        return False

    @staticmethod
    def _clean_label(s) -> str:
        if not s:
            return ""
        s = str(s).replace("\n", " ").strip()
        s = re.sub(r"\s{2,}", " ", s)
        s = re.sub(r"^(I{1,3}|IV|V|VI{1,3}|IX|X{1,2})\s+", "", s)
        return s.strip()

    @staticmethod
    def _parse_num(s) -> float | None:
        if s is None:
            return None
        s = str(s).strip().replace("\n", "")
        if not s or s in ["-", "—", "", "None"]:
            return None
        neg = False
        if s.startswith("(") and s.endswith(")"):
            neg = True
            s = s[1:-1]
        if s.startswith("-"):
            neg = True
            s = s[1:]
        s = s.replace(" ", "").replace("\xa0", "").replace(",", ".")
        try:
            v = float(s)
            return -v if neg else v
        except ValueError:
            return None

    def _page_text(self, idx: int) -> str:
        if idx >= self.n_pages:
            return ""
        return self.pages[idx].extract_text() or ""

    def __del__(self):
        try:
            self.pdf.close()
        except Exception:
            pass
