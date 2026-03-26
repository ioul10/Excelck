"""
core/pdf_parser.py  v2
Extraction PDF → valeurs structurées via pdfplumber tables.

Structure des colonnes (détectée empiriquement) :
  Bilan Actif  (page 2) : [latéral, label, brut, vide, amort, net_n, net_n1]  → idx 1,2,4,6
  Bilan Passif (page 3) : [latéral, label, vide, val_n, val_n1]               → idx 1,3,4
  CPC          (pages 4-5): [latéral, num, label, propre_n, prec_n, total_n, total_n1] → idx 2,3,4,5,6
"""

import re
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

# Labels de section/total à ne pas injecter (ils sont calculés par formule)
# Labels dont l'extraction CPC doit être évitée (lignes totales/résultats)
# ATTENTION : utiliser des patterns qui ne bloquent pas les sous-postes
TOTAL_SKIP_EXACT = {
    "total i", "total ii", "total iii", "total général", "total general",
    "total des produits", "total des charges",
    "trésorerie-actif", "trésorerie passif", "tresorerie passif",
    "capitaux propres", "dettes du passif circulant",
}
# Préfixes exacts qui indiquent un total/résultat (doit être en début de label)
TOTAL_SKIP_PREFIXES = (
    "résultat d'exploitation", "résultat financier", "résultat courant",
    "résultat non courant", "résultat avant impôts",
    "resultat d'exploitation", "resultat financier", "resultat courant",
    "resultat non courant", "resultat avant impots",
    # "resultat net" retiré intentionnellement : on veut extraire RESULTAT NET (XI-XII)
    # pour l'injecter dans Bilan Passif > Résultat net de l'exercice
    "total i ", "total ii", "total iii", "total des ", "total général",
    "produits d'exploitation", "charges d'exploitation",
    "charges financières", "charges financieres",
    "produits non courants", "charges non courants",
)
# Alias pour compatibilité
TOTAL_SKIP = TOTAL_SKIP_EXACT


class PDFParser:

    def __init__(self, pdf_path: str):
        self.path    = pdf_path
        self.pdf     = pdfplumber.open(pdf_path)
        self.pages   = self.pdf.pages
        self.n_pages = len(self.pages)
        logger.info(f"PDF chargé : {pdf_path} — {self.n_pages} pages")

    def parse(self) -> dict:
        result = {
            "info":          self._parse_info(),
            "actif_values":  self._parse_actif(),
            "passif_values": self._parse_passif(),
            "cpc_values":    self._parse_cpc(),
        }
        self._enrich_passif(result)
        return result

    def _enrich_passif(self, data: dict):
        """
        Complète les valeurs passif manquantes en les inférant depuis d'autres sections.
        Résultat net de l'exercice : calculé depuis RESULTAT NET (XI-XII) du CPC.
        """
        pv = data["passif_values"]
        cpc = data["cpc_values"]

        # Résultat net : somme propre_n + prec_n = total_n, total_n1 = net_n1
        rn_key = "Résultat net de l'exercice"
        if rn_key not in pv:
            for k, v in cpc.items():
                if "RESULTAT NET (XI-XII)" in k or "RESULTAT NET (XI" in k:
                    # v = [propre_n, prec_n, total_n1]
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
        # Utiliser le tableau structuré page 1
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

        # Exercice sur pages 2+
        for i in range(1, min(3, self.n_pages)):
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
        # Retourner la dernière cellule non vide et non label
        for c in reversed(cells):
            if len(c) > 2 and not any(k in c.lower() for k in ["raison", "taxe", "identifiant", "adresse", ":"]):
                return c
        return cells[-1] if cells else ""

    # ── Bilan Actif (page 2) ──────────────────────────────────────────────────
    # Colonnes : [0=latéral, 1=label, 2=brut, 3=vide, 4=amort, 5=net_n, 6=net_n1]

    def _parse_actif(self) -> dict:
        values = {}
        for page_idx in range(1, min(3, self.n_pages)):
            tables = self.pages[page_idx].extract_tables()
            for table in tables:
                for row in table:
                    if not row or len(row) < 3:
                        continue
                    label = self._clean_label(row[1] if len(row) > 1 else row[0])
                    if not label or self._should_skip(label):
                        continue

                    # Colonnes : brut=col2, amort=col4, net_n=col5, net_n1=col6
                    brut   = self._parse_num(row[2] if len(row) > 2 else None)
                    amort  = self._parse_num(row[4] if len(row) > 4 else None)
                    net_n1 = self._parse_num(row[6] if len(row) > 6 else None)

                    if any(v is not None for v in [brut, amort, net_n1]):
                        if label not in values:
                            values[label] = [brut, amort, net_n1]

        logger.info(f"Actif : {len(values)} postes avec valeurs")
        return values

    # ── Bilan Passif (page 3) ─────────────────────────────────────────────────
    # Colonnes : [0=latéral, 1=label, 2=vide, 3=val_n, 4=val_n1]

    def _parse_passif(self) -> dict:
        values = {}
        for page_idx in range(2, min(4, self.n_pages)):
            tables = self.pages[page_idx].extract_tables()
            for table in tables:
                for row in table:
                    if not row or len(row) < 3:
                        continue
                    label = self._clean_label(row[1] if len(row) > 1 else row[0])
                    if not label or self._should_skip(label):
                        continue

                    val_n  = self._parse_num(row[3] if len(row) > 3 else None)
                    val_n1 = self._parse_num(row[4] if len(row) > 4 else None)

                    if any(v is not None for v in [val_n, val_n1]):
                        if label not in values:
                            values[label] = [val_n, val_n1]

        logger.info(f"Passif : {len(values)} postes avec valeurs")
        return values

    # ── CPC (pages 4-5) ───────────────────────────────────────────────────────
    # Colonnes : [0=latéral, 1=num_romain, 2=label, 3=propre_n, 4=prec_n, 5=total_n, 6=total_n1]

    def _parse_cpc(self) -> dict:
        """
        Structure des tableaux CPC selon le nombre de colonnes :
          7 cols: [lat, num, label, propre_n, prec_n, total_n, total_n1]
          8 cols: [lat, num, label, propre_n, VIDE,   prec_n, total_n, total_n1]
        On prend toujours : propre_n, prec_n, total_n1 (exercice précédent)
        """
        values = {}
        for page_idx in range(3, min(6, self.n_pages)):
            tables = self.pages[page_idx].extract_tables()
            for table in tables:
                for row in table:
                    if not row or len(row) < 4:
                        continue

                    label_raw = row[2] if len(row) > 2 else None
                    label = self._clean_label(label_raw)
                    if not label or self._should_skip(label):
                        continue

                    n = len(row)
                    if n >= 8:
                        # 8 cols : col 4 vide → prec_n en col 5, total_n1 en col 7
                        propre_n = self._parse_num(row[3])
                        prec_n   = self._parse_num(row[5])
                        total_n1 = self._parse_num(row[7])
                    else:
                        # 7 cols : propre_n col 3, prec_n col 4, total_n1 col 6
                        propre_n = self._parse_num(row[3])
                        prec_n   = self._parse_num(row[4])
                        total_n1 = self._parse_num(row[6] if n > 6 else None)

                    if any(v is not None for v in [propre_n, prec_n, total_n1]):
                        # Mettre à jour si on trouve de meilleures valeurs (prec_n non None)
                        if label not in values:
                            values[label] = [propre_n, prec_n, total_n1]
                        else:
                            existing = values[label]
                            # Enrichir avec prec_n si manquant
                            if existing[1] is None and prec_n is not None:
                                values[label] = [
                                    existing[0] if existing[0] is not None else propre_n,
                                    prec_n,
                                    existing[2] if existing[2] is not None else total_n1,
                                ]

        logger.info(f"CPC : {len(values)} postes avec valeurs")
        return values

    # ── Utilitaires ───────────────────────────────────────────────────────────

    def _should_skip(self, label: str) -> bool:
        l = label.lower()
        if l in SKIP_LABELS:
            return True
        if any(l.startswith(p) for p in SKIP_PREFIXES):
            return True
        # Exact match sur labels de totaux
        if l in TOTAL_SKIP_EXACT:
            return True
        # Préfixe exact sur résultats/totaux
        if any(l.startswith(p) for p in TOTAL_SKIP_PREFIXES):
            return True
        # Ignorer les labels trop courts ou purement numériques
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
        # Enlever numéros romains en début
        s = re.sub(r"^(I{1,3}|IV|V|VI{1,3}|IX|X{1,2})\s+", "", s)
        return s.strip()

    @staticmethod
    def _parse_num(s) -> float | None:
        if s is None:
            return None
        s = str(s).strip().replace("\n", "")
        if not s or s in ["-", "—", "", "None"]:
            return None
        # Négatifs entre parenthèses
        neg = False
        if s.startswith("(") and s.endswith(")"):
            neg = True
            s = s[1:-1]
        if s.startswith("-"):
            neg = True
            s = s[1:]
        # Format marocain : espace = milliers, virgule = décimale
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
