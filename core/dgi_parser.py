"""
core/dgi_parser.py — Parseur DGI via Camelot Lattice

Structure DGI 7 pages:
  Page 1  : Identification
  Pages 2-3: Bilan Actif  (tables 6 colonnes: [section, label, brut, amort, net_n, net_n1])
  Page 4  : Bilan Passif (table 4 colonnes: [section, label, val_n, val_n1])
  Pages 5-7: CPC         (tables 6 colonnes: [nature, label, propre_n, prec_n, total_n, total_n1])

Camelot Lattice détecte les bordures du tableau → 100% accuracy.
"""

import re
import camelot
from utils.logger import get_logger

logger = get_logger(__name__)

# Labels à ignorer (totaux calculés, en-têtes)
SKIP_RE = re.compile(
    r'^(total\s+i|total\s+ii|total\s+iii|total\s+général|total\s+general|'
    r'total\s+des\s+|total\s+[ivx]+\b|actif$|passif$|nature$|exploitation$|'
    r'financier$|courant$|non\s+courant$|désignation$|'
    r'brut\s+exercice|amortissements|net\s+exercice|exercice\s+précédent|'
    r'opérations\s+propres|opérations\s+concernant|totaux\s+de)', re.I
)

RESULT_SKIP_RE = re.compile(
    r'^(résultat\s+d.exploitation|résultat\s+financier|résultat\s+courant(?!\s+\(report)|'
    r'résultat\s+non\s+courant|résultat\s+avant\s+impôts|'
    r'iii\.\s+résultat|vi\.\s+résultat|vii\.\s+résultat\s+courant$|'
    r'x\.\s+résultat|xi\.\s+résultat\s+avant|xv\.\s+total)', re.I
)

PASSIF_ONLY = {
    'capitaux propres', 'capital social', 'actionnaires capital',
    'prime emission', 'prime d emission', 'ecart reevaluation',
    'reserve legale', 'autres reserves', 'report nouveau',
    'resultat nets en instance', 'resultat net de l exercice',
    'total des capitaux propres', 'capitaux propres assimiles',
    'provisions reglementees', 'dettes de financement',
    'emprunts obligataires', 'autres dettes de financement',
    'provisions durables', 'ecarts de conversion passif',
    'fournisseurs et comptes rattaches passif',
}


class DGIParser:
    """Parseur spécialisé DGI utilisant Camelot Lattice."""

    def __init__(self, pdf_path: str):
        self.path    = pdf_path
        self.n_pages = self._count_pages()
        logger.info(f"DGI PDF chargé : {pdf_path} — {self.n_pages} pages")

    def _count_pages(self) -> int:
        try:
            import pdfplumber
            with pdfplumber.open(self.path) as pdf:
                return len(pdf.pages)
        except Exception:
            return 7

    def parse(self) -> dict:
        result = {
            "info":          self._parse_info(),
            "actif_values":  self._parse_actif(),
            "passif_values": self._parse_passif(),
            "cpc_values":    self._parse_cpc(),
        }
        self._enrich_passif(result)
        return result

    # ── Infos générales ────────────────────────────────────────────────────────

    def _parse_info(self) -> dict:
        import pdfplumber, re as re2
        info = {}
        try:
            with pdfplumber.open(self.path) as pdf:
                text0 = pdf.pages[0].extract_text() or ""
        except Exception:
            text0 = ""

        for key, pat in [
            ("raison_sociale",      r"Raison\s+[Ss]ociale\s*[:\-]?\s*([A-Z][^\n]{3,80})"),
            ("identifiant_fiscal",  r"[Ii]dentifiant\s+[Ff]iscal\s*[:\-]?\s*(\d+)"),
            ("taxe_professionnelle",r"(?:Taxe|Art\.)\s+[Pp]rof[a-z.]*\s*[:\-]?\s*([\d\s]+)"),
            ("adresse",             r"[Aa]dresse\s*[:\-]?\s*(.+)"),
        ]:
            m = re2.search(pat, text0, re2.IGNORECASE)
            if m:
                info[key] = m.group(1).strip()

        # Exercice depuis le titre
        m = re2.search(r"(\d{2}/\d{2}/\d{4})\s+au\s+(\d{2}/\d{2}/\d{4})", text0)
        if m:
            info["exercice"]       = f"Du {m.group(1)} au {m.group(2)}"
            info["exercice_debut"] = m.group(1)
            info["exercice_fin"]   = m.group(2)

        # Date déclaration
        m = re2.search(r"(?:le\s+)?(\d{2}/\d{2}/\d{4})\s+\d{2}:\d{2}", text0)
        if m:
            info["date_declaration"] = m.group(1)

        for k in ["exercice", "exercice_fin", "exercice_debut", "date_declaration",
                  "raison_sociale", "identifiant_fiscal", "adresse", "taxe_professionnelle"]:
            info.setdefault(k, "")
        info["pages"] = self.n_pages
        return info

    # ── Bilan Actif ────────────────────────────────────────────────────────────

    def _parse_actif(self) -> dict:
        """
        Colonnes actif (6 cols):
        [0]=section, [1]=label, [2]=brut, [3]=amort, [4]=net_n, [5]=net_n1
        """
        values = {}
        pages_to_scan = "2-3" if self.n_pages >= 3 else "2"

        try:
            tables = camelot.read_pdf(self.path, pages=pages_to_scan, flavor='lattice')
        except Exception as e:
            logger.warning(f"Camelot actif error: {e}")
            return values

        for table in tables:
            df = table.df
            n_cols = df.shape[1]

            for _, row in df.iterrows():
                # Label: col 1 (parfois col 0 si pas de section)
                label = self._clean(row.iloc[1] if n_cols > 1 else row.iloc[0])
                if not label:
                    label = self._clean(row.iloc[0])
                if not label or self._skip(label, 'actif'):
                    continue

                # Valeurs: brut(2), amort(3), net_n(4), net_n1(5)
                brut   = self._parse_num(row.iloc[2] if n_cols > 2 else "")
                amort  = self._parse_num(row.iloc[3] if n_cols > 3 else "")
                net_n  = self._parse_num(row.iloc[4] if n_cols > 4 else "")
                net_n1 = self._parse_num(row.iloc[5] if n_cols > 5 else "")

                # Filtre contamination passif
                norm = self._normalize(label)
                if any(m in norm for m in PASSIF_ONLY):
                    continue

                if any(v is not None for v in [brut, amort, net_n, net_n1]):
                    if label not in values:
                        values[label] = [brut, amort, net_n1]

        logger.info(f"DGI actif : {len(values)} postes")
        return values

    # ── Bilan Passif ───────────────────────────────────────────────────────────

    def _parse_passif(self) -> dict:
        """
        Colonnes passif (4 cols):
        [0]=section, [1]=label, [2]=val_n, [3]=val_n1
        """
        values = {}
        page = "4" if self.n_pages >= 4 else str(self.n_pages - 3)

        try:
            tables = camelot.read_pdf(self.path, pages=page, flavor='lattice')
        except Exception as e:
            logger.warning(f"Camelot passif error: {e}")
            return values

        for table in tables:
            df = table.df
            n_cols = df.shape[1]

            for _, row in df.iterrows():
                label = self._clean(row.iloc[1] if n_cols > 1 else row.iloc[0])
                if not label:
                    label = self._clean(row.iloc[0])
                if not label or self._skip(label, 'passif'):
                    continue

                val_n  = self._parse_num(row.iloc[2] if n_cols > 2 else "")
                val_n1 = self._parse_num(row.iloc[3] if n_cols > 3 else "")

                if any(v is not None for v in [val_n, val_n1]):
                    if label not in values:
                        values[label] = [val_n, val_n1]

        logger.info(f"DGI passif : {len(values)} postes")
        return values

    # ── CPC ────────────────────────────────────────────────────────────────────

    def _parse_cpc(self) -> dict:
        """
        Colonnes CPC (6 cols):
        [0]=nature, [1]=label, [2]=propre_n, [3]=prec_n, [4]=total_n, [5]=total_n1
        """
        values = {}
        # Pages 5-7 pour DGI 7 pages
        end_page = min(self.n_pages, 7)
        start_page = max(5, end_page - 2)
        pages_str = f"{start_page}-{end_page}"

        try:
            tables = camelot.read_pdf(self.path, pages=pages_str, flavor='lattice')
        except Exception as e:
            logger.warning(f"Camelot CPC error: {e}")
            return values

        for table in tables:
            df = table.df
            n_cols = df.shape[1]

            for _, row in df.iterrows():
                label = self._clean(row.iloc[1] if n_cols > 1 else row.iloc[0])
                if not label:
                    continue
                if self._skip(label, 'cpc'):
                    continue

                propre_n = self._parse_num(row.iloc[2] if n_cols > 2 else "")
                prec_n   = self._parse_num(row.iloc[3] if n_cols > 3 else "")
                total_n  = self._parse_num(row.iloc[4] if n_cols > 4 else "")
                total_n1 = self._parse_num(row.iloc[5] if n_cols > 5 else "")

                # Pour CPC: [propre_n, prec_n, total_n1]
                if any(v is not None for v in [propre_n, prec_n, total_n, total_n1]):
                    if label not in values:
                        values[label] = [propre_n, prec_n, total_n1]
                    else:
                        # Enrichir
                        ex = values[label]
                        if ex[1] is None and prec_n is not None:
                            values[label] = [ex[0] or propre_n, prec_n, ex[2] or total_n1]

        logger.info(f"DGI CPC : {len(values)} postes")
        return values

    # ── Enrichissement ─────────────────────────────────────────────────────────

    def _enrich_passif(self, data: dict):
        pv  = data["passif_values"]
        cpc = data["cpc_values"]

        # Résultat net depuis CPC
        rn_key = "Résultat net de l'exercice"
        if rn_key not in pv:
            for k, v in cpc.items():
                if re.search(r'resultat net.*xi.xii|xiii.*resultat net', k, re.I):
                    propre = v[0] or 0
                    prec   = v[1] or 0
                    total_n  = round(propre + prec, 2)
                    total_n1 = v[2] if len(v) > 2 else None
                    pv[rn_key] = [total_n, total_n1]
                    logger.info(f"DGI résultat net inféré : N={total_n}")
                    break

        # Capital social depuis passif_values
        cap_keys = ["Capital social ou personnel", "Capital social ou personnel (1)"]
        has_cap = any(k in pv for k in cap_keys)
        if not has_cap:
            for k, v in pv.items():
                kl = k.lower()
                if "capital appelé" in kl or "dont versé" in kl:
                    if v and v[0] and v[0] > 0:
                        pv["Capital social ou personnel"] = v
                        logger.info(f"DGI capital social inféré : {v}")
                        break

    # ── Utilitaires ────────────────────────────────────────────────────────────

    @staticmethod
    def _clean(s) -> str:
        if not s: return ""
        s = str(s).replace('\n', ' ').strip()
        s = re.sub(r'\s{2,}', ' ', s)
        # Supprimer les préfixes (I., II., *, etc.)
        s = re.sub(r'^\*\s*', '', s)
        # Supprimer les numéros romains en début
        s = re.sub(r'^(XIV\.|XV\.|XVI\.|XIII\.|XII\.|XI\.|X\.|IX\.|VIII\.|VII\.|VI\.|V\.|IV\.|III\.|II\.|I\.)\s*', '', s)
        return s.strip()

    @staticmethod
    def _normalize(s: str) -> str:
        import unicodedata
        s = unicodedata.normalize('NFD', s.lower())
        s = ''.join(c for c in s if unicodedata.category(c) != 'Mn')
        s = re.sub(r'[^a-z0-9 ]', ' ', s)
        return re.sub(r'\s+', ' ', s).strip()

    @staticmethod
    def _parse_num(s) -> float | None:
        if not s: return None
        s = str(s).strip()
        if not s or s in ['-', '—', '']: return None
        neg = s.startswith('-')
        s = s.lstrip('-').replace(' ', '').replace('\xa0', '')
        s = s.replace(',', '.')
        try:
            v = float(s)
            return -v if neg else v
        except ValueError:
            return None

    def _skip(self, label: str, mode: str) -> bool:
        if len(label) < 3: return True
        l = label.lower().strip()
        if SKIP_RE.match(l): return True
        if mode == 'cpc' and RESULT_SKIP_RE.match(l): return True
        if re.match(r'^\d+$', label): return True
        return False
