"""
core/extractor.py
Étape 1 & 2 : Extraction PDF → données brutes structurées
Version améliorée avec extraction par tableaux + texte
"""
import re
import pdfplumber
from utils.logger import get_logger

logger = get_logger(__name__)

class PDFExtractor:
    """
    Lit un PDF de pièces annexes IS (Modèle Normal) et extrait :
    - Infos générales (page 1)
    - Bilan Actif     (pages 2-3)
    - Bilan Passif    (pages 3-4)
    - CPC             (pages 4-5)
    """
    
    def __init__(self, pdf_path: str):
        self.pdf_path = pdf_path
        self._pdf = pdfplumber.open(pdf_path)
        self.pages = self._pdf.pages
        self.num_pages = len(self.pages)
        logger.info(f"PDF ouvert : {pdf_path} ({self.num_pages} pages)")
    
    # ── API publique ──────────────────────────────────────────────────────────
    def extract_all(self) -> dict:
        return {
            "info":         self._extract_info(),
            "bilan_actif":  self._extract_bilan_actif(),
            "bilan_passif": self._extract_bilan_passif(),
            "cpc":          self._extract_cpc(),
        }
    
    def get_raw_text(self, page_index: int) -> str:
        if page_index >= self.num_pages:
            return ""
        return self.pages[page_index].extract_text() or ""
    
    def get_all_text(self) -> str:
        return "\n".join(self.get_raw_text(i) for i in range(self.num_pages))
    
    # ── Infos générales ───────────────────────────────────────────────────────
    def _extract_info(self) -> dict:
        text = self.get_raw_text(0)
        info = {}
        
        patterns = {
            "raison_sociale":       r"Raison sociale\s*[:\-]?\s*(.+?)(?:\n|Taxe|$)",
            "taxe_professionnelle": r"Taxe professionnelle\s*[:\-]?\s*(\d+)",
            "identifiant_fiscal":   r"Identifiant fiscal\s*[:\-]?\s*(\d+)",
            "adresse":              r"Adresse\s*[:\-]?\s*(.+?)(?:\n|$)",
            "date_declaration":     r"(?:FES|FÈS|RABAT|CASABLANCA)\s+[Ll]e\s+(\d{2}/\d{2}/\d{4})",
        }
        
        for key, pattern in patterns.items():
            m = re.search(pattern, text, re.IGNORECASE | re.DOTALL)
            if m:
                info[key] = m.group(1).strip()
            else:
                info[key] = ""
        
        # Exercice : chercher sur la 2e page si pas trouvé
        for page_idx in range(min(3, self.num_pages)):
            t = self.get_raw_text(page_idx)
            m = re.search(
                r"[Ee]xercice\s+du\s+(\d{2}/\d{2}/\d{4})\s+au\s+(\d{2}/\d{2}/\d{4})", 
                t
            )
            if m:
                info["exercice_debut"] = m.group(1)
                info["exercice_fin"]   = m.group(2)
                info["exercice"]       = f"Du {m.group(1)} au {m.group(2)}"
                break
        
        if not info.get("exercice"):
            info["exercice"] = ""
        
        # Raison sociale : parfois sur la 2e page
        if not info.get("raison_sociale"):
            for page_idx in range(1, min(3, self.num_pages)):
                t = self.get_raw_text(page_idx)
                m = re.search(r"(?:AGENCE|SOCIETE|SOCIÉTÉ|ENTREPRISE)[^\n]{3,80}", t)
                if m:
                    info["raison_sociale"] = m.group(0).strip()
                    break
        
        info["pages"] = self.num_pages
        return info
    
    # ── Bilan Actif ───────────────────────────────────────────────────────────
    def _extract_bilan_actif(self) -> list:
        """
        Retourne une liste de tuples :
        (label, brut, amortissements, net_n, net_n1)
        """
        rows = []
        
        # Le bilan actif est généralement sur la page 2 (index 1)
        for page_idx in range(1, min(4, self.num_pages)):
            # Essayer d'abord l'extraction par tableaux
            tables = self.pages[page_idx].extract_tables()
            
            if tables:
                for table in tables:
                    page_rows = self._parse_table_to_rows(table, type="actif")
                    rows.extend(page_rows)
            
            # Si pas de tableaux, utiliser le texte
            if not rows:
                text = self.get_raw_text(page_idx)
                if self._is_bilan_actif_page(text):
                    page_rows = self._parse_bilan_actif_page(text)
                    rows.extend(page_rows)
        
        # Dédoublonner les totaux si spread sur 2 pages
        rows = self._deduplicate(rows)
        logger.info(f"Bilan Actif : {len(rows)} lignes extraites")
        return rows
    
    def _is_bilan_actif_page(self, text: str) -> bool:
        keywords = ["Immobilisations", "ACTIF", "Stocks", "Créances", "Trésorerie"]
        return sum(1 for k in keywords if k in text) >= 3
    
    def _parse_bilan_actif_page(self, text: str) -> list:
        """Parse ligne par ligne le texte du bilan actif."""
        rows = []
        lines = [l.strip() for l in text.split("\n") if l.strip()]
        
        for line in lines:
            # Ignorer en-têtes
            if any(h in line for h in [
                "Brut", "Amortissements", "EXERCICE", "Tableau", "Bilan",
                "identifiant", "Exercice du", "ACTIF", "NET"
            ]):
                continue
            
            label, nums = self._split_label_nums(line)
            if not label:
                continue
            
            # Normaliser
            label = self._normalize_label(label)
            
            # Filtrer les lignes trop courtes ou sans chiffres
            if len(label) < 3 and not nums:
                continue
            
            brut = nums[0] if len(nums) > 0 else None
            amort = nums[1] if len(nums) > 1 else None
            net_n = nums[2] if len(nums) > 2 else None
            net_n1 = nums[3] if len(nums) > 3 else None
            
            rows.append((label, brut, amort, net_n, net_n1))
        
        return rows
    
    # ── Bilan Passif ──────────────────────────────────────────────────────────
    def _extract_bilan_passif(self) -> list:
        rows = []
        
        for page_idx in range(2, min(5, self.num_pages)):
            # Essayer tableaux d'abord
            tables = self.pages[page_idx].extract_tables()
            
            if tables:
                for table in tables:
                    page_rows = self._parse_table_to_rows(table, type="passif")
                    rows.extend(page_rows)
            
            # Fallback texte
            if not rows:
                text = self.get_raw_text(page_idx)
                if self._is_bilan_passif_page(text):
                    rows.extend(self._parse_bilan_passif_page(text))
        
        rows = self._deduplicate(rows)
        logger.info(f"Bilan Passif : {len(rows)} lignes extraites")
        return rows
    
    def _is_bilan_passif_page(self, text: str) -> bool:
        keywords = ["PASSIF", "Capitaux propres", "Capital", "Dettes", "Subvention"]
        return sum(1 for k in keywords if k in text) >= 3
    
    def _parse_bilan_passif_page(self, text: str) -> list:
        rows = []
        lines = [l.strip() for l in text.split("\n") if l.strip()]
        
        for line in lines:
            if any(h in line for h in [
                "PASSIF", "EXERCICE", "Tableau", "Bilan", "identifiant", 
                "Exercice du", "(1)Capital", "(2)Bénéficiaire"
            ]):
                continue
            
            label, nums = self._split_label_nums(line)
            if not label:
                continue
            
            label = self._normalize_label(label)
            val_n  = nums[0] if len(nums) > 0 else None
            val_n1 = nums[1] if len(nums) > 1 else None
            
            rows.append((label, val_n, val_n1))
        
        return rows
    
    # ── CPC ───────────────────────────────────────────────────────────────────
    def _extract_cpc(self) -> list:
        rows = []
        
        for page_idx in range(3, min(6, self.num_pages)):
            # Essayer tableaux
            tables = self.pages[page_idx].extract_tables()
            
            if tables:
                for table in tables:
                    page_rows = self._parse_table_to_rows(table, type="cpc")
                    rows.extend(page_rows)
            
            # Fallback texte
            if not rows:
                text = self.get_raw_text(page_idx)
                if self._is_cpc_page(text):
                    rows.extend(self._parse_cpc_page(text))
        
        rows = self._deduplicate(rows)
        logger.info(f"CPC : {len(rows)} lignes extraites")
        return rows
    
    def _is_cpc_page(self, text: str) -> bool:
        keywords = ["Produits", "Charges", "EXPLOITATION", "RÉSULTAT", "RESULTAT"]
        return sum(1 for k in keywords if k in text) >= 3
    
    def _parse_cpc_page(self, text: str) -> list:
        rows = []
        lines = [l.strip() for l in text.split("\n") if l.strip()]
        
        for line in lines:
            if any(h in line for h in [
                "DESIGNATION", "DÉSIGNATION", "Propres à", "Concernant",
                "TOTAUX", "Tableau", "identifiant", "Exercice du",
                "1)Variation", "2)Achats"
            ]):
                continue
            
            label, nums = self._split_label_nums(line)
            if not label:
                continue
            
            label = self._normalize_label(label)
            prop_n  = nums[0] if len(nums) > 0 else None
            prec_n  = nums[1] if len(nums) > 1 else None
            total_n = nums[2] if len(nums) > 2 else None
            total_n1= nums[3] if len(nums) > 3 else None
            
            rows.append((label, prop_n, prec_n, total_n, total_n1))
        
        return rows
    
    # ── Nouvelle méthode : Parser les tableaux extraits ───────────────────────
    def _parse_table_to_rows(self, table: list, type: str = "actif") -> list:
        """
        Parse un tableau extrait par pdfplumber.extract_tables()
        Retourne une liste de tuples selon le type
        """
        rows = []
        
        if not table:
            return rows
        
        for row_idx, row in enumerate(table):
            if not row:
                continue
            
            # Nettoyer les cellules
            cleaned_row = [self._clean_cell(cell) for cell in row]
            
            # Filtrer les lignes d'en-tête
            if any(h in str(cleaned_row).upper() for h in [
                "BRUT", "AMORT", "EXERCICE", "NET", "TOTAL", "ACTIF", "PASSIF"
            ]):
                continue
            
            # Extraire label et valeurs
            label = None
            values = []
            
            for cell in cleaned_row:
                if not cell:
                    continue
                
                # Tester si c'est un nombre
                num_val = self._parse_number(cell)
                if num_val is not None:
                    values.append(num_val)
                elif not label and len(cell) > 2:
                    # C'est probablement le label
                    label = self._normalize_label(cell)
            
            if label and values:
                if type == "actif":
                    # (label, brut, amort, net_n, net_n1)
                    while len(values) < 4:
                        values.append(None)
                    rows.append((label, *values[:4]))
                elif type == "passif":
                    # (label, val_n, val_n1)
                    while len(values) < 2:
                        values.append(None)
                    rows.append((label, *values[:2]))
                elif type == "cpc":
                    # (label, propre_n, prec_n, total_n, total_n1)
                    while len(values) < 4:
                        values.append(None)
                    rows.append((label, *values[:4]))
        
        return rows
    
    def _clean_cell(self, cell) -> str:
        """Nettoie une cellule de tableau"""
        if cell is None:
            return ""
        return str(cell).strip()
    
    # ── Utilitaires ───────────────────────────────────────────────────────────
    @staticmethod
    def _split_label_nums(line: str):
        """
        Sépare une ligne en (texte_label, [liste_de_nombres]).
        Ex: "Terrains  6 100 375,00  6 100 375,00" → ("Terrains", [6100375.0, 6100375.0])
        """
        # Regex pour nombres marocains : 1 234 567,89 ou 1234567.89
        num_pattern = r"-?\d[\d\s]*(?:[,\.]\d{2})?"
        nums_found = re.findall(num_pattern, line)
        
        parsed_nums = []
        for n in nums_found:
            v = PDFExtractor._parse_number(n)
            if v is not None:
                parsed_nums.append(v)
        
        # Le label est ce qui reste quand on retire les nombres
        label = re.sub(num_pattern, "", line).strip()
        label = re.sub(r"\s{2,}", " ", label).strip("[]()- \t")
        
        # Ignorer les lignes purement numériques sans label
        if not label or len(label) < 3:
            return None, []
        
        return label, parsed_nums
    
    @staticmethod
    def _parse_number(s: str):
        if not s:
            return None
        
        s = s.strip().replace(" ", "").replace("\xa0", "")
        # Format marocain : virgule = décimale, espace = milliers
        s = s.replace(",", ".")
        
        try:
            v = float(s)
            return v if abs(v) < 1e12 else None  # sanity check
        except ValueError:
            return None
    
    @staticmethod
    def _normalize_label(label: str) -> str:
        """Nettoie et normalise un libellé."""
        label = label.strip()
        # Supprimer caractères parasites en début/fin
        label = re.sub(r"^[:\-\.\|]+", "", label).strip()
        label = re.sub(r"[:\-\.\|]+$", "", label).strip()
        # Normaliser espaces
        label = re.sub(r"\s{2,}", " ", label)
        return label
    
    @staticmethod
    def _deduplicate(rows: list) -> list:
        """Supprime les doublons consécutifs de même label."""
        seen = set()
        result = []
        
        for row in rows:
            key = row[0] if row else None
            if key and key not in seen:
                seen.add(key)
                result.append(row)
        
        return result
    
    def __del__(self):
        try:
            self._pdf.close()
        except Exception:
            pass
