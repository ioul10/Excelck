"""
core/extractor.py
Étape 1 & 2 : Extraction PDF → données brutes structurées
Version corrigée : extraction par tableaux avec mapping de colonnes
"""
import re
import pdfplumber
from utils.logger import get_logger

logger = get_logger(__name__)

class PDFExtractor:
    """
    Lit un PDF de pièces annexes IS (Modèle Normal) et extrait :
    - Infos générales (page 1) : tableau 2 colonnes
    - Bilan Actif     (pages 2-3) : tableau 5 colonnes
    - Bilan Passif    (page 4)    : tableau 3 colonnes  
    - CPC             (page 5)    : tableau 5 colonnes
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
        """Extraction depuis le tableau page 1 : Champ | Valeur"""
        info = {
            "raison_sociale": "",
            "taxe_professionnelle": "",
            "identifiant_fiscal": "",
            "adresse": "",
            "exercice": "",
            "date_declaration": "",
            "pages": self.num_pages
        }
        
        # Mapping des champs attendus
        field_mapping = {
            "raison sociale": "raison_sociale",
            "taxe professionnelle": "taxe_professionnelle", 
            "identifiant fiscal": "identifiant_fiscal",
            "adresse": "adresse",
            "exercice": "exercice",
            "date de déclaration": "date_declaration",
            "date de declaration": "date_declaration",  # sans accent fallback
        }
        
        # Extraire depuis les tableaux de la page 1 (et 2 si besoin)
        for page_idx in range(min(2, self.num_pages)):
            tables = self.pages[page_idx].extract_tables()
            
            for table in tables:
                if not table:
                    continue
                    
                for row in table:
                    if not row or len(row) < 2:
                        continue
                    
                    # Nettoyer les cellules
                    label = self._clean_cell(row[0]).lower()
                    value = self._clean_cell(row[1])
                    
                    # Matching avec le mapping
                    for key_phrase, field_name in field_mapping.items():
                        if key_phrase in label and not info.get(field_name):
                            # Nettoyer la valeur
                            if value and value not in ["", "-", "None", "—"]:
                                info[field_name] = value.strip()
                            break
        
        # Fallback : chercher l'exercice dans le texte si pas trouvé en tableau
        if not info["exercice"]:
            text = self.get_all_text()
            m = re.search(
                r"[Ee]xercice\s+(?:du\s+)?(\d{2}/\d{2}/\d{4})\s+(?:au|–|—)\s*(\d{2}/\d{2}/\d{4})?",
                text
            )
            if m:
                if m.group(2):
                    info["exercice"] = f"Du {m.group(1)} au {m.group(2)}"
                else:
                    info["exercice"] = m.group(1)
        
        logger.info(f"Infos extraites: {info}")
        return info
    
    # ── Bilan Actif ───────────────────────────────────────────────────────────
    def _extract_bilan_actif(self) -> list:
        """
        Extrait le Bilan Actif (pages 2-3)
        Structure: [Label, Brut N, Amort/Prov, Net N, Net N-1]
        Retourne: list de tuples (label, brut, amort, net_n, net_n1)
        """
        rows = []
        
        for page_idx in range(1, min(4, self.num_pages)):  # pages 2, 3, 4 (index 1,2,3)
            tables = self.pages[page_idx].extract_tables()
            
            for table in tables:
                page_rows = self._parse_fiscal_table(table, type="actif")
                if page_rows:
                    rows.extend(page_rows)
                    break  # Un seul tableau par page pour le bilan
        
        rows = self._deduplicate(rows)
        logger.info(f"Bilan Actif : {len(rows)} lignes extraites")
        return rows
    
    def _is_bilan_actif_page(self, text: str) -> bool:
        keywords = ["ACTIF", "Immobilisations", "Stocks", "Créances", "Trésorerie"]
        return sum(1 for k in keywords if k.upper() in text.upper()) >= 2
    
    # ── Bilan Passif ──────────────────────────────────────────────────────────
    def _extract_bilan_passif(self) -> list:
        """
        Extrait le Bilan Passif (page 4)
        Structure: [Label, Exercice N, Exercice N-1]
        Retourne: list de tuples (label, val_n, val_n1)
        """
        rows = []
        
        for page_idx in range(2, min(5, self.num_pages)):  # pages 3, 4, 5
            tables = self.pages[page_idx].extract_tables()
            
            for table in tables:
                page_rows = self._parse_fiscal_table(table, type="passif")
                if page_rows:
                    rows.extend(page_rows)
                    break
        
        rows = self._deduplicate(rows)
        logger.info(f"Bilan Passif : {len(rows)} lignes extraites")
        return rows
    
    def _is_bilan_passif_page(self, text: str) -> bool:
        keywords = ["PASSIF", "Capitaux propres", "Dettes", "Subvention"]
        return sum(1 for k in keywords if k.upper() in text.upper()) >= 2
    
    # ── CPC ───────────────────────────────────────────────────────────────────
    def _extract_cpc(self) -> list:
        """
        Extrait le CPC (page 5)
        Structure: [Label, Propre N, Exerc Préc, Total N, Total N-1]
        Retourne: list de tuples (label, propre_n, prec_n, total_n, total_n1)
        """
        rows = []
        
        for page_idx in range(3, min(6, self.num_pages)):  # pages 4, 5, 6
            tables = self.pages[page_idx].extract_tables()
            
            for table in tables:
                page_rows = self._parse_fiscal_table(table, type="cpc")
                if page_rows:
                    rows.extend(page_rows)
                    break
        
        rows = self._deduplicate(rows)
        logger.info(f"CPC : {len(rows)} lignes extraites")
        return rows
    
    def _is_cpc_page(self, text: str) -> bool:
        keywords = ["PRODUITS", "CHARGES", "EXPLOITATION", "RÉSULTAT"]
        return sum(1 for k in keywords if k.upper() in text.upper()) >= 3
    
    # ── NOUVELLE MÉTHODE PRINCIPALE : Parser tableau fiscal ───────────────────
    def _parse_fiscal_table(self, table: list, type: str) -> list:
        """
        Parse un tableau fiscal extrait par pdfplumber
        Gère le mapping des colonnes selon le type de bilan
        """
        rows = []
        
        if not table:
            return rows
        
        # Définir le nombre de colonnes attendues selon le type
        col_config = {
            "actif":  {"label_col": 0, "value_cols": [1, 2, 3, 4], "expected": 5},
            "passif": {"label_col": 0, "value_cols": [1, 2], "expected": 3},
            "cpc":    {"label_col": 0, "value_cols": [1, 2, 3, 4], "expected": 5},
        }
        
        config = col_config.get(type)
        if not config:
            return rows
        
        label_col = config["label_col"]
        value_cols = config["value_cols"]
        
        for row_idx, row in enumerate(table):
            if not row or len(row) <= label_col:
                continue
            
            # Nettoyer la cellule label
            label = self._clean_cell(row[label_col])
            
            # Ignorer les lignes vides, en-têtes et séparateurs
            if self._should_skip_row(label, type):
                continue
            
            # Extraire les valeurs numériques des colonnes appropriées
            values = []
            for col_idx in value_cols:
                if col_idx < len(row):
                    cell_val = self._clean_cell(row[col_idx])
                    num_val = self._parse_number(cell_val)
                    values.append(num_val)
                else:
                    values.append(None)
            
            # Ajouter la ligne si on a au moins un label valide
            if label and len(label) >= 2:
                rows.append((label, *values))
        
        return rows
    
    def _should_skip_row(self, label: str, type: str) -> bool:
        """Détermine si une ligne doit être ignorée"""
        if not label or len(label) < 2:
            return True
        
        label_upper = label.upper()
        
        # Mots-clés d'en-tête à ignorer
        skip_keywords = [
            "BILAN", "ACTIF", "PASSIF", "CPC", "COMPTE DE PRODUITS",
            "EXERCICE", "BRUT", "AMORT", "NET", "TOTAL", "DÉSIGNATION",
            "PROPRE", "PRÉCÉDENTS", "CHARGES", "PRODUITS",
            "TABLEAU", "PIÈCES ANNEXES", "IMPÔTS", "MODÈLE",
            "AGENCE DU BASSIN", "IF:", "HORS TAXES"
        ]
        
        if any(kw in label_upper for kw in skip_keywords):
            return True
        
        # Ignorer les lignes qui ne contiennent que des tirets ou symboles
        if re.match(r'^[\s\-\–—\|\.]+$', label):
            return True
        
        return False
    
    # ── Utilitaires ───────────────────────────────────────────────────────────
    def _clean_cell(self, cell) -> str:
        """Nettoie et normalise une cellule de tableau"""
        if cell is None:
            return ""
        
        result = str(cell).strip()
        
        # Remplacer les tirets/vides par chaîne vide
        if result in ["-", "–", "—", ".", "None", "none", ""]:
            return ""
        
        # Nettoyer les espaces multiples
        result = re.sub(r'\s+', ' ', result)
        
        return result
    
    @staticmethod
    def _parse_number(s: str):
        """Convertit une chaîne en float, gère le format marocain"""
        if not s or s in ["-", "–", "—", ".", "None", ""]:
            return None
        
        # Nettoyer: espaces, NBSP, puis convertir virgule en point
        cleaned = s.strip().replace(" ", "").replace("\xa0", "").replace(",", ".")
        
        # Gérer les parenthèses pour nombres négatifs: (123.45) → -123.45
        if cleaned.startswith("(") and cleaned.endswith(")"):
            cleaned = "-" + cleaned[1:-1]
        
        try:
            v = float(cleaned)
            # Sanity check: rejeter les valeurs absurdes
            return v if abs(v) < 1e15 else None
        except (ValueError, TypeError):
            return None
    
    @staticmethod
    def _normalize_label(label: str) -> str:
        """Normalise un libellé pour le matching"""
        if not label:
            return ""
        
        # Supprimer les caractères parasites en début/fin
        label = re.sub(r'^[\s:\-\.\|\[\]\(\)]+', '', label).strip()
        label = re.sub(r'[\s:\-\.\|\[\]\(\)]+$', '', label).strip()
        
        # Normaliser les espaces multiples
        label = re.sub(r'\s{2,}', ' ', label)
        
        return label.strip()
    
    @staticmethod
    def _deduplicate(rows: list) -> list:
        """Supprime les doublons en gardant la première occurrence"""
        seen = {}
        result = []
        
        for row in rows:
            if not row:
                continue
            key = row[0]  # Le label est la clé
            if key and key not in seen:
                seen[key] = True
                result.append(row)
        
        return result
    
    def __del__(self):
        """Fermer proprement le PDF"""
        try:
            if hasattr(self, '_pdf') and self._pdf:
                self._pdf.close()
        except Exception:
            pass
