"""
core/template_filler.py
Remplit un template Excel existant avec les données extraites
"""
from openpyxl import load_workbook
from utils.logger import get_logger

logger = get_logger(__name__)

class TemplateFiller:
    """
    Remplit un template Excel avec les données fiscales extraites
    """
    
    def __init__(self, template_path: str):
        self.template_path = template_path
        self.wb = load_workbook(template_path)
        logger.info(f"Template chargé : {template_path}")
    
    def fill_from_data(self, fiscal_data: dict, output_path: str) -> dict:
        """
        Remplit le template avec les données
        fiscal_data = {
            "info": {...},
            "bilan_actif": [...],
            "bilan_passif": [...],
            "cpc": [...]
        }
        """
        # Feuille 1: Infos
        self._fill_info_sheet(fiscal_data.get("info", {}))
        
        # Feuille 2: Bilan Actif
        self._fill_bilan_actif(fiscal_data.get("bilan_actif", []))
        
        # Feuille 3: Bilan Passif
        self._fill_bilan_passif(fiscal_data.get("bilan_passif", []))
        
        # Feuille 4: CPC
        self._fill_cpc(fiscal_data.get("cpc", []))
        
        # Sauvegarder
        self.wb.save(output_path)
        logger.info(f"Template rempli sauvegardé : {output_path}")
        
        return {
            "sheets": len(self.wb.sheetnames),
            "rows_filled": self._count_filled_rows(fiscal_data)
        }
    
    def _fill_info_sheet(self, info: dict):
        """Remplit la feuille Infos Générales"""
        ws = self.wb["1 - Infos Générales"] if "1 - Infos Générales" in self.wb.sheetnames else self.wb.worksheets[0]
        
        mapping = {
            "raison_sociale": ("B5", "Raison sociale"),
            "taxe_professionnelle": ("B6", "Taxe professionnelle"),
            "identifiant_fiscal": ("B7", "Identifiant fiscal"),
            "adresse": ("B8", "Adresse"),
            "exercice": ("B9", "Exercice"),
            "date_declaration": ("B10", "Date de déclaration"),
        }
        
        for key, (cell_ref, label) in mapping.items():
            value = info.get(key, "")
            ws[cell_ref] = value
            logger.debug(f"{label}: {value}")
    
    def _fill_bilan_actif(self, rows: list):
        """Remplit le Bilan Actif par matching de labels"""
        ws = self.wb["2 - Bilan Actif"] if "2 - Bilan Actif" in self.wb.sheetnames else self.wb.worksheets[1]
        
        # Construire un dict label → valeurs
        data_dict = {row[0]: row[1:] for row in rows if row and row[0]}
        
        # Parcourir les lignes du template et remplir
        for row_idx in range(4, ws.max_row + 1):
            label_cell = ws[f"A{row_idx}"].value
            if label_cell and label_cell in data_dict:
                values = data_dict[label_cell]
                # Colonnes: B=Brut, C=Amort, D=Net N, E=Net N-1
                for col_idx, val in enumerate(values[:4], 2):
                    if val is not None:
                        ws.cell(row=row_idx, column=col_idx).value = val
        
        logger.info(f"Bilan Actif: {len(data_dict)} lignes remplies")
    
    def _fill_bilan_passif(self, rows: list):
        """Remplit le Bilan Passif"""
        ws = self.wb["3 - Bilan Passif"] if "3 - Bilan Passif" in self.wb.sheetnames else self.wb.worksheets[2]
        
        data_dict = {row[0]: row[1:] for row in rows if row and row[0]}
        
        for row_idx in range(4, ws.max_row + 1):
            label_cell = ws[f"A{row_idx}"].value
            if label_cell and label_cell in data_dict:
                values = data_dict[label_cell]
                # Colonnes: B=N, C=N-1
                for col_idx, val in enumerate(values[:2], 2):
                    if val is not None:
                        ws.cell(row=row_idx, column=col_idx).value = val
        
        logger.info(f"Bilan Passif: {len(data_dict)} lignes remplies")
    
    def _fill_cpc(self, rows: list):
        """Remplit le CPC"""
        ws = self.wb["4 - CPC"] if "4 - CPC" in self.wb.sheetnames else self.wb.worksheets[3]
        
        data_dict = {row[0]: row[1:] for row in rows if row and row[0]}
        
        for row_idx in range(4, ws.max_row + 1):
            label_cell = ws[f"A{row_idx}"].value
            if label_cell and label_cell in data_dict:
                values = data_dict[label_cell]
                # Colonnes: B=Propre N, C=Exerc Préc, D=Total N, E=Total N-1
                for col_idx, val in enumerate(values[:4], 2):
                    if val is not None:
                        ws.cell(row=row_idx, column=col_idx).value = val
        
        logger.info(f"CPC: {len(data_dict)} lignes remplies")
    
    def _count_filled_rows(self, fiscal_data: dict) -> int:
        return (
            len(fiscal_data.get("bilan_actif", [])) +
            len(fiscal_data.get("bilan_passif", [])) +
            len(fiscal_data.get("cpc", []))
        )
