import pdfplumber
import pandas as pd
import re
import os
from typing import Dict, List, Any

def clean_number(val):
    """
    Nettoie une chaîne de caractères pour la convertir en nombre flottant.
    Gère les formats marocains/français : "1 234,56" -> 1234.56
    """
    if val is None:
        return None
    if isinstance(val, (int, float)):
        return val
    
    text = str(val).strip()
    
    # Cas spéciaux
    if not text or text == "-" or text == "":
        return None
    
    # Supprime les espaces (séparateurs de milliers) et remplace la virgule par un point
    text = text.replace(" ", "").replace(",", ".")
    
    try:
        return float(text)
    except ValueError:
        # Si ce n'est pas un nombre, on retourne le texte original (ex: libellé)
        return text

def extract_metadata(pages):
    """
    Extrait les métadonnées (Raison sociale, Exercice, IF) depuis la première page.
    """
    info = {
        "raison_sociale": "Inconnu",
        "exercice_debut": "",
        "exercice_fin": "",
        "identifiant_fiscal": ""
    }
    
    if not pages:
        return info
        
    text_page_0 = pages[0].extract_text()
    if not text_page_0:
        return info

    # Regex Raison Sociale
    match_rs = re.search(r'Raison sociale\s*[:\n]?\s*(.+?)(?:\n|Identifiant|Adresse)', text_page_0, re.IGNORECASE)
    if match_rs:
        info["raison_sociale"] = match_rs.group(1).strip()
    
    # Regex Identifiant Fiscal
    match_if = re.search(r'Identifiant fiscal\s*[:\n]?\s*(\d+)', text_page_0, re.IGNORECASE)
    if match_if:
        info["identifiant_fiscal"] = match_if.group(1)

    # Regex Exercice (Du ... au ...)
    match_ex = re.search(r'Du\s+(\d{2}/\d{2}/\d{4})\s+au\s+(\d{2}/\d{2}/\d{4})', text_page_0)
    if match_ex:
        info["exercice_debut"] = match_ex.group(1)
        info["exercice_fin"] = match_ex.group(2)
    else:
        # Fallback sur l'année seule si format différent
        match_year = re.search(r'(\d{4})', text_page_0)
        if match_year:
            info["exercice_fin"] = f"31/12/{match_year.group(1)}"

    return info

def process_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    """
    Nettoie un DataFrame extrait :
    1. Supprime les lignes totalement vides.
    2. Convertit les colonnes numériques (sauf la première qui est souvent le libellé).
    3. Renomme les colonnes si nécessaire pour éviter les doublons.
    """
    if df.empty:
        return df

    # Suppression des lignes vides
    df = df.dropna(how='all')
    
    # Réinitialisation des index
    df = df.reset_index(drop=True)

    # Conversion des nombres : on suppose que la colonne 0 est le texte, le reste sont des chiffres
    # On itère à partir de la colonne 1
    for col in df.columns[1:]:
        df[col] = df[col].apply(clean_number)
        
    return df

def merge_section_tables(tables_list: List[pd.DataFrame], section_name: str) -> pd.DataFrame:
    """
    Fusionne plusieurs DataFrames (pages coupées) en un seul tableau continu.
    Supprime les en-têtes répétitifs (lignes contenant 'Tableau', 'Brut', 'Net', etc.)
    """
    if not tables_list:
        return pd.DataFrame()
    
    final_df = tables_list[0].copy()
    
    for next_df in tables_list[1:]:
        # Filtrage intelligent : on ignore les lignes qui ressemblent à des en-têtes répétés
        # On vérifie la première colonne. Si elle contient des mots-clés d'en-tête, on saute la ligne.
        keywords_to_skip = ['Tableau', 'Brut', 'Amortissements', 'Net', 'Exercice', 'DESIGNATION', 'ACTIF', 'PASSIF']
        
        mask = ~next_df.iloc[:, 0].astype(str).str.contains('|'.join(keywords_to_skip), na=False, case=False)
        clean_next = next_df[mask]
        
        if not clean_next.empty:
            final_df = pd.concat([final_df, clean_next], ignore_index=True)
            
    return final_df

def convert(pdf_path: str, output_path: str) -> Dict[str, Any]:
    """
    Fonction principale d'extraction et de conversion.
    """
    stats = {
        "info": {},
        "tables": 0,
        "rows": 0,
        "pages": 0
    }
    
    all_sheets = {}
    
    # Listes temporaires pour stocker les chunks par section
    bilan_actif_chunks = []
    bilan_passif_chunks = []
    cpc_chunks = []
    annexes_chunks = []
    
    current_section = None
    
    try:
        with pdfplumber.open(pdf_path) as pdf:
            stats["pages"] = len(pdf.pages)
            stats["info"] = extract_metadata(pdf.pages)
            
            for i, page in enumerate(pdf.pages):
                text = page.extract_text()
                if not text:
                    continue
                
                # --- Détection de la section courante ---
                text_upper = text.upper()
                
                if "BILAN ACTIF" in text_upper or ("ACTIF" in text_upper and "BILAN" in text_upper):
                    current_section = "ACTIF"
                elif "BILAN PASSIF" in text_upper or ("PASSIF" in text_upper and "BILAN" in text_upper):
                    current_section = "PASSIF"
                elif "COMPTE DE PRODUITS ET CHARGES" in text_upper or "CPC" in text_upper:
                    current_section = "CPC"
                elif "ANNEXES" in text_upper or "PIÈCES ANNEXES" in text_upper:
                    current_section = "ANNEXES"
                
                # Extraction des tableaux bruts de la page
                tables = page.extract_tables()
                
                for table in tables:
                    if len(table) < 2: # Ignorer les tableaux trop petits (moins de 2 lignes)
                        continue
                    
                    df = pd.DataFrame(table)
                    df_clean = process_dataframe(df)
                    
                    if df_clean.empty:
                        continue
                    
                    # Routage vers la bonne liste
                    if current_section == "ACTIF":
                        bilan_actif_chunks.append(df_clean)
                    elif current_section == "PASSIF":
                        bilan_passif_chunks.append(df_clean)
                    elif current_section == "CPC":
                        cpc_chunks.append(df_clean)
                    elif current_section == "ANNEXES":
                        annexes_chunks.append(df_clean)
                    else:
                        # Si aucune section détectée mais qu'on a un tableau d'identification
                        first_cell = str(table[0][0]).upper() if table and table[0] else ""
                        if "IDENTIFICATION" in first_cell or "RAISON SOCIALE" in first_cell:
                            if "Metadata" not in all_sheets:
                                all_sheets["Metadata"] = df_clean

            # --- Fusion et Finalisation ---
            
            # 1. Bilan Actif
            if bilan_actif_chunks:
                df_actif = merge_section_tables(bilan_actif_chunks, "Actif")
                all_sheets["Bilan_Actif"] = df_actif
            
            # 2. Bilan Passif
            if bilan_passif_chunks:
                df_passif = merge_section_tables(bilan_passif_chunks, "Passif")
                all_sheets["Bilan_Passif"] = df_passif
            
            # 3. CPC (Compte de Produits et Charges)
            if cpc_chunks:
                df_cpc = merge_section_tables(cpc_chunks, "CPC")
                all_sheets["CPC"] = df_cpc
            
            # 4. Annexes
            if annexes_chunks:
                df_annexes = merge_section_tables(annexes_chunks, "Annexes")
                all_sheets["Annexes"] = df_annexes

    except Exception as e:
        raise Exception(f"Erreur lors de la lecture du PDF : {str(e)}")

    if not all_sheets:
        raise Exception("Aucun tableau financier valide détecté dans ce PDF.")

    # --- Écriture dans Excel avec XlsxWriter ---
    try:
        with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
            workbook = writer.book
            
            # Formats communs
            header_format = workbook.add_format({'bold': True, 'bg_color': '#D9E1F2', 'border': 1, 'align': 'center'})
            number_format = workbook.add_format({'num_format': '#,##0.00', 'border': 1})
            text_format = workbook.add_format({'border': 1})
            
            for sheet_name, df in all_sheets.items():
                # Nettoyage du nom de la feuille (max 31 chars, pas de caractères interdits)
                clean_name = re.sub(r'[\\/\[\]:*?]', '', sheet_name)[:31]
                
                # Écriture des données
                df.to_excel(writer, sheet_name=clean_name, index=False, startrow=1)
                
                worksheet = writer.sheets[clean_name]
                
                # Application des formats
                for col_num, value in enumerate(df.columns.values):
                    worksheet.write(0, col_num, value, header_format)
                    
                    # Ajustement largeur colonne
                    max_len = max(df[col_num].astype(str).map(len).max(), len(str(value))) + 2
                    worksheet.set_column(col_num, col_num, min(max_len, 50))
                    
                    # Formatage des cellules de données (sauf la colonne 0 qui est du texte)
                    if col_num > 0:
                        # On applique le format nombre à toute la colonne sauf l'en-tête
                        worksheet.set_column(col_num, col_num, min(max_len, 20), number_format)

        # Mise à jour des statistiques
        stats["tables"] = len(all_sheets)
        stats["rows"] = sum(len(df) for df in all_sheets.values())
        
        return stats
        
    except Exception as e:
        raise Exception(f"Erreur lors de l'écriture du fichier Excel : {str(e)}")
