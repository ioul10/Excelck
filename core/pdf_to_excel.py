import pdfplumber
import pandas as pd
import re
from datetime import datetime
import os

def clean_number(val):
    """Nettoie une chaîne pour en faire un nombre flottant valide."""
    if val is None:
        return None
    if isinstance(val, (int, float)):
        return val
    
    text = str(val).strip()
    if not text or text == "-":
        return None
    
    # Supprime les espaces (séparateurs de milliers) et remplace la virgule décimale par un point
    # Ex: "1 234,56" -> "1234.56"
    text = text.replace(" ", "").replace(",", ".")
    
    try:
        return float(text)
    except ValueError:
        return text  # Retourne le texte original si ce n'est pas un nombre

def extract_metadata(pages):
    """Extrait les métadonnées (Raison sociale, Exercice) depuis la première page."""
    info = {
        "raison_sociale": "Inconnu",
        "exercice_debut": "",
        "exercice_fin": "",
        "identifiant_fiscal": ""
    }
    
    # Recherche simple dans le texte brut de la première page
    text_page_0 = pages[0].extract_text() if pages else ""
    
    # Regex pour Raison Sociale (souvent après "Raison sociale" ou en gros titre)
    match_rs = re.search(r'Raison sociale\s*[:\n]?\s*(.+?)(?:\n|Identifiant)', text_page_0, re.IGNORECASE)
    if match_rs:
        info["raison_sociale"] = match_rs.group(1).strip()
    
    # Regex pour Identifiant Fiscal
    match_if = re.search(r'Identifiant fiscal\s*[:\n]?\s*(\d+)', text_page_0, re.IGNORECASE)
    if match_if:
        info["identifiant_fiscal"] = match_if.group(1)

    # Regex pour Exercice (Du ... au ...)
    match_ex = re.search(r'Du\s+(\d{2}/\d{2}/\d{4})\s+au\s+(\d{2}/\d{2}/\d{4})', text_page_0)
    if match_ex:
        info["exercice_debut"] = match_ex.group(1)
        info["exercice_fin"] = match_ex.group(2)
    else:
        # Fallback simple sur l'année
        match_year = re.search(r'(\d{4})', text_page_0)
        if match_year:
            info["exercice_fin"] = f"31/12/{match_year.group(1)}"

    return info

def process_table(table_data):
    """Nettoie un tableau extrait : convertit les nombres et supprime les lignes vides."""
    if not table_data:
        return pd.DataFrame()
    
    df = pd.DataFrame(table_data)
    
    # Suppression des lignes entièrement vides
    df = df.dropna(how='all')
    
    # Nettoyage des colonnes numériques (on suppose que tout sauf la première colonne est numérique dans les bilans)
    # Attention : cela dépend de la structure exacte, ici on essaie de convertir toutes les colonnes sauf la 0 (Libellé)
    for col in df.columns[1:]:
        df[col] = df[col].apply(clean_number)
        
    return df

def merge_cpc_tables(tables_list):
    """Fusionne les tableaux CPC (souvent coupés en 1/2 et 2/2) en un seul tableau continu."""
    if not tables_list:
        return pd.DataFrame()
    
    # On prend le premier tableau comme base
    final_df = tables_list[0].copy()
    
    for next_df in tables_list[1:]:
        # On ignore les lignes d'en-tête répétées (qui contiennent souvent "Tableau n°" ou "DESIGNATION")
        # On garde seulement les lignes de données réelles
        # Astuce : Si la première colonne contient "RESULTAT COURANT" ou similaire, c'est la suite
        
        # Filtrage grossier : on saute les lignes qui ressemblent à des titres de section répétés
        mask = ~next_df.iloc[:, 0].astype(str).str.contains(r'Tableau|Compte de Produits|DESIGNATION', na=False)
        clean_next = next_df[mask]
        
        if not clean_next.empty:
            final_df = pd.concat([final_df, clean_next], ignore_index=True)
            
    return final_df

def convert(pdf_path, output_path):
    stats = {
        "info": {},
        "tables": 0,
        "rows": 0,
        "pages": 0
    }
    
    all_sheets = {}
    
    with pdfplumber.open(pdf_path) as pdf:
        stats["pages"] = len(pdf.pages)
        stats["info"] = extract_metadata(pdf.pages)
        
        # Listes temporaires pour regrouper les sections
        bilan_actif_chunks = []
        bilan_passif_chunks = []
        cpc_chunks = []
        annexes_chunks = []
        
        current_section = None
        
        for i, page in enumerate(pdf.pages):
            text = page.extract_text()
            if not text:
                continue
                
            # Détection de la section courante basée sur les mots-clés
            if "BILAN ACTIF" in text.upper():
                current_section = "ACTIF"
            elif "BILAN PASSIF" in text.upper():
                current_section = "PASSIF"
            elif "COMPTE DE PRODUITS ET CHARGES" in text.upper() or "CPC" in text.upper():
                current_section = "CPC"
            elif "ANNEXES" in text.upper() or "PIÈCES ANNEXES" in text.upper():
                current_section = "ANNEXES"
            
            # Extraction des tableaux de la page
            tables = page.extract_tables()
            
            for table in tables:
                if len(table) < 2: # Ignorer les tableaux trop petits
                    continue
                    
                df = process_table(table)
                if df.empty:
                    continue
                
                # Routage vers la bonne section
                if current_section == "ACTIF":
                    bilan_actif_chunks.append(df)
                elif current_section == "PASSIF":
                    bilan_passif_chunks.append(df)
                elif current_section == "CPC":
                    cpc_chunks.append(df)
                elif current_section == "ANNEXES":
                    annexes_chunks.append(df)
                else:
                    # Si aucune section détectée mais qu'on a un tableau avec "Identification", on le met dans Info
                    first_cell = str(table[0][0]).upper() if table and table[0] else ""
                    if "IDENTIFICATION" in first_cell or "RAISON SOCIALE" in first_cell:
                        all_sheets["Metadata"] = df

        # Fusion et nettoyage final par section
        
        # 1. Bilan Actif
        if bilan_actif_chunks:
            # Concaténation simple, puis nettoyage des doublons d'en-têtes si nécessaire
            df_actif = pd.concat(bilan_actif_chunks, ignore_index=True)
            # Supprimer les lignes où la première colonne est vide ou contient juste des totaux intermédiaires mal placés
            # (Optionnel : affiner selon besoin)
            all_sheets["Bilan_Actif"] = df_actif
            
        # 2. Bilan Passif
        if bilan_passif_chunks:
            df_passif = pd.concat(bilan_passif_chunks, ignore_index=True)
            all_sheets["Bilan_Passif"] = df_passif
            
        # 3. CPC (Spécial : fusion intelligente)
        if cpc_chunks:
            df_cpc = merge_cpc_tables(cpc_chunks)
            all_sheets["CPC"] = df_cpc
            
        # 4. Annexes / Autres
        if annexes_chunks:
            df_annexes = pd.concat(annexes_chunks, ignore_index=True)
            all_sheets["Annexes"] = df_annexes

    # Écriture dans Excel
    if not all_sheets:
        raise Exception("Aucun tableau financier détecté dans le PDF.")

    with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
        for sheet_name, df in all_sheets.items():
            # Nommer les feuilles proprement (max 31 chars)
            clean_name = sheet_name[:31]
            df.to_excel(writer, sheet_name=clean_name, index=False)
            
            # Ajustement automatique de la largeur des colonnes (optionnel mais recommandé)
            worksheet = writer.sheets[clean_name]
            for i, col in enumerate(df.columns):
                max_len = max(df[col].astype(str).map(len).max(), len(str(col))) + 2
                worksheet.set_column(i, i, max_len)

    # Mise à jour des stats
    stats["tables"] = len(all_sheets)
    stats["rows"] = sum(len(df) for df in all_sheets.values())
    
    return stats
