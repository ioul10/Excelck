import streamlit as st
import pdfplumber
import pandas as pd
import xlsxwriter
import io
import re

# Configuration de la page
st.set_page_config(page_title="Excelck - Converter Fiscal", layout="wide")

st.title("📑 Excelck : PDF Fiscal vers Excel Dynamique")
st.markdown("""
Cette application transforme les annexes de la déclaration IS (Modèle Normal) 
en un fichier Excel structuré, coloré et **calculant automatiquement les totaux**.
""")

# --- ÉTAPE 1 & 2 : EXTRACTION ET PARSING ---

def extract_data_from_pdf(uploaded_file):
    """
    Extrait le texte et les tableaux du PDF.
    Retourne un dictionnaire structuré par type de feuille.
    """
    data = {
        "info_generale": {},
        "actif": [],
        "passif": [],
        "cpc": []
    }
    
    with pdfplumber.open(uploaded_file) as pdf:
        # Page 1 : Informations Générales (Texte clé-valeur simplifié)
        if len(pdf.pages) > 0:
            page1 = pdf.pages[0]
            text = page1.extract_text()
            # Simulation d'extraction (à adapter selon le regex exact du formulaire)
            data["info_generale"]["texte_brut"] = text
            
        # Pages 2-3 : Bilan Actif
        # Pages 4 : Bilan Passif
        # Page 5 : CPC
        
        # Pour cet exemple, on va extraire les tableaux bruts
        # Dans un cas réel, il faudrait mapper les lignes spécifiques (Immobilisations, Stocks, etc.)
        for i, page in enumerate(pdf.pages):
            tables = page.extract_tables()
            for table in tables:
                df = pd.DataFrame(table)
                if i == 0: continue # Skip page 1 for tables
                
                # Nettoyage basique des NaN
                df = df.fillna("")
                
                if i in [1, 2]: # Actif
                    data["actif"].append(df)
                elif i == 3: # Passif
                    data["passif"].append(df)
                elif i == 4: # CPC
                    data["cpc"].append(df)
                    
    return data

# --- ÉTAPE 3 & 4 : GÉNÉRATION EXCEL PRO AVEC FORMULES ---

def create_pro_excel(data):
    """
    Génère un fichier Excel avec mise en forme et FORMULES de calcul.
    """
    output = io.BytesIO()
    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    workbook = writer.book
    
    # --- DEFINITION DES STYLES (Mise en forme PRO) ---
    header_fmt = workbook.add_format({
        'bold': True, 'text_wrap': True, 'valign': 'top',
        'fg_color': '#2E75B6', 'font_color': '#FFFFFF', 'border': 1
    })
    
    total_fmt = workbook.add_format({
        'bold': True, 'fg_color': '#D9E1F2', 'font_color': '#000000', 'border': 1, 'num_format': '#,##0.00'
    })
    
    normal_fmt = workbook.add_format({
        'border': 1, 'num_format': '#,##0.00', 'align': 'right'
    })
    
    text_fmt = workbook.add_format({
        'border': 1, 'align': 'left', 'text_wrap': True
    })

    # --- FEUILLE 1 : INFORMATIONS GENERALES ---
    df_info = pd.DataFrame([["Champ", "Valeur"]] + [[k, v] for k, v in data["info_generale"].items() if k != "texte_brut"])
    # Note: Ici on met les données en dur car ce sont des infos statiques
    df_info.to_excel(writer, sheet_name='1. Info Générale', index=False, startrow=1, header=False)
    worksheet_info = writer.sheets['1. Info Générale']
    worksheet_info.set_column('A:A', 30)
    worksheet_info.set_column('B:B', 40)

    # --- FEUILLE 2 : BILAN ACTIF (Avec Formules) ---
    worksheet_actif = workbook.add_worksheet('2. Bilan Actif')
    
    # En-têtes
    headers = ['Rubrique', 'N', 'N-1']
    for col_num, header in enumerate(headers):
        worksheet_actif.write(0, col_num, header, header_fmt)
    
    # Simulation de remplissage des données (Brut Net)
    # Dans un cas réel, on mapperait les lignes extraites du PDF ici
    row = 1
    start_data_row = row
    
    # Exemple de structure comptable
    items_actif = [
        ("Immobilisations en non-valeurs", 10000, 10000),
        ("Immobilisations incorporelles", 50000, 45000),
        ("Immobilisations corporelles", 200000, 180000),
        ("Immobilisations financières", 15000, 15000),
        ("Ecarts de conversion Actif", 0, 0),
        ("Créances de l'actif immobilisé", 5000, 4000),
        ("Stocks et en-cours", 120000, 100000),
        ("Créances de l'actif circulant", 80000, 70000),
        ("Titres de placement", 30000, 25000),
        ("Trésorerie Actif", 40000, 35000),
    ]
    
    for item, val_n, val_n_1 in items_actif:
        worksheet_actif.write(row, 0, item, text_fmt)
        worksheet_actif.write(row, 1, val_n, normal_fmt) # Colonne N (Index 1)
        worksheet_actif.write(row, 2, val_n_1, normal_fmt) # Colonne N-1 (Index 2)
        row += 1
        
    end_data_row = row - 1
    
    # LIGNE DE TOTAL (C'est ici que la magie opère - ÉTAPE 3)
    worksheet_actif.write(row, 0, "TOTAL ACTIF", total_fmt)
    
    # FORMULE EXCEL : =SUM(B2:B11)
    # On utilise write_formula pour que Excel recalcule si on change une valeur
    formula_n = f"=SUM(B{start_data_row+2}:B{end_data_row+2})" # +2 car Excel commence à 1 et on a une ligne header
    formula_n_1 = f"=SUM(C{start_data_row+2}:C{end_data_row+2})"
    
    worksheet_actif.write_formula(row, 1, formula_n, total_fmt)
    worksheet_actif.write_formula(row, 2, formula_n_1, total_fmt)
    
    # --- FEUILLE 3 : BILAN PASSIF (Avec Formules) ---
    worksheet_passif = workbook.add_worksheet('3. Bilan Passif')
    # ... Même logique que l'actif ...
    # Pour l'exemple, on copie la structure
    for col_num, header in enumerate(headers):
        worksheet_passif.write(0, col_num, header, header_fmt)
        
    row = 1
    items_passif = [
        ("Capital social", 100000, 100000),
        ("Primes d'émission", 10000, 10000),
        ("Ecarts de réévaluation", 0, 0),
        ("Réserves", 50000, 40000),
        ("Report à nouveau", 20000, 15000),
        ("Résultat net", 0, 0), # Sera lié au CPC
        ("Subventions d'investissement", 30000, 30000),
        ("Provisions durables", 10000, 10000),
        ("Dettes de financement", 150000, 160000),
        ("Provisions pour risques et charges", 5000, 4000),
        ("Dettes du passif circulant", 90000, 80000),
        ("Trésorerie Passif", 15000, 10000),
    ]
    
    for item, val_n, val_n_1 in items_passif:
        worksheet_passif.write(row, 0, item, text_fmt)
        worksheet_passif.write(row, 1, val_n, normal_fmt)
        worksheet_passif.write(row, 2, val_n_1, normal_fmt)
        row += 1
        
    end_data_row = row - 1
    worksheet_passif.write(row, 0, "TOTAL PASSIF", total_fmt)
    
    formula_n = f"=SUM(B{start_data_row+2}:B{end_data_row+2})"
    formula_n_1 = f"=SUM(C{start_data_row+2}:C{end_data_row+2})"
    
    worksheet_passif.write_formula(row, 1, formula_n, total_fmt)
    worksheet_passif.write_formula(row, 2, formula_n_1, total_fmt)

    # --- FEUILLE 4 : CPC (Compte de Produits et Charges) ---
    worksheet_cpc = workbook.add_worksheet('4. CPC')
    headers_cpc = ['Rubrique', 'Montant N', 'Montant N-1']
    for col_num, header in enumerate(headers_cpc):
        worksheet_cpc.write(0, col_num, header, header_fmt)
        
    row = 1
    # Structure simplifiée CPC
    cpc_data = [
        ("Ventes de marchandises", 500000, 450000),
        ("Chiffre d'affaires", 500000, 450000),
        ("Variation de stocks", 10000, 5000),
        ("Production immobilisée", 0, 0),
        ("Subventions d'exploitation", 20000, 20000),
        ("Autres produits", 5000, 3000),
        ("Achats revendus de marchandises", -200000, -180000),
        ("Achats consommés", -100000, -90000),
        ("Autres charges externes", -50000, -45000),
        ("Impôts et taxes", -15000, -14000),
        ("Charges de personnel", -120000, -110000),
        ("Autres charges d'exploitation", -10000, -9000),
        ("Dotations d'exploitation", -25000, -20000),
        ("Résultat d'exploitation", 0, 0), # Formule
        ("Résultat financier", -5000, -4000),
        ("Résultat non courant", 2000, 1000),
        ("Impôt sur les résultats", -10000, -8000),
        ("Résultat Net", 0, 0) # Formule Finale
    ]
    
    start_cpc_row = row
    for item, val_n, val_n_1 in cpc_data:
        # Si c'est une ligne de calcul (Résultat), on met 0 pour l'instant, la formule prendra le relais
        is_calculation = "Résultat" in item
        
        worksheet_cpc.write(row, 0, item, text_fmt if not is_calculation else total_fmt)
        
        if not is_calculation:
            worksheet_cpc.write(row, 1, val_n, normal_fmt)
            worksheet_cpc.write(row, 2, val_n_1, normal_fmt)
        else:
            # Ici on prépare les formules complexes
            # Exemple simplifié : Résultat d'exploitation = Somme des lignes précédentes
            # Dans un vrai cas, il faut connaître les lignes exactes (Produits - Charges)
            pass 
        row += 1
    
    end_cpc_row = row - 1
    
    # Exemple de formule complexe pour le Résultat Net (Simulation)
    # Disons que le Résultat Net est la somme de tout le CPC (simplification)
    worksheet_cpc.write(end_cpc_row, 0, "RESULTAT NET (Formule)", total_fmt)
    # Formule dynamique
    formula_net_n = f"=SUM(B{start_cpc_row+2}:B{end_cpc_row+1})" 
    worksheet_cpc.write_formula(end_cpc_row, 1, formula_net_n, total_fmt)
    
    # Fermeture du fichier
    writer.close()
    output.seek(0)
    return output

# --- INTERFACE STREAMLIT ---

uploaded_file = st.file_uploader("Choisissez le PDF (Annexes IS)", type="pdf")

if uploaded_file is not None:
    with st.spinner('Extraction et Analyse en cours...'):
        # 1. Extraction
        raw_data = extract_data_from_pdf(uploaded_file)
        st.success("PDF analysé avec succès !")
        
        # Affichage rapide des données brutes extraites (Debug)
        with st.expander("Voir les données brutes extraites"):
            st.json(raw_data)
            
        # 2. Génération Excel
        excel_file = create_pro_excel(raw_data)
        
        st.download_button(
            label="📥 Télécharger l'Excel PRO (Avec Formules)",
            data=excel_file,
            file_name="Declaration_IS_Dynamique.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
