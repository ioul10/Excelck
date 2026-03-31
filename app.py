# app.py
import streamlit as st
import tempfile
import os
import traceback
from pathlib import Path

# Import du module core
from core.pdf_to_excel import convert

# ── Configuration de la page ──────────────────────────────────────────────────
st.set_page_config(
    page_title="📊 Convertisseur Fiscal PDF → Excel",
    page_icon="📑",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ── Styles CSS personnalisés ─────────────────────────────────────────────────
st.markdown("""
<style>
    .stButton>button { width: 100%; }
    .success-box { padding: 1rem; border-radius: 0.5rem; background: #d4edda; border: 1px solid #c3e6cb; }
    .error-box { padding: 1rem; border-radius: 0.5rem; background: #f8d7da; border: 1px solid #f5c6cb; }
    .info-card { padding: 1rem; border-radius: 0.5rem; background: #e7f1fa; border-left: 4px solid #2E75B6; }
</style>
""", unsafe_allow_html=True)

# ── Titre et description ─────────────────────────────────────────────────────
st.title("📑 Convertisseur Fiscal PDF → Excel")
st.markdown("""
**Extraction automatique de documents fiscaux marocains**  
✅ Supporte les formats **AMMC (5 pages)** et **DGI (7 pages)**  
✅ Reconnaissance intelligente des tableaux (avec fallback X/Y)  
✅ Mise en forme professionnelle avec styles, couleurs et formats numériques
""")

# ── Sidebar : Informations et aide ───────────────────────────────────────────
with st.sidebar:
    st.header("ℹ️ Informations")
    st.info("""
    **Formats supportés :**
    - 📄 AMMC : 5 pages (Bilan Actif/Passif + CPC)
    - 📄 DGI : 7 pages (Modèle comptable normal)
    
    **Fonctionnalités :**
    - Extraction via `pdfplumber` (tables + fallback XY)
    - Parsing des nombres au format français
    - Filtrage intelligent des lignes non pertinentes
    - Styles Excel professionnels
    """)
    
    st.divider()
    st.header("📋 Prérequis")
    st.code("""
    pip install streamlit pdfplumber openpyxl
    """, language="bash")
    
    st.divider()
    st.caption("🔐 Les fichiers sont traités localement et supprimés après conversion.")

# ── Zone principale : Upload et conversion ───────────────────────────────────
col1, col2 = st.columns([2, 1])

with col1:
    uploaded_file = st.file_uploader(
        "📁 Importez votre PDF fiscal",
        type=["pdf"],
        help="Glissez-déposez ou cliquez pour sélectionner un fichier PDF (AMMC ou DGI)"
    )

if uploaded_file is not None:
    # Affichage des infos du fichier
    file_info = f"**Fichier :** {uploaded_file.name} • **Taille :** {uploaded_file.size / 1024:.1f} Ko"
    st.markdown(f"<div class='info-card'>{file_info}</div>", unsafe_allow_html=True)
    
    # Aperçu rapide (premières pages)
    with st.expander("🔍 Aperçu du PDF (optionnel)", expanded=False):
        try:
            import pdfplumber
            with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
                tmp.write(uploaded_file.read())
                tmp_path = tmp.name
            
            with pdfplumber.open(tmp_path) as pdf:
                st.write(f"📄 Nombre de pages détectées : **{len(pdf.pages)}**")
                if len(pdf.pages) == 5:
                    st.success("✅ Format AMMC détecté (5 pages)")
                elif len(pdf.pages) == 7:
                    st.success("✅ Format DGI détecté (7 pages)")
                else:
                    st.warning(f"⚠️ Format inattendu : {len(pdf.pages)} pages")
                    
                # Aperçu texte page 1
                st.subheader("Extrait page 1")
                text = pdf.pages[0].extract_text() or "*Aucun texte extrait*"
                st.text_area("", text, height=150, label_visibility="collapsed")
            
            os.unlink(tmp_path)
        except Exception as e:
            st.error(f"Erreur lors de l'aperçu : {e}")

    # Bouton de conversion
    st.divider()
    convert_btn = st.button("🔄 Lancer la conversion", type="primary", disabled=uploaded_file is None)
    
    if convert_btn:
        with st.spinner("⏳ Extraction et conversion en cours..."):
            try:
                # Création d'un fichier temporaire pour le PDF uploadé
                with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf", dir=tempfile.gettempdir()) as tmp_in:
                    tmp_in.write(uploaded_file.read())
                    pdf_path = tmp_in.name
                
                # Chemin de sortie temporaire pour l'Excel
                with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx", dir=tempfile.gettempdir()) as tmp_out:
                    excel_path = tmp_out.name
                
                # 🎯 Appel au module de conversion
                result = convert(pdf_path, excel_path)
                
                # ✅ Succès : affichage des résultats
                st.success("✅ Conversion réussie !")
                
                # Cartes de résumé
                res_col1, res_col2, res_col3 = st.columns(3)
                with res_col1:
                    st.metric("📄 Pages traitées", result['pages'])
                with res_col2:
                    st.metric("📊 Feuilles Excel", result['tables'])
                with res_col3:
                    st.metric("📝 Lignes extraites", result['rows'])
                
                # Détails metadata
                with st.expander("🔎 Métadonnées extraites", expanded=True):
                    info = result['info']
                    meta_col1, meta_col2 = st.columns(2)
                    with meta_col1:
                        st.markdown(f"**Raison sociale :** {info.get('raison_sociale', '—')}")
                        st.markdown(f"**IF :** {info.get('identifiant_fiscal', '—')}")
                    with meta_col2:
                        st.markdown(f"**TP :** {info.get('taxe_professionnelle', '—')}")
                        st.markdown(f"**Exercice :** {info.get('exercice', '—')}")
                
                # 📥 Bouton de téléchargement
                with open(excel_path, "rb") as f:
                    excel_bytes = f.read()
                
                st.download_button(
                    label="📥 Télécharger le fichier Excel",
                    data=excel_bytes,
                    file_name=f"{Path(uploaded_file.name).stem}_converti.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary"
                )
                
                # Nettoyage fichiers temporaires
                os.unlink(pdf_path)
                os.unlink(excel_path)
                
            except FileNotFoundError:
                st.error("❌ Fichier introuvable. Veuillez réessayer.")
            except PermissionError:
                st.error("❌ Erreur de permission. Vérifiez les accès au dossier temporaire.")
            except Exception as e:
                st.error(f"❌ Erreur lors de la conversion : {str(e)}")
                with st.expander("🐞 Détails de l'erreur (debug)"):
                    st.code(traceback.format_exc())

# ── Footer ───────────────────────────────────────────────────────────────────
st.divider()
st.caption("""
💡 **Conseil** : Pour de meilleurs résultats, utilisez des PDF générés numériquement (non scannés).  
🛠️ **Développeur** : Module basé sur `pdfplumber` + `openpyxl` avec fallback X/Y pour tableaux sans bordures.
""")
