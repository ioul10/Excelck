"""
FiscalXL — Convertisseur PDF Fiscal → Excel Pro
Point d'entrée Streamlit
"""

import streamlit as st
import tempfile
import os
from pathlib import Path

from core.extractor import PDFExtractor
from core.transformer import FiscalTransformer
from core.excel_builder import ExcelBuilder
from utils.validator import validate_pdf_structure
from utils.logger import get_logger

logger = get_logger(__name__)

# ── Config page ──────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="FiscalXL — PDF → Excel",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── CSS personnalisé ─────────────────────────────────────────────────────────
st.markdown("""
<style>
    .main-header {
        background: linear-gradient(135deg, #1F3864 0%, #2E75B6 100%);
        padding: 2rem 2.5rem;
        border-radius: 12px;
        margin-bottom: 2rem;
        color: white;
    }
    .main-header h1 { color: white; margin: 0; font-size: 2.2rem; }
    .main-header p  { color: #BDD7EE; margin: 0.3rem 0 0; font-size: 1rem; }

    .step-card {
        background: #f8f9fa;
        border-left: 4px solid #2E75B6;
        padding: 1rem 1.2rem;
        border-radius: 0 8px 8px 0;
        margin: 0.5rem 0;
    }
    .step-card h4 { margin: 0 0 0.3rem; color: #1F3864; }
    .step-card p  { margin: 0; color: #555; font-size: 0.9rem; }

    .stat-box {
        background: white;
        border: 1px solid #BDD7EE;
        border-radius: 8px;
        padding: 1rem;
        text-align: center;
    }
    .stat-box .value { font-size: 1.6rem; font-weight: bold; color: #1F3864; }
    .stat-box .label { font-size: 0.8rem; color: #888; margin-top: 0.2rem; }

    .success-banner {
        background: #E2EFDA;
        border: 1px solid #70AD47;
        border-radius: 8px;
        padding: 1rem 1.5rem;
        color: #375623;
    }
    .error-banner {
        background: #FCE4D6;
        border: 1px solid #C55A11;
        border-radius: 8px;
        padding: 1rem 1.5rem;
        color: #7B2C00;
    }
    [data-testid="stFileUploader"] { border: 2px dashed #2E75B6 !important; border-radius: 10px; }
    div[data-testid="stDownloadButton"] button {
        background: linear-gradient(135deg, #1F3864, #2E75B6);
        color: white;
        border: none;
        padding: 0.7rem 2rem;
        font-size: 1rem;
        border-radius: 8px;
        width: 100%;
    }
</style>
""", unsafe_allow_html=True)

# ── Header ───────────────────────────────────────────────────────────────────
st.markdown("""
<div class="main-header">
    <h1>📊 FiscalXL</h1>
    <p>Convertisseur automatique · Pièces annexes Déclaration IS → Excel structuré avec formules</p>
</div>
""", unsafe_allow_html=True)

# ── Sidebar ──────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("### ⚙️ Options")
    
    # NOUVEAU : Mode de génération
    generation_mode = st.radio(
        "Mode de génération",
        ["Excel avec formules (nouveau)", "Remplir template existant"],
        help="Choisis entre créer un nouvel Excel ou remplir ton template"
    )
    
    if generation_mode == "Remplir template existant":
        template_file = st.file_uploader(
            "📁 Template Excel",
            type=["xlsx"],
            help="Ton fichier Excel template à remplir"
        )
    
    opt_dashboard = st.toggle("Feuille Tableau de Bord", value=True)
    st.image("https://img.icons8.com/color/96/microsoft-excel-2019.png", width=60)
    st.markdown("### ⚙️ Options")
    opt_formulas  = st.toggle("Formules dynamiques",     value=True)
    opt_colors    = st.toggle("Mise en forme colorée",   value=True)

    st.markdown("---")
    st.markdown("### 📋 Structure détectée")
    st.markdown("""
    Le PDF doit contenir :
    - **Page 1** — Infos générales
    - **Page 2-3** — Bilan Actif
    - **Page 3-4** — Bilan Passif
    - **Page 4-5** — CPC
    """)

    st.markdown("---")
    st.markdown("### ℹ️ À propos")
    st.caption("FiscalXL v1.0 · Modèle Comptable Normal (loi 9-88)")

# ── Layout principal ─────────────────────────────────────────────────────────
col_upload, col_info = st.columns([3, 2])

with col_upload:
    st.markdown("### 📂 Importer le PDF")
    uploaded = st.file_uploader(
        "Glissez-déposez ou cliquez pour choisir",
        type=["pdf"],
        help="PDF de pièces annexes à la déclaration IS (Modèle Normal)",
    )

with col_info:
    st.markdown("### 🔄 Étapes de traitement")
    steps = [
        ("1 · Extraction",   "Lecture du PDF avec pdfplumber"),
        ("2 · Structuration","Détection des tableaux et des zones"),
        ("3 · Relations",    "Reconstruction des formules & totaux"),
        ("4 · Excel Pro",    "Génération du classeur multi-feuilles"),
    ]
    for title, desc in steps:
        st.markdown(f"""
        <div class="step-card">
            <h4>{title}</h4>
            <p>{desc}</p>
        </div>""", unsafe_allow_html=True)

# ── Traitement ───────────────────────────────────────────────────────────────
if uploaded:
    st.markdown("---")

    # Sauvegarder temporairement
    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
        tmp.write(uploaded.getbuffer())
        tmp_path = tmp.name

    try:
        # ── Progression ──
        progress = st.progress(0, text="Initialisation...")
        status   = st.empty()

        # Étape 1 : Validation
        status.info("🔍 **Étape 1/4** — Validation de la structure PDF...")
        progress.progress(10, text="Validation structure...")

        with st.spinner("Lecture du PDF..."):
            extractor = PDFExtractor(tmp_path)
            validation = validate_pdf_structure(extractor)

        if not validation["valid"]:
            st.markdown(f"""
            <div class="error-banner">
                ⚠️ <strong>Structure non reconnue</strong><br>
                {validation['message']}
            </div>""", unsafe_allow_html=True)
            st.stop()

        progress.progress(25, text="Structure validée ✓")

        # Afficher méta-données détectées
        meta = validation["meta"]
        c1, c2, c3, c4 = st.columns(4)
        for col, (label, val) in zip(
            [c1, c2, c3, c4],
            [("Raison Sociale", meta.get("raison_sociale", "—")[:22]),
             ("Identifiant Fiscal", meta.get("identifiant_fiscal", "—")),
             ("Exercice", meta.get("exercice", "—")),
             ("Pages détectées", str(meta.get("pages", "—")))]):
            col.markdown(f"""
            <div class="stat-box">
                <div class="value">{val}</div>
                <div class="label">{label}</div>
            </div>""", unsafe_allow_html=True)

        st.markdown("<br>", unsafe_allow_html=True)

        # Étape 2 : Extraction
        status.info("📄 **Étape 2/4** — Extraction des données...")
        progress.progress(40, text="Extraction des tableaux...")

        with st.spinner("Extraction en cours..."):
            data = extractor.extract_all()

        progress.progress(60, text="Extraction complète ✓")

        # Étape 3 : Transformation
        status.info("🔗 **Étape 3/4** — Reconstruction des relations...")
        progress.progress(70, text="Calcul des relations...")

        with st.spinner("Structuration des données..."):
            transformer = FiscalTransformer(data)
            fiscal_data = transformer.transform()

        progress.progress(80, text="Relations reconstruites ✓")

        # Étape 4 : Excel
        status.info("📊 **Étape 4/4** — Génération du fichier Excel...")
        progress.progress(88, text="Génération Excel...")

        output_path = tmp_path.replace(".pdf", ".xlsx")

        with st.spinner("Construction du classeur Excel..."):
            if generation_mode == "Excel avec formules (nouveau)":
                # Mode actuel
                builder = ExcelBuilder(
                    fiscal_data,
                    with_dashboard=opt_dashboard,
                    with_formulas=opt_formulas,
                    with_colors=opt_colors,
                )
                stats = builder.build(output_path)
            else:
               # NOUVEAU : Remplir template
               if template_file:
                  # Sauvegarder temporairement le template
                  with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp_tpl:
                       tmp_tpl.write(template_file.getbuffer())
                       template_path = tmp_tpl.name
            
                  filler = TemplateFiller(template_path)
                  stats = filler.fill_from_data(fiscal_data, output_path)
            
                  # Nettoyage
                  os.unlink(template_path)
               else:
                  st.error("Veuillez uploader un template Excel")
                  st.stop()
        progress.progress(100, text="✅ Terminé !")
        status.empty()

        # ── Succès ──
        st.markdown(f"""
        <div class="success-banner">
            ✅ <strong>Fichier Excel généré avec succès !</strong>
            &nbsp;·&nbsp; {stats['sheets']} feuilles
            &nbsp;·&nbsp; {stats['formulas']} formules
            &nbsp;·&nbsp; {stats['rows']} lignes de données
        </div>""", unsafe_allow_html=True)

        st.markdown("<br>", unsafe_allow_html=True)

        # ── Prévisualisation ──
        with st.expander("👁️ Aperçu des données extraites", expanded=False):
            tabs = st.tabs(["Infos", "Bilan Actif", "Bilan Passif", "CPC"])

            with tabs[0]:
                st.json(fiscal_data.get("info", {}))

            with tabs[1]:
                actif = fiscal_data.get("bilan_actif", [])
                if actif:
                    import pandas as pd
                    df = pd.DataFrame(actif,
                        columns=["Poste", "Brut", "Amort.", "Net N", "Net N-1"])
                    st.dataframe(df, use_container_width=True)

            with tabs[2]:
                passif = fiscal_data.get("bilan_passif", [])
                if passif:
                    import pandas as pd
                    df = pd.DataFrame(passif, columns=["Poste", "Exercice N", "Exercice N-1"])
                    st.dataframe(df, use_container_width=True)

            with tabs[3]:
                cpc = fiscal_data.get("cpc", [])
                if cpc:
                    import pandas as pd
                    df = pd.DataFrame(cpc,
                        columns=["Désignation", "Propre N", "Exerc. Préc.", "Total N", "Total N-1"])
                    st.dataframe(df, use_container_width=True)

        # ── Bouton de téléchargement ──
        st.markdown("### ⬇️ Télécharger")
        fname = Path(uploaded.name).stem + "_fiscal.xlsx"
        with open(output_path, "rb") as f:
            st.download_button(
                label="📥 Télécharger le fichier Excel",
                data=f,
                file_name=fname,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

    except Exception as e:
        logger.exception("Erreur lors du traitement")
        st.markdown(f"""
        <div class="error-banner">
            ❌ <strong>Erreur de traitement</strong><br>
            <code>{str(e)}</code>
        </div>""", unsafe_allow_html=True)
        with st.expander("Détails de l'erreur"):
            import traceback
            st.code(traceback.format_exc())

    finally:
        # Nettoyage
        for f in [tmp_path, tmp_path.replace(".pdf", ".xlsx")]:
            if os.path.exists(f):
                try:
                    os.unlink(f)
                except Exception:
                    pass

else:
    # État vide
    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown("""
    <div style="text-align:center; padding: 3rem; color: #888; border: 2px dashed #BDD7EE; border-radius: 12px; background: #f8fafd;">
        <div style="font-size: 3rem;">📄</div>
        <h3 style="color: #2E75B6;">Importez un PDF pour commencer</h3>
        <p>Pièces annexes à la déclaration IS — Modèle Comptable Normal</p>
    </div>
    """, unsafe_allow_html=True)
