"""FiscalXL v7 — TableParser + injector_v7"""
import streamlit as st
import tempfile, os
from pathlib import Path
import pandas as pd

from core.table_parser import TableParser
from core.injector_v7  import inject
from utils.validator   import validate_pdf_structure_v2
from core.pdf_parser   import PDFParser
from utils.logger      import get_logger

logger = get_logger(__name__)
TEMPLATE = Path(__file__).parent / "EX_template.xlsx"

st.set_page_config(page_title="FiscalXL v7", page_icon="📊", layout="wide",
                   initial_sidebar_state="expanded")
st.markdown("""<style>
.hdr{background:linear-gradient(135deg,#1F3864,#2E75B6);padding:1.4rem 2rem;
  border-radius:12px;margin-bottom:1.2rem;}
.hdr h1{color:white;margin:0;font-size:1.8rem;}
.hdr p{color:#BDD7EE;margin:.3rem 0 0;font-size:.88rem;}
.kpi{background:white;border:1px solid #BDD7EE;border-radius:8px;padding:.7rem;text-align:center;}
.kpi .v{font-size:1.1rem;font-weight:bold;color:#1F3864;}
.kpi .l{font-size:.72rem;color:#888;margin-top:.2rem;}
.ok{background:#E2EFDA;border:1px solid #70AD47;border-radius:8px;padding:.9rem 1.3rem;color:#375623;}
.er{background:#FCE4D6;border:1px solid #C55A11;border-radius:8px;padding:.9rem 1.3rem;color:#7B2C00;}
div[data-testid="stDownloadButton"] button{
  background:linear-gradient(135deg,#1F3864,#2E75B6);color:white;border:none;
  padding:.8rem 2.5rem;font-size:1rem;border-radius:8px;width:100%;}
</style>""", unsafe_allow_html=True)

st.markdown("""<div class="hdr"><h1>📊 FiscalXL v7</h1>
<p>Convertisseur PDF → Excel · Pièces annexes IS — MCN loi 9-88 Maroc</p></div>""",
unsafe_allow_html=True)

with st.sidebar:
    st.markdown("### ⚙️ Options")
    show_preview = st.toggle("Aperçu des données", value=True)
    show_debug   = st.toggle("Mode débogage",      value=False)
    st.markdown("---")
    st.markdown("""**Méthode :**
- Bordures détectées → `extract_tables()`
- Sinon → Algorithme X/Y
- Injection directe par label exact""")
    st.markdown("---")
    st.success("✅ Template chargé") if TEMPLATE.exists() else st.error("⚠️ Template manquant")
    st.caption("FiscalXL v7 · MCN loi 9-88")

col_up, col_pipe = st.columns([3, 2])
with col_up:
    st.markdown("### 📂 Importer le PDF")
    uploaded = st.file_uploader("Glissez-déposez ou cliquez", type=["pdf"])
with col_pipe:
    st.markdown("### 🔄 Pipeline")
    for step, desc in [
        ("1 · Validation",  "Structure du PDF"),
        ("2 · Extraction",  "Tableaux → données"),
        ("3 · Injection",   "Label → cellule Excel"),
        ("4 · Formules",    "Totaux calculés auto"),
    ]:
        st.markdown(f'<div style="background:#f8f9fa;border-left:4px solid #1F3864;'
                    f'padding:.5rem .8rem;border-radius:0 6px 6px 0;margin:.3rem 0;">'
                    f'<strong style="color:#1F3864">{step}</strong><br>'
                    f'<span style="color:#555;font-size:.83rem">{desc}</span></div>',
                    unsafe_allow_html=True)

if not uploaded:
    st.markdown("""<div style="text-align:center;padding:3rem;color:#888;
      border:2px dashed #BDD7EE;border-radius:12px;background:#f8fafd;margin-top:1rem;">
      <div style="font-size:3rem;">📄</div>
      <h3 style="color:#2E75B6;">Importez un PDF pour commencer</h3>
      <p>DGI (7 pages) · AMMC (5 pages) · Tous formats MCN</p></div>""",
      unsafe_allow_html=True)
    st.stop()

st.markdown("---")
if not TEMPLATE.exists():
    st.markdown('<div class="er">❌ EX_template.xlsx manquant.</div>', unsafe_allow_html=True)
    st.stop()

with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
    tmp.write(uploaded.getbuffer())
    pdf_path = tmp.name
output_path = pdf_path.replace(".pdf", "_out.xlsx")

try:
    progress = st.progress(0)
    status   = st.empty()

    # Validation
    status.info("🔍 Validation...")
    progress.progress(15)
    base = PDFParser(pdf_path)
    v    = validate_pdf_structure_v2(base)
    if not v["valid"]:
        st.markdown(f'<div class="er">⚠️ {v["message"]}</div>', unsafe_allow_html=True)
        st.stop()
    meta = v["meta"]

    for col, (lbl, val) in zip(st.columns(4), [
        ("Raison Sociale",    (meta.get("raison_sociale") or "—")[:22]),
        ("Identifiant Fiscal", meta.get("identifiant_fiscal") or "—"),
        ("Fin exercice",       meta.get("exercice_fin") or "—"),
        ("Pages",              str(meta.get("pages", "—"))),
    ]):
        col.markdown(f'<div class="kpi"><div class="v">{val}</div>'
                     f'<div class="l">{lbl}</div></div>', unsafe_allow_html=True)
    st.markdown("<br>", unsafe_allow_html=True)
    progress.progress(25)

    # Extraction
    status.info("📄 Extraction des tableaux...")
    progress.progress(45)
    with st.spinner("Lecture en cours..."):
        parser    = TableParser(pdf_path)
        extracted = parser.parse()
    progress.progress(65)

    if show_debug:
        with st.expander("🐛 Données brutes"):
            for sec, d in extracted.items():
                if isinstance(d, dict):
                    st.write(f"**{sec}** ({len(d)} postes)")
                    st.json({k: v for k,v in list(d.items())[:8]})

    # Injection
    status.info("🔗 Injection dans le template Excel...")
    progress.progress(78)
    stats = inject(extracted, str(TEMPLATE), output_path)
    progress.progress(100)
    status.empty()

    st.markdown(f"""<div class="ok">
      ✅ <strong>Fichier Excel généré !</strong>
      &nbsp;·&nbsp; {stats['injected']} valeurs injectées
    </div>""", unsafe_allow_html=True)
    st.markdown("<br>", unsafe_allow_html=True)

    # Aperçu
    if show_preview:
        with st.expander("👁️ Aperçu des valeurs extraites", expanded=False):
            tabs = st.tabs(["ℹ️ Infos", "📋 Actif", "📋 Passif", "📈 CPC"])
            with tabs[0]:
                for k, v in extracted.get("info", {}).items():
                    if k != "pages" and v:
                        st.markdown(f"**{k.replace('_',' ').title()}** : {v}")
            with tabs[1]:
                av = extracted.get("actif_values", {})
                if av:
                    df = pd.DataFrame(
                        [(k, f"{v[0]:,.2f}" if v and v[0] is not None else "—",
                             f"{v[1]:,.2f}" if len(v)>1 and v[1] is not None else "—")
                         for k, v in av.items()],
                        columns=["Poste","Brut","Amort."])
                    st.dataframe(df, use_container_width=True, height=300)
            with tabs[2]:
                pv = extracted.get("passif_values", {})
                if pv:
                    df = pd.DataFrame(
                        [(k, f"{v[0]:,.2f}" if v and v[0] is not None else "—",
                             f"{v[1]:,.2f}" if len(v)>1 and v[1] is not None else "—")
                         for k, v in pv.items()],
                        columns=["Poste","N","N-1"])
                    st.dataframe(df, use_container_width=True, height=300)
            with tabs[3]:
                cv = extracted.get("cpc_values", {})
                if cv:
                    df = pd.DataFrame(
                        [(k, f"{v[0]:,.2f}" if v and v[0] is not None else "—",
                             f"{v[2]:,.2f}" if len(v)>2 and v[2] is not None else "—")
                         for k, v in cv.items()],
                        columns=["Désignation","Propre N","Total N-1"])
                    st.dataframe(df, use_container_width=True, height=300)

    # Téléchargement
    st.markdown("### ⬇️ Télécharger")
    raison = (extracted.get("info",{}).get("raison_sociale") or "fiscal").replace(" ","_")[:20]
    date   = (extracted.get("info",{}).get("exercice_fin") or "").replace("/","-")
    fname  = f"{raison}_{date}_fiscal.xlsx"
    with open(output_path, "rb") as f:
        st.download_button("📥 Télécharger le fichier Excel", data=f,
            file_name=fname,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

except Exception as e:
    logger.exception("Erreur pipeline")
    st.markdown(f'<div class="er">❌ <strong>Erreur :</strong> <code>{e}</code></div>',
                unsafe_allow_html=True)
    if show_debug:
        import traceback; st.code(traceback.format_exc())
finally:
    for f in [pdf_path, output_path]:
        try:
            if os.path.exists(f): os.unlink(f)
        except: pass
