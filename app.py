"""FiscalXL — Convertisseur PDF fiscal → Excel structuré (MCN loi 9-88)"""
import streamlit as st
import tempfile, os
from utils.logger import get_logger

logger = get_logger(__name__)

st.set_page_config(page_title="FiscalXL", page_icon="📊", layout="wide")
st.markdown("""<style>
.hdr{background:linear-gradient(135deg,#1F3864,#2E75B6);padding:1.4rem 2rem;
  border-radius:12px;margin-bottom:1.2rem;}
.hdr h1{color:white;margin:0;font-size:1.8rem;}
.hdr p{color:#BDD7EE;margin:.3rem 0 0;}
.kpi{background:white;border:1px solid #BDD7EE;border-radius:8px;
  padding:.8rem;text-align:center;}
.kpi .v{font-size:1.1rem;font-weight:bold;color:#1F3864;}
.kpi .l{font-size:.72rem;color:#888;margin-top:.3rem;}
.ok{background:#E2EFDA;border:1px solid #70AD47;border-radius:8px;
  padding:.9rem 1.3rem;color:#375623;margin:.5rem 0;}
.er{background:#FCE4D6;border:1px solid #C55A11;border-radius:8px;
  padding:.9rem 1.3rem;color:#7B2C00;}
div[data-testid="stDownloadButton"] button{
  background:linear-gradient(135deg,#1F3864,#2E75B6);color:white;
  border:none;padding:.8rem 2.5rem;font-size:1rem;border-radius:8px;width:100%;}
</style>""", unsafe_allow_html=True)

st.markdown("""<div class="hdr">
<h1>📊 FiscalXL</h1>
<p>Convertisseur PDF → Excel structuré · MCN loi 9-88 Maroc · Structure fixe garantie</p>
</div>""", unsafe_allow_html=True)

# ── Sidebar ──────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("### 📋 Format PDF")
    fmt_choice = st.radio("", ["📄 AMMC — 5 pages", "🏛️ DGI — 7 pages"], index=0)
    is_dgi = "DGI" in fmt_choice
    fmt_label = "DGI" if is_dgi else "AMMC"

    st.markdown("---")
    st.markdown(f"""
**Structure Excel générée :**
- `1 - Identification`
- `2 - Bilan Actif` → **49 lignes**
- `3 - Bilan Passif` → **44 lignes**
- `4 - CPC` → **54 lignes**

_Identique pour tous les PDFs {fmt_label}_
""")
    st.caption("FiscalXL · MCN loi 9-88")

# ── Upload ────────────────────────────────────────────────────────────────────
st.markdown("### 📂 Importer le PDF")
uploaded = st.file_uploader(
    f"PDF {fmt_label} ({'5 pages' if not is_dgi else '7 pages'})",
    type=["pdf"])

if not uploaded:
    st.markdown(f"""<div style="text-align:center;padding:3rem;color:#888;
      border:2px dashed #BDD7EE;border-radius:12px;background:#f8fafd;">
      <div style="font-size:3rem;">📄</div>
      <h3 style="color:#2E75B6;">Importez un PDF {fmt_label}</h3>
      <p>Format : {'5 pages (Actif/Passif/CPC×2)' if not is_dgi else '7 pages (Actif×2/Passif/CPC×3)'}</p>
    </div>""", unsafe_allow_html=True)
    st.stop()

st.markdown("---")

with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
    tmp.write(uploaded.getbuffer())
    pdf_path = tmp.name
output_path = pdf_path.replace(".pdf", "_out.xlsx")

try:
    progress = st.progress(0)
    status   = st.empty()
    status.info("📄 Extraction en cours...")
    progress.progress(25)

    if is_dgi:
        from core.dgi_parser import parse
    else:
        from core.ammc_parser import parse
    from core.excel_writer import write

    with st.spinner("Analyse du PDF..."):
        parsed = parse(pdf_path)
    progress.progress(65)

    status.info("📊 Génération Excel...")
    with st.spinner("Écriture Excel..."):
        stats = write(parsed, output_path)
    progress.progress(100)
    status.empty()

    info    = parsed['info']
    raison  = (info.get('raison_sociale') or '—')[:28]
    exercice = info.get('exercice_fin') or '—'

    cols = st.columns(4)
    for col, (lbl, val) in zip(cols, [
        ("Raison Sociale", raison),
        ("Fin exercice",   exercice),
        ("Format",         stats['format']),
        ("Lignes Excel",   f"{stats['rows']} (fixe)"),
    ]):
        col.markdown(f'<div class="kpi"><div class="v">{val}</div>'
                     f'<div class="l">{lbl}</div></div>', unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown(f"""<div class="ok">
      ✅ <strong>Excel généré !</strong>
      &nbsp;·&nbsp; Format <strong>{stats['format']}</strong>
      &nbsp;·&nbsp; <strong>{stats['actif']}</strong> lignes Actif
      &nbsp;·&nbsp; <strong>{stats['passif']}</strong> lignes Passif
      &nbsp;·&nbsp; <strong>{stats['cpc']}</strong> lignes CPC
    </div>""", unsafe_allow_html=True)

    fname = f"{raison.replace(' ','_')[:20]}_{exercice.replace('/','_')}_{stats['format']}.xlsx"
    with open(output_path, "rb") as f:
        st.download_button(
            "📥 Télécharger le fichier Excel", data=f,
            file_name=fname,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

except Exception as e:
    logger.exception("Erreur")
    st.markdown(f'<div class="er">❌ <strong>Erreur :</strong> <code>{e}</code></div>',
                unsafe_allow_html=True)
    import traceback; st.code(traceback.format_exc())
finally:
    for f in [pdf_path, output_path]:
        try:
            if os.path.exists(f): os.unlink(f)
        except: pass
