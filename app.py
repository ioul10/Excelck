"""FiscalXL v7 — TableParser hybride (extract_tables + X/Y fallback)"""

import streamlit as st
import tempfile, os, shutil
from pathlib import Path
import pandas as pd
import openpyxl

from core.table_parser import TableParser
from core.direct_injector import build_excel_index, ALIASES
from utils.validator import validate_pdf_structure_v2
from core.pdf_parser import PDFParser
from utils.logger import get_logger

logger = get_logger(__name__)
TEMPLATE_PATH = Path(__file__).parent / "EX_template.xlsx"

st.set_page_config(page_title="FiscalXL v7", page_icon="📊", layout="wide",
                   initial_sidebar_state="expanded")

st.markdown("""<style>
.main-header{background:linear-gradient(135deg,#1F3864,#2E75B6);padding:1.5rem 2rem;
  border-radius:12px;margin-bottom:1.2rem;}
.main-header h1{color:white;margin:0;font-size:1.8rem;}
.main-header p{color:#BDD7EE;margin:.3rem 0 0;font-size:.88rem;}
.kpi-box{background:white;border:1px solid #BDD7EE;border-radius:8px;padding:.7rem;text-align:center;}
.kpi-box .val{font-size:1.15rem;font-weight:bold;color:#1F3864;}
.kpi-box .lbl{font-size:.72rem;color:#888;margin-top:.2rem;}
.success-box{background:#E2EFDA;border:1px solid #70AD47;border-radius:8px;padding:.9rem 1.3rem;color:#375623;}
.warn-box{background:#FFF2CC;border:1px solid #FFD700;border-radius:8px;padding:.7rem 1.1rem;color:#7B5900;}
.error-box{background:#FCE4D6;border:1px solid #C55A11;border-radius:8px;padding:.9rem 1.3rem;color:#7B2C00;}
div[data-testid="stDownloadButton"] button{
  background:linear-gradient(135deg,#1F3864,#2E75B6);color:white;border:none;
  padding:.8rem 2.5rem;font-size:1rem;border-radius:8px;width:100%;}
</style>""", unsafe_allow_html=True)

st.markdown("""<div class="main-header">
  <h1>📊 FiscalXL v7</h1>
  <p>Convertisseur PDF → Excel · Pièces annexes IS (Modèle Comptable Normal, loi 9-88 Maroc)</p>
</div>""", unsafe_allow_html=True)

# ── Sidebar ───────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("### ⚙️ Configuration")
    show_preview = st.toggle("Aperçu des données", value=True)
    show_debug   = st.toggle("Mode débogage",      value=False)
    st.markdown("---")
    st.markdown("""**Formats supportés :**
- 🏛️ DGI — 7 pages
- 📄 AMMC / Standard — 5 pages
- 📋 Tous formats MCN loi 9-88

**Méthode d'extraction :**
- Tableaux avec bordures → `extract_tables()`
- Autres formats → Algorithme X/Y""")
    st.markdown("---")
    if TEMPLATE_PATH.exists(): st.success("✅ Template chargé")
    else: st.error("⚠️ EX_template.xlsx manquant")
    st.caption("FiscalXL v7 · MCN loi 9-88")

# ── Injection ─────────────────────────────────────────────────────────────────
def inject_data(extracted: dict, wb) -> int:
    """Injecte les données extraites dans l'Excel."""
    idx = build_excel_index(wb)
    count = 0
    for section, data in [
        ('actif',  extracted.get('actif_values',  {})),
        ('passif', extracted.get('passif_values', {})),
        ('cpc',    extracted.get('cpc_values',    {})),
    ]:
        for label, vals in data.items():
            n = ALIASES.get(label, label)
            ci = idx.get(n)
            if not ci: continue
            sheet, row = ci
            ws = wb[sheet]
            if section == 'actif':
                b, a, n1 = (vals+[None,None,None])[:3]
                if b  is not None: ws.cell(row,3).value = b;  count += 1
                if a  is not None: ws.cell(row,4).value = a;  count += 1
                if n1 is not None: ws.cell(row,6).value = n1
            elif section == 'passif':
                vn, vn1 = (vals+[None,None])[:2]
                if vn  is not None: ws.cell(row,3).value = vn;  count += 1
                if vn1 is not None: ws.cell(row,4).value = vn1; count += 1
            elif section == 'cpc':
                p, pr, t1 = (vals+[None,None,None])[:3]
                if p  is not None: ws.cell(row,3).value = p;  count += 1
                if pr is not None: ws.cell(row,4).value = pr; count += 1
                if t1 is not None: ws.cell(row,6).value = t1; count += 1
    return count

def update_headers(wb, info: dict):
    raison   = info.get("raison_sociale") or "—"
    id_fisc  = info.get("identifiant_fiscal") or ""
    exercice = info.get("exercice") or ""
    sub = f"{raison}  —  IF: {id_fisc}" if id_fisc else raison
    for sheet, title in [
        ("2 - Bilan Actif",  f"BILAN — ACTIF  |  {exercice}"),
        ("3 - Bilan Passif", f"BILAN — PASSIF  |  {exercice}"),
        ("4 - CPC",          f"COMPTE DE PRODUITS ET CHARGES  |  {exercice}"),
    ]:
        if sheet not in wb.sheetnames: continue
        ws = wb[sheet]
        if ws.cell(1,2).value: ws.cell(1,2).value = title
        if ws.cell(2,1).value is not None: ws.cell(2,1).value = sub

# ── Layout ────────────────────────────────────────────────────────────────────
col_up, col_info = st.columns([3,2])
with col_up:
    st.markdown("### 📂 Importer le PDF")
    uploaded = st.file_uploader("Glissez-déposez ou cliquez", type=["pdf"])

with col_info:
    st.markdown("### 🔄 Pipeline")
    for step, desc in [
        ("1 · Validation",   "Vérification structure PDF"),
        ("2 · Extraction",   "Tableaux → données structurées"),
        ("3 · Injection",    "Label → cellule Excel exacte"),
        ("4 · Formules",     "Totaux calculés automatiquement"),
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
      <p>Formats supportés : DGI (7 pages) · AMMC (5 pages) · Tous formats MCN</p>
    </div>""", unsafe_allow_html=True)
    st.stop()

# ── Processing ────────────────────────────────────────────────────────────────
st.markdown("---")
if not TEMPLATE_PATH.exists():
    st.markdown('<div class="error-box">❌ EX_template.xlsx manquant.</div>', unsafe_allow_html=True)
    st.stop()

with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
    tmp.write(uploaded.getbuffer())
    pdf_path = tmp.name
output_path = pdf_path.replace(".pdf", "_out.xlsx")
shutil.copy(str(TEMPLATE_PATH), output_path)

try:
    progress = st.progress(0)
    status   = st.empty()

    # Étape 1 — Validation
    status.info("🔍 Validation du PDF...")
    progress.progress(15)
    base = PDFParser(pdf_path)
    v    = validate_pdf_structure_v2(base)
    if not v["valid"]:
        st.markdown(f'<div class="error-box">⚠️ {v["message"]}</div>', unsafe_allow_html=True)
        st.stop()

    meta    = v["meta"]
    n_pages = meta.get("pages", 0)
    for col, (lbl, val) in zip(st.columns(4), [
        ("Raison Sociale",    (meta.get("raison_sociale") or "—")[:22]),
        ("Identifiant Fiscal", meta.get("identifiant_fiscal") or "—"),
        ("Fin exercice",       meta.get("exercice_fin") or "—"),
        ("Pages",              str(n_pages)),
    ]):
        col.markdown(f'<div class="kpi-box"><div class="val">{val}</div>'
                     f'<div class="lbl">{lbl}</div></div>', unsafe_allow_html=True)
    st.markdown("<br>", unsafe_allow_html=True)
    progress.progress(25)

    # Étape 2 — Extraction
    status.info("📄 Extraction des données...")
    progress.progress(45)
    with st.spinner("Lecture des tableaux..."):
        parser    = TableParser(pdf_path)
        extracted = parser.parse()
    progress.progress(65)

    if show_debug:
        with st.expander("🐛 Données brutes"):
            for section, data in extracted.items():
                if isinstance(data, dict):
                    st.write(f"**{section}** ({len(data)} postes)")
                    st.json({k: v for k,v in list(data.items())[:10]})

    # Étape 3 — Injection
    status.info("🔗 Injection dans le template...")
    progress.progress(75)
    wb    = openpyxl.load_workbook(output_path)
    count = inject_data(extracted, wb)
    update_headers(wb, extracted.get("info", {}))
    wb.save(output_path)
    progress.progress(95)

    # Compte formules
    wb_check  = openpyxl.load_workbook(output_path)
    n_formulas = sum(1 for ws in wb_check.worksheets
                     for row in ws.iter_rows()
                     for c in row if isinstance(c.value,str) and c.value.startswith("="))
    progress.progress(100)
    status.empty()

    # Résumé
    st.markdown(f"""<div class="success-box">
      ✅ <strong>Fichier Excel généré !</strong>
      &nbsp;·&nbsp; {count} valeurs injectées
      &nbsp;·&nbsp; {n_formulas} formules intactes
      &nbsp;·&nbsp; Détection automatique du format
    </div>""", unsafe_allow_html=True)
    st.markdown("<br>", unsafe_allow_html=True)

    # Aperçu
    if show_preview:
        with st.expander("👁️ Aperçu des valeurs extraites", expanded=False):
            tabs = st.tabs(["ℹ️ Infos","📋 Actif","📋 Passif","📈 CPC"])
            with tabs[0]:
                for k,v in extracted.get("info",{}).items():
                    if k != "pages" and v: st.markdown(f"**{k.replace('_',' ').title()}** : {v}")
            with tabs[1]:
                av = extracted.get("actif_values",{})
                if av:
                    df = pd.DataFrame([(k,
                        f"{v[0]:,.2f}" if v and v[0] is not None else "—",
                        f"{v[1]:,.2f}" if len(v)>1 and v[1] is not None else "—")
                        for k,v in av.items()], columns=["Poste","Brut","Amort."])
                    st.dataframe(df, use_container_width=True, height=300)
            with tabs[2]:
                pv = extracted.get("passif_values",{})
                if pv:
                    df = pd.DataFrame([(k,
                        f"{v[0]:,.2f}" if v and v[0] is not None else "—",
                        f"{v[1]:,.2f}" if len(v)>1 and v[1] is not None else "—")
                        for k,v in pv.items()], columns=["Poste","N","N-1"])
                    st.dataframe(df, use_container_width=True, height=300)
            with tabs[3]:
                cv = extracted.get("cpc_values",{})
                if cv:
                    df = pd.DataFrame([(k,
                        f"{v[0]:,.2f}" if v and v[0] is not None else "—",
                        f"{v[2]:,.2f}" if len(v)>2 and v[2] is not None else "—")
                        for k,v in cv.items()], columns=["Désignation","Propre N","Total N-1"])
                    st.dataframe(df, use_container_width=True, height=300)

    # Téléchargement
    st.markdown("### ⬇️ Télécharger")
    raison_slug = (extracted.get("info",{}).get("raison_sociale") or "fiscal").replace(" ","_")[:20]
    date_slug   = (extracted.get("info",{}).get("exercice_fin") or "").replace("/","-")
    fname = f"{raison_slug}_{date_slug}_fiscal.xlsx"
    with open(output_path,"rb") as f:
        st.download_button("📥 Télécharger le fichier Excel", data=f,
            file_name=fname,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

except Exception as e:
    logger.exception("Erreur pipeline")
    st.markdown(f'<div class="error-box">❌ <strong>Erreur :</strong> <code>{e}</code></div>',
                unsafe_allow_html=True)
    if show_debug:
        import traceback; st.code(traceback.format_exc())
finally:
    for f in [pdf_path, output_path]:
        try:
            if os.path.exists(f): os.unlink(f)
        except: pass
