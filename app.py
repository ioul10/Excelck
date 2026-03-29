"""FiscalXL v6 — DirectInjector + EX_template.xlsx"""

import streamlit as st
import tempfile, os, shutil
from pathlib import Path
import pandas as pd
import openpyxl

from core.direct_injector import DirectInjector, build_excel_index, soft_normalize
from utils.validator import validate_pdf_structure_v2
from core.pdf_parser import PDFParser
from utils.logger import get_logger

logger = get_logger(__name__)
TEMPLATE_PATH = Path(__file__).parent / "EX_template.xlsx"

st.set_page_config(page_title="FiscalXL v6", page_icon="📊", layout="wide",
                   initial_sidebar_state="expanded")

st.markdown("""
<style>
.main-header{background:linear-gradient(135deg,#1F3864,#2E75B6);padding:1.5rem 2rem;
  border-radius:12px;margin-bottom:1.2rem;}
.main-header h1{color:white;margin:0;font-size:1.8rem;}
.main-header p{color:#BDD7EE;margin:.3rem 0 0;font-size:.88rem;}
.kpi-box{background:white;border:1px solid #BDD7EE;border-radius:8px;
  padding:.7rem;text-align:center;}
.kpi-box .val{font-size:1.15rem;font-weight:bold;color:#1F3864;}
.kpi-box .lbl{font-size:.72rem;color:#888;margin-top:.2rem;}
.success-box{background:#E2EFDA;border:1px solid #70AD47;border-radius:8px;
  padding:.9rem 1.3rem;color:#375623;}
.warn-box{background:#FFF2CC;border:1px solid #FFD700;border-radius:8px;
  padding:.7rem 1.1rem;color:#7B5900;}
.error-box{background:#FCE4D6;border:1px solid #C55A11;border-radius:8px;
  padding:.9rem 1.3rem;color:#7B2C00;}
div[data-testid="stDownloadButton"] button{
  background:linear-gradient(135deg,#1F3864,#2E75B6);color:white;
  border:none;padding:.8rem 2.5rem;font-size:1rem;border-radius:8px;width:100%;}
</style>""", unsafe_allow_html=True)

st.markdown("""
<div class="main-header">
  <h1>📊 FiscalXL v6</h1>
  <p>Convertisseur PDF → Excel · Pièces annexes IS (Modèle Comptable Normal, loi 9-88 Maroc)</p>
</div>""", unsafe_allow_html=True)

# ── Sidebar ───────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("### ⚙️ Configuration")
    mode = st.radio(
        "**Format PDF**",
        options=["🏛️ DGI — 7 pages", "📄 AMMC / Standard — 5 pages"],
        index=0,
    )
    is_dgi = "DGI" in mode

    st.markdown("---")
    show_preview    = st.toggle("Aperçu des données", value=True)
    show_comparison = st.toggle("Comparaison PDF ↔ Excel", value=True)
    show_debug      = st.toggle("Mode débogage", value=False)
    st.markdown("---")

    if is_dgi:
        st.markdown('<span style="background:#1F3864;color:white;padding:3px 10px;border-radius:20px;font-size:.78rem;font-weight:bold;">MODE DGI — 7 PAGES</span>', unsafe_allow_html=True)
        st.markdown("- **Page 1** — Identification\n- **Pages 2-3** — Bilan Actif\n- **Page 4** — Bilan Passif\n- **Pages 5-7** — CPC")
    else:
        st.markdown('<span style="background:#70AD47;color:white;padding:3px 10px;border-radius:20px;font-size:.78rem;font-weight:bold;">MODE AMMC — 5 PAGES</span>', unsafe_allow_html=True)
        st.markdown("- **Page 1** — Identification\n- **Page 2** — Bilan Actif\n- **Page 3** — Bilan Passif\n- **Pages 4-5** — CPC")

    if TEMPLATE_PATH.exists():
        st.success("✅ Template EX chargé")
    else:
        st.error("⚠️ EX_template.xlsx manquant")
    st.caption("FiscalXL v6 · MCN loi 9-88")

# ── Helpers ───────────────────────────────────────────────────────────────────

def update_headers(wb, info: dict):
    raison   = info.get("raison_sociale") or "—"
    id_fisc  = info.get("identifiant_fiscal") or ""
    exercice = info.get("exercice") or ""
    sub = f"{raison}  —  IF: {id_fisc}" if id_fisc else raison

    for sheet, title in [
        ("2 - Bilan Actif",   f"BILAN — ACTIF  |  {exercice}"),
        ("3 - Bilan Passif",  f"BILAN — PASSIF  |  {exercice}"),
        ("4 - CPC",           f"COMPTE DE PRODUITS ET CHARGES  |  {exercice}"),
        ("5 - Tableau de Bord","TABLEAU DE BORD — SYNTHÈSE FINANCIÈRE"),
    ]:
        if sheet not in wb.sheetnames: continue
        ws = wb[sheet]
        if ws.cell(1, 2).value is not None:
            ws.cell(1, 2).value = title
        if ws.cell(2, 1).value is not None or True:
            pass  # Headers dynamiques gérés par les formules =Infos!B4


def build_comparison(wb, pdf_path: str) -> list:
    """Compare les valeurs PDF avec l'Excel généré poste par poste."""
    from core.direct_injector import extract_page_rows, detect_columns, _parse_num_str
    import pdfplumber

    checks = []
    pdf = pdfplumber.open(pdf_path)
    n   = len(pdf.pages)

    # Extraire quelques valeurs clés du PDF pour comparer
    page_idx = 3 if n == 7 else 2  # page passif
    rows = extract_page_rows(pdf.pages[page_idx])
    rows_cols = detect_columns(rows, 'passif')
    pdf.close()

    # Lire l'Excel généré
    wa = wb['2 - Bilan Actif']
    wp = wb['3 - Bilan Passif']
    wc = wb['4 - CPC']

    key_checks = [
        ("Passif", "Capital appelé",           wp, 7,  3),
        ("Passif", "Résultat net",              wp, 15, 3),
        ("Passif", "Subventions investissement",wp, 18, 3),
        ("Passif", "Fournisseurs",              wp, 31, 3),
        ("Passif", "Personnel passif",          wp, 33, 3),
        ("Actif",  "Terrains brut",             wa, 15, 3),
        ("Actif",  "Constructions brut",        wa, 16, 3),
        ("Actif",  "Installations brut",        wa, 17, 3),
        ("Actif",  "Clients brut",              wa, 39, 3),
        ("CPC",    "Ventes biens/services",     wc, 7,  3),
        ("CPC",    "Charges personnel",         wc, 20, 3),
        ("CPC",    "Dotations exploitation",    wc, 22, 3),
        ("CPC",    "Impôts résultats",          wc, 54, 3),
    ]

    for section, label, ws_ref, row, col in key_checks:
        val = ws_ref.cell(row, col).value
        if isinstance(val, str) and val.startswith('='): val = None
        checks.append({
            "Section": section,
            "Poste":   label,
            "Cellule": f"{ws_ref.title[:10]}!{ws_ref.cell(row,col).coordinate}",
            "Valeur":  f"{val:,.2f}" if isinstance(val,(int,float)) else ("—" if val is None else str(val)),
            "Statut":  "✅" if val is not None and val != 0 else "⚠️",
        })
    return checks


# ── Layout ────────────────────────────────────────────────────────────────────
col_up, col_pipe = st.columns([3, 2])
with col_up:
    st.markdown("### 📂 Importer le PDF")
    uploaded = st.file_uploader("Glissez-déposez ou cliquez", type=["pdf"])

with col_pipe:
    st.markdown("### 🔄 Pipeline")
    for step, desc in [
        ("1 · Validation",      "Structure et format du PDF"),
        ("2 · Extraction",      "Labels + valeurs par position X/Y"),
        ("3 · Correspondance",  "Label PDF → cellule Excel exacte"),
        ("4 · Injection",       "Remplir uniquement les cellules vides"),
        ("5 · Vérification",    "Formules recalculées vs PDF"),
    ]:
        color = "#1F3864" if is_dgi else "#70AD47"
        st.markdown(
            f'<div style="background:#f8f9fa;border-left:4px solid {color};'
            f'padding:.5rem .8rem;border-radius:0 6px 6px 0;margin:.3rem 0;">'
            f'<strong style="color:#1F3864">{step}</strong><br>'
            f'<span style="color:#555;font-size:.83rem">{desc}</span></div>',
            unsafe_allow_html=True)

if not uploaded:
    st.markdown("""
    <div style="text-align:center;padding:3rem;color:#888;border:2px dashed #BDD7EE;
      border-radius:12px;background:#f8fafd;margin-top:1rem;">
      <div style="font-size:3rem;">📄</div>
      <h3 style="color:#2E75B6;">Importez un PDF pour commencer</h3>
      <p>DGI (7 pages) ou AMMC (5 pages) — sélectionnez le format dans la sidebar</p>
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

# Copier le template dans un fichier temporaire de sortie
output_path = pdf_path.replace(".pdf", "_out.xlsx")
shutil.copy(str(TEMPLATE_PATH), output_path)

try:
    progress = st.progress(0)
    status   = st.empty()

    # ── Étape 1 : Validation ──────────────────────────────────────────────────
    status.info("🔍 Étape 1/5 — Validation du PDF...")
    progress.progress(10)

    base_parser = PDFParser(pdf_path)
    validation  = validate_pdf_structure_v2(base_parser, mode="dgi" if is_dgi else "ammc")
    if not validation["valid"]:
        st.markdown(f'<div class="error-box">⚠️ {validation["message"]}</div>', unsafe_allow_html=True)
        st.stop()

    meta    = validation["meta"]
    n_pages = meta.get("pages", 0)

    if is_dgi and n_pages != 7:
        st.warning(f"⚠️ Mode DGI sélectionné mais le PDF a {n_pages} pages (attendu : 7).")
    elif not is_dgi and n_pages not in (5, 6):
        st.warning(f"⚠️ Mode AMMC sélectionné mais le PDF a {n_pages} pages (attendu : 5).")

    for col, (lbl, val) in zip(
        st.columns(4),
        [("Raison Sociale",    (meta.get("raison_sociale") or "—")[:22]),
         ("Identifiant Fiscal", meta.get("identifiant_fiscal") or "—"),
         ("Fin exercice",       meta.get("exercice_fin") or "—"),
         ("Pages",              str(n_pages))]
    ):
        col.markdown(f'<div class="kpi-box"><div class="val">{val}</div>'
                     f'<div class="lbl">{lbl}</div></div>', unsafe_allow_html=True)
    st.markdown("<br>", unsafe_allow_html=True)
    progress.progress(20)

    # ── Étape 2-3-4 : Extraction + Correspondance + Injection ─────────────────
    status.info("📄 Étapes 2-4/5 — Extraction et injection...")
    progress.progress(40)

    with st.spinner("Traitement en cours..."):
        wb  = openpyxl.load_workbook(output_path)
        inj = DirectInjector(wb)
        stats = inj.inject_pdf(pdf_path)
        wb.save(output_path)

    progress.progress(75)

    # ── Étape 5 : Headers + Vérification ──────────────────────────────────────
    status.info("🏷️ Étape 5/5 — Headers et vérification...")
    wb_final = openpyxl.load_workbook(output_path)
    update_headers(wb_final, meta)
    wb_final.save(output_path)

    # Compter les formules intactes
    wb_check = openpyxl.load_workbook(output_path)
    n_formulas = sum(
        1 for ws in wb_check.worksheets
        for row in ws.iter_rows()
        for c in row if isinstance(c.value, str) and c.value.startswith("=")
    )
    progress.progress(100)
    status.empty()

    # ── Résumé ────────────────────────────────────────────────────────────────
    st.markdown(f"""
    <div class="success-box">
      ✅ <strong>Fichier Excel généré !</strong>
      &nbsp;·&nbsp; {stats['injected']} valeurs injectées
      &nbsp;·&nbsp; {n_formulas} formules intactes
      &nbsp;·&nbsp; {len(stats['skipped'])} postes non trouvés
    </div>""", unsafe_allow_html=True)

    # Postes non trouvés (filtrer les totaux normaux)
    real_skipped = [x for x in stats['skipped']
                    if len(x) > 3
                    and not any(t in x.upper() for t in
                        ['TOTAL','RÉSULTAT','CHIFFRE','PRODUITS D\'EX',
                         'CHARGES D\'EX','EXPLOITATION','FINANCIER','COURANT'])]
    if real_skipped:
        st.markdown(
            f'<div class="warn-box">⚠️ <strong>{len(real_skipped)} postes non mappés</strong>'
            f' (label PDF non trouvé dans le template) :<br>'
            f'<code>{"  ·  ".join(real_skipped[:8])}</code></div>',
            unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    # ── Aperçu ────────────────────────────────────────────────────────────────
    if show_preview:
        with st.expander("👁️ Valeurs injectées", expanded=False):
            st.markdown("**Cellules remplies depuis le PDF :**")
            df = pd.DataFrame(inj._injected, columns=["Label PDF", "Feuille", "Ligne"])
            st.dataframe(df, use_container_width=True, height=300)

    # ── Comparaison ───────────────────────────────────────────────────────────
    if show_comparison:
        st.markdown("### 🔍 Vérification des postes clés")
        try:
            comp = build_comparison(wb_check, pdf_path)
            df_comp = pd.DataFrame(comp)
            
            errors = [r for r in comp if r["Statut"] == "⚠️"]
            ok     = [r for r in comp if r["Statut"] == "✅"]
            
            c1, c2 = st.columns(2)
            c1.metric("✅ Postes injectés", len(ok))
            c2.metric("⚠️ Postes vides",    len(errors))
            
            if errors:
                st.markdown("**Postes vides — à vérifier :**")
                st.dataframe(pd.DataFrame(errors)[["Section","Poste","Cellule"]], 
                           use_container_width=True, height=200)
            else:
                st.success("✅ Tous les postes clés sont remplis !")
            
            with st.expander("📋 Tous les postes vérifiés"):
                st.dataframe(df_comp, use_container_width=True, height=450)
        except Exception as e:
            st.warning(f"Comparaison non disponible : {e}")

    # ── Debug ─────────────────────────────────────────────────────────────────
    if show_debug:
        with st.expander("🐛 Labels non mappés (complets)"):
            for s in stats['skipped']:
                st.text(s)

    # ── Téléchargement ────────────────────────────────────────────────────────
    st.markdown("### ⬇️ Télécharger")
    raison_slug = (meta.get("raison_sociale") or "fiscal").replace(" ","_")[:20]
    date_slug   = (meta.get("exercice_fin") or "").replace("/","-")
    fname = f"{raison_slug}_{date_slug}_fiscal.xlsx"

    with open(output_path, "rb") as f:
        st.download_button(
            "📥 Télécharger le fichier Excel",
            data=f, file_name=fname,
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
