"""
FiscalXL v5 — PDF Fiscal → Excel Pro
Mode DGI (7 pages) et AMMC (5 pages)
Headers Excel dynamiques depuis le PDF
Comparaison précise PDF vs Excel
"""

import streamlit as st
import tempfile, os
from pathlib import Path
import pandas as pd
import openpyxl

from core.pdf_parser import PDFParser
from core.injector import TemplateInjector
from utils.validator import validate_pdf_structure_v2
from utils.logger import get_logger

logger = get_logger(__name__)
TEMPLATE_PATH = Path(__file__).parent / "template_fiscal.xlsx"

# ─── Config ───────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="FiscalXL — PDF → Excel",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown("""
<style>
.main-header {
    background: linear-gradient(135deg, #1F3864 0%, #2E75B6 100%);
    padding: 1.6rem 2rem; border-radius: 12px; margin-bottom: 1.2rem;
}
.main-header h1 { color: white; margin: 0; font-size: 1.8rem; }
.main-header p  { color: #BDD7EE; margin: 0.3rem 0 0; font-size: 0.88rem; }
.kpi-box {
    background: white; border: 1px solid #BDD7EE;
    border-radius: 8px; padding: 0.7rem; text-align: center;
}
.kpi-box .val { font-size: 1.2rem; font-weight: bold; color: #1F3864; }
.kpi-box .lbl { font-size: 0.72rem; color: #888; margin-top: 0.2rem; }
.success-box { background:#E2EFDA; border:1px solid #70AD47; border-radius:8px; padding:0.9rem 1.3rem; color:#375623; }
.warn-box    { background:#FFF2CC; border:1px solid #FFD700; border-radius:8px; padding:0.7rem 1.1rem; color:#7B5900; }
.error-box   { background:#FCE4D6; border:1px solid #C55A11; border-radius:8px; padding:0.9rem 1.3rem; color:#7B2C00; }
.diff-ok     { color: #375623; background: #E2EFDA; padding: 2px 6px; border-radius: 4px; font-size: 0.82rem; }
.diff-warn   { color: #7B5900; background: #FFF2CC; padding: 2px 6px; border-radius: 4px; font-size: 0.82rem; }
.diff-error  { color: #7B2C00; background: #FCE4D6; padding: 2px 6px; border-radius: 4px; font-size: 0.82rem; }
.mode-badge-dgi  { background:#1F3864; color:white; padding:3px 10px; border-radius:20px; font-size:0.78rem; font-weight:bold; }
.mode-badge-ammc { background:#70AD47; color:white; padding:3px 10px; border-radius:20px; font-size:0.78rem; font-weight:bold; }
div[data-testid="stDownloadButton"] button {
    background: linear-gradient(135deg, #1F3864, #2E75B6);
    color:white; border:none; padding:0.8rem 2.5rem;
    font-size:1rem; border-radius:8px; width:100%;
}
</style>
""", unsafe_allow_html=True)

st.markdown("""
<div class="main-header">
    <h1>📊 FiscalXL v5</h1>
    <p>Convertisseur PDF → Excel · Pièces annexes IS (Modèle Comptable Normal, loi 9-88 Maroc)</p>
</div>
""", unsafe_allow_html=True)

# ─── Sidebar ──────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("### ⚙️ Configuration")

    mode = st.radio(
        "**Format PDF**",
        options=["🏛️ DGI — 7 pages", "📄 AMMC / Standard — 5 pages"],
        index=0,
        help=(
            "**DGI** : Étatde synthèse DGI (7 pages) — Page 1: Infos, "
            "Pages 2-3: Actif, Page 4: Passif, Pages 5-7: CPC\n\n"
            "**AMMC** : Liasse fiscale standard (5 pages) — Page 1: Infos, "
            "Page 2: Actif, Page 3: Passif, Pages 4-5: CPC"
        ),
    )
    is_dgi = "DGI" in mode

    st.markdown("---")
    show_preview    = st.toggle("Aperçu des données extraites", value=True)
    show_comparison = st.toggle("Comparaison PDF ↔ Excel",      value=True)
    show_debug      = st.toggle("Mode débogage",                value=False)

    st.markdown("---")
    if is_dgi:
        st.markdown('<span class="mode-badge-dgi">MODE DGI — 7 PAGES</span>', unsafe_allow_html=True)
        st.markdown("""
        - **Page 1** — Identification
        - **Pages 2-3** — Bilan Actif
        - **Page 4** — Bilan Passif
        - **Pages 5-7** — CPC
        """)
    else:
        st.markdown('<span class="mode-badge-ammc">MODE AMMC — 5 PAGES</span>', unsafe_allow_html=True)
        st.markdown("""
        - **Page 1** — Identification
        - **Page 2** — Bilan Actif
        - **Page 3** — Bilan Passif
        - **Pages 4-5** — CPC
        """)

    if not TEMPLATE_PATH.exists():
        st.error("⚠️ Template introuvable")
    else:
        st.success("✅ Template chargé")
    st.caption("FiscalXL v5 · MCN loi 9-88")

# ─── Upload ───────────────────────────────────────────────────────────────────
col_up, col_info = st.columns([3, 2])
with col_up:
    st.markdown("### 📂 Importer le PDF")
    uploaded = st.file_uploader("Glissez-déposez ou cliquez", type=["pdf"])

with col_info:
    st.markdown("### 🔄 Pipeline")
    for step, desc in [
        ("1 · Lecture PDF",         "pdfplumber extrait mots et positions X/Y"),
        ("2 · Parsing valeurs",     "Nombres alignés sur labels par colonnes"),
        ("3 · Headers dynamiques",  "Raison sociale, exercice depuis le PDF"),
        ("4 · Injection template",  "Valeurs → cellules exactes du modèle"),
        ("5 · Comparaison",         "Vérification PDF ↔ Excel poste par poste"),
    ]:
        st.markdown(
            f'<div style="background:#f8f9fa;border-left:4px solid #2E75B6;'
            f'padding:0.5rem 0.8rem;border-radius:0 6px 6px 0;margin:0.3rem 0;">'
            f'<strong style="color:#1F3864">{step}</strong><br>'
            f'<span style="color:#555;font-size:0.83rem">{desc}</span></div>',
            unsafe_allow_html=True
        )

# ─── Helpers ──────────────────────────────────────────────────────────────────

def update_excel_headers(wb, info: dict):
    """Met à jour les headers de chaque feuille avec les infos du PDF."""
    raison   = info.get("raison_sociale") or "—"
    id_fisc  = info.get("identifiant_fiscal") or ""
    exercice = info.get("exercice") or ""
    date_fin = info.get("exercice_fin") or ""

    sub = f"{raison}  —  IF: {id_fisc}" if id_fisc else raison

    titles = {
        "2 - Bilan Actif":   f"BILAN — ACTIF  |  {exercice}",
        "3 - Bilan Passif":  f"BILAN — PASSIF  |  {exercice}",
        "4 - CPC":           f"COMPTE DE PRODUITS ET CHARGES  |  {exercice}",
        "5 - Tableau de Bord": f"TABLEAU DE BORD — SYNTHÈSE FINANCIÈRE",
    }
    for sheet_name, title in titles.items():
        if sheet_name not in wb.sheetnames:
            continue
        ws = wb[sheet_name]
        # Ligne 1 = titre, Ligne 2 = sous-titre (raison sociale)
        ws["A1"] = title
        ws["A2"] = sub if sheet_name != "5 - Tableau de Bord" else raison

    # Feuille Infos Générales
    if "1 - Infos Générales" in wb.sheetnames:
        ws = wb["1 - Infos Générales"]
        # Mettre à jour les valeurs dans la colonne B
        info_map = {
            "Raison sociale":        raison,
            "Identifiant fiscal":    id_fisc,
            "Taxe professionnelle":  info.get("taxe_professionnelle") or "—",
            "Adresse":               info.get("adresse") or "—",
            "Exercice":              exercice,
            "Date de déclaration":   info.get("date_declaration") or "—",
            "Nombre de pages":       str(info.get("pages") or "—"),
        }
        for row in ws.iter_rows():
            if row[0].value in info_map:
                row[1].value = info_map[row[0].value]


def build_comparison(extracted: dict, wb) -> list:
    """
    Compare les valeurs extraites du PDF avec ce qui est dans l'Excel.
    Retourne une liste de dicts avec les écarts.
    """
    rows = []

    # Mapping postes clés → (feuille, cellule_N, cellule_N1, label_affichage)
    key_checks = [
        # Bilan Actif
        ("actif",  "Immobilisations en non-valeurs",     "2 - Bilan Actif",  "B5",  "E5",  "Immobilisations en non-valeurs [A]"),
        ("actif",  "Immobilisations incorporelles",       "2 - Bilan Actif",  "B9",  "E9",  "Immobilisations incorporelles [B]"),
        ("actif",  "Immobilisations corporelles",         "2 - Bilan Actif",  "B14", "E14", "Immobilisations corporelles [C]"),
        ("actif",  "Terrains",                            "2 - Bilan Actif",  "B15", "E15", "Terrains"),
        ("actif",  "Constructions",                       "2 - Bilan Actif",  "B16", "E16", "Constructions"),
        ("actif",  "Matières et fournitures consommables","2 - Bilan Actif",  "B34", "E34", "Matières et fournitures consommables"),
        ("actif",  "Clients et comptes rattachés",        "2 - Bilan Actif",  "B40", "E40", "Clients et comptes rattachés"),
        ("actif",  "Banques, T.G et C.C.P",               "2 - Bilan Actif",  "B51", "E51", "Banques, T.G et C.C.P"),
        # Bilan Passif
        ("passif", "Capital social ou personnel",         "3 - Bilan Passif", "B6",  "C6",  "Capital social"),
        ("passif", "Résultat net de l'exercice",          "3 - Bilan Passif", "B15", "C15", "Résultat net exercice"),
        ("passif", "Subvention d'investissement",         "3 - Bilan Passif", "B18", "C18", "Subventions d'investissement"),
        ("passif", "Fournisseurs et comptes rattachés",   "3 - Bilan Passif", "B30", "C30", "Fournisseurs et ctes rattachés"),
        ("passif", "Autres dettes de financement",        "3 - Bilan Passif", "B22", "C22", "Autres dettes de financement"),
        # CPC
        ("cpc",    "Ventes de biens et services produits","4 - CPC",          "B6",  "E6",  "Ventes de biens et services"),
        ("cpc",    "Achats consommés",                    "4 - CPC",          "B11", "E11", "Achats consommés matières"),
        ("cpc",    "Charges de personnel",                "4 - CPC",          "B19", "E19", "Charges de personnel"),
        ("cpc",    "Dotations d'exploitation",            "4 - CPC",          "B20", "E20", "Dotations d'exploitation"),
        ("cpc",    "IMPOTS SUR LES BENEFICES",            "4 - CPC",          "B53", "E53", "Impôts sur les résultats"),
    ]

    actif  = extracted.get("actif_values",  {})
    passif = extracted.get("passif_values", {})
    cpc    = extracted.get("cpc_values",    {})
    sources = {"actif": actif, "passif": passif, "cpc": cpc}

    for section, search_key, sheet, cell_n, cell_n1, display in key_checks:
        src = sources[section]

        # Trouver la valeur PDF (recherche approximative)
        pdf_vals = None
        for k, v in src.items():
            if search_key.lower() in k.lower() or k.lower() in search_key.lower():
                pdf_vals = v
                break

        pdf_n  = pdf_vals[0] if pdf_vals and len(pdf_vals) > 0 else None
        pdf_n1 = pdf_vals[1] if pdf_vals and len(pdf_vals) > 1 else None
        if section == "actif" and pdf_vals and len(pdf_vals) >= 3:
            pdf_n1 = pdf_vals[2]  # net N-1

        # Lire la valeur Excel
        xl_n  = None
        xl_n1 = None
        if sheet in wb.sheetnames:
            ws = wb[sheet]
            try:
                v = ws[cell_n].value
                xl_n = float(v) if isinstance(v, (int, float)) else None
            except Exception:
                pass
            try:
                v = ws[cell_n1].value
                # Ne pas lire si c'est une formule
                if isinstance(v, (int, float)):
                    xl_n1 = float(v)
            except Exception:
                pass

        # Comparer
        def fmt(v):
            if v is None: return "—"
            return f"{v:,.2f}"

        def status(a, b):
            if a is None and b is None: return "vide"
            if a is None or b is None:  return "manquant"
            if abs(a - b) < 1:          return "ok"
            if abs(a - b) / max(abs(a), 1) < 0.01: return "ok"
            return "erreur"

        st_n  = status(pdf_n,  xl_n)
        st_n1 = status(pdf_n1, xl_n1)

        rows.append({
            "Section":   section.upper(),
            "Poste":     display,
            "Cellule N": cell_n,
            "PDF — N":   fmt(pdf_n),
            "Excel — N": fmt(xl_n),
            "Statut N":  st_n,
            "Cellule N-1": cell_n1,
            "PDF — N-1": fmt(pdf_n1),
            "Excel — N-1": fmt(xl_n1),
            "Statut N-1": st_n1,
        })

    return rows


def render_comparison(rows: list):
    """Affiche le tableau de comparaison avec couleurs."""
    if not rows:
        return

    errors   = [r for r in rows if r["Statut N"] == "erreur" or r["Statut N-1"] == "erreur"]
    warnings = [r for r in rows if r["Statut N"] == "manquant" or r["Statut N-1"] == "manquant"]
    ok_count = len([r for r in rows if r["Statut N"] == "ok" and r["Statut N-1"] in ("ok", "vide")])

    col1, col2, col3 = st.columns(3)
    col1.metric("✅ Postes corrects",   ok_count)
    col2.metric("⚠️ Valeurs manquantes", len(warnings))
    col3.metric("❌ Erreurs de valeur",  len(errors))

    if errors:
        st.markdown("#### ❌ Erreurs de valeur — à corriger manuellement")
        df_err = pd.DataFrame(errors)[["Section","Poste","Cellule N","PDF — N","Excel — N","Cellule N-1","PDF — N-1","Excel — N-1"]]
        st.dataframe(df_err, use_container_width=True, height=min(200, 50 + len(errors)*35))

    if warnings:
        st.markdown("#### ⚠️ Valeurs non injectées")
        df_warn = pd.DataFrame(warnings)[["Section","Poste","Cellule N","PDF — N","Excel — N"]]
        st.dataframe(df_warn, use_container_width=True, height=min(200, 50 + len(warnings)*35))

    if not errors and not warnings:
        st.success("✅ Toutes les valeurs vérifiées sont correctes !")

    # Tableau complet
    with st.expander("📋 Tableau complet de comparaison"):
        rows_display = []
        for r in rows:
            def badge(s):
                if s == "ok":       return "✅"
                if s == "manquant": return "⚠️"
                if s == "erreur":   return "❌"
                return "—"
            rows_display.append({
                "Section":  r["Section"],
                "Poste":    r["Poste"],
                "Cel.N":    r["Cellule N"],
                "PDF N":    r["PDF — N"],
                "Excel N":  r["Excel — N"],
                "":         badge(r["Statut N"]),
                "Cel.N-1":  r["Cellule N-1"],
                "PDF N-1":  r["PDF — N-1"],
                "Excel N-1":r["Excel — N-1"],
                " ":        badge(r["Statut N-1"]),
            })
        st.dataframe(pd.DataFrame(rows_display), use_container_width=True, height=500)


# ─── Main ─────────────────────────────────────────────────────────────────────

if uploaded:
    st.markdown("---")

    if not TEMPLATE_PATH.exists():
        st.markdown('<div class="error-box">❌ <strong>Template manquant.</strong></div>', unsafe_allow_html=True)
        st.stop()

    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
        tmp.write(uploaded.getbuffer())
        pdf_path = tmp.name
    output_path = pdf_path.replace(".pdf", "_out.xlsx")

    try:
        progress = st.progress(0)
        status   = st.empty()

        # ── Étape 1 : Validation ──────────────────────────────────────────────
        status.info("🔍 Étape 1/5 — Validation du PDF...")
        progress.progress(10)
        parser = PDFParser(pdf_path)
        validation = validate_pdf_structure_v2(parser)
        if not validation["valid"]:
            st.markdown(f'<div class="error-box">⚠️ {validation["message"]}</div>', unsafe_allow_html=True)
            st.stop()

        meta = validation["meta"]
        n_pages = meta.get("pages", 0)

        # Avertissement si le mode ne correspond pas
        if is_dgi and n_pages != 7:
            st.warning(f"⚠️ Mode DGI sélectionné mais le PDF a {n_pages} pages (attendu : 7). Résultats possiblement incomplets.")
        elif not is_dgi and n_pages not in (5, 6):
            st.warning(f"⚠️ Mode AMMC sélectionné mais le PDF a {n_pages} pages (attendu : 5). Résultats possiblement incomplets.")

        # KPIs
        badge_html = f'<span class="mode-badge-dgi">DGI</span>' if is_dgi else f'<span class="mode-badge-ammc">AMMC</span>'
        c1, c2, c3, c4, c5 = st.columns(5)
        for col, (lbl, val) in zip([c1,c2,c3,c4,c5], [
            ("Raison Sociale",    (meta.get("raison_sociale") or "—")[:20]),
            ("Identifiant Fiscal", meta.get("identifiant_fiscal") or "—"),
            ("Fin exercice",      meta.get("exercice_fin") or "—"),
            ("Pages",             str(n_pages)),
            ("Mode",              "DGI" if is_dgi else "AMMC"),
        ]):
            col.markdown(f'<div class="kpi-box"><div class="val">{val}</div><div class="lbl">{lbl}</div></div>',
                         unsafe_allow_html=True)
        st.markdown("<br>", unsafe_allow_html=True)
        progress.progress(20)

        # ── Étape 2 : Extraction ──────────────────────────────────────────────
        status.info("📄 Étape 2/5 — Extraction des valeurs PDF...")
        progress.progress(35)
        with st.spinner("Parsing en cours..."):
            extracted = parser.parse()
        progress.progress(55)

        if show_debug:
            with st.expander("🐛 Données brutes extraites"):
                st.json({k: {str(l): v for l,v in d.items()} if isinstance(d, dict) else d
                         for k,d in extracted.items()})

        # ── Étape 3 : Injection ───────────────────────────────────────────────
        status.info("🔗 Étape 3/5 — Injection dans le template...")
        progress.progress(65)
        with st.spinner("Injection..."):
            injector = TemplateInjector(str(TEMPLATE_PATH))
            stats = injector.inject(extracted, output_path)
        progress.progress(78)

        # ── Étape 4 : Headers dynamiques ─────────────────────────────────────
        status.info("🏷️ Étape 4/5 — Mise à jour des headers...")
        wb = openpyxl.load_workbook(output_path)
        update_excel_headers(wb, extracted["info"])
        wb.save(output_path)
        progress.progress(88)

        # ── Étape 5 : Vérification ────────────────────────────────────────────
        status.info("✅ Étape 5/5 — Vérification...")
        wb_check = openpyxl.load_workbook(output_path)
        n_formulas = sum(
            1 for ws in wb_check.worksheets
            for row in ws.iter_rows()
            for c in row
            if isinstance(c.value, str) and c.value.startswith("=")
        )
        progress.progress(100)
        status.empty()

        # ── Résumé ────────────────────────────────────────────────────────────
        st.markdown(f"""
        <div class="success-box">
            ✅ <strong>Fichier Excel généré !</strong>
            &nbsp;·&nbsp; {len(wb_check.sheetnames)} feuilles
            &nbsp;·&nbsp; {n_formulas} formules intactes
            &nbsp;·&nbsp; {stats['injected']} valeurs injectées
            &nbsp;·&nbsp; {badge_html}
        </div>""", unsafe_allow_html=True)

        # Alertes postes non mappés — enrichies
        if stats.get("not_found"):
            nf = stats["not_found"]
            # Filtrer les totaux et sections normaux
            real_missing = [x for x in nf if len(x) > 4 and not any(
                t in x.upper() for t in ["TOTAL", "CAPITAUX PROPRES", "RESULTAT D'EX", "CHARGES D'EX"]
            )]
            if real_missing:
                st.markdown(
                    f'<div class="warn-box">⚠️ <strong>{len(real_missing)} postes non mappés</strong>'
                    f' — valeurs extraites du PDF mais cellule Excel non trouvée :<br>'
                    f'<code>{"  ·  ".join(real_missing[:8])}</code></div>',
                    unsafe_allow_html=True
                )

        st.markdown("<br>", unsafe_allow_html=True)

        # ── Aperçu données ────────────────────────────────────────────────────
        if show_preview:
            with st.expander("👁️ Aperçu des valeurs extraites", expanded=False):
                tabs = st.tabs(["ℹ️ Infos", "📋 Bilan Actif", "📋 Bilan Passif", "📈 CPC"])

                with tabs[0]:
                    info = extracted.get("info", {})
                    for k, v in info.items():
                        if k != "pages" and v:
                            st.markdown(f"**{k.replace('_',' ').title()}** : {v}")

                with tabs[1]:
                    av = extracted.get("actif_values", {})
                    if av:
                        df = pd.DataFrame(
                            [(k,
                              f"{v[0]:,.2f}" if v and v[0] is not None else "—",
                              f"{v[1]:,.2f}" if len(v)>1 and v[1] is not None else "—",
                              f"{v[2]:,.2f}" if len(v)>2 and v[2] is not None else "—")
                             for k, v in av.items()],
                            columns=["Poste", "Brut", "Amort.", "Net N-1"])
                        st.dataframe(df, use_container_width=True, height=320)

                with tabs[2]:
                    pv = extracted.get("passif_values", {})
                    if pv:
                        df = pd.DataFrame(
                            [(k,
                              f"{v[0]:,.2f}" if v and v[0] is not None else "—",
                              f"{v[1]:,.2f}" if len(v)>1 and v[1] is not None else "—")
                             for k, v in pv.items()],
                            columns=["Poste", "Exercice N", "Exercice N-1"])
                        st.dataframe(df, use_container_width=True, height=320)

                with tabs[3]:
                    cv = extracted.get("cpc_values", {})
                    if cv:
                        df = pd.DataFrame(
                            [(k,
                              f"{v[0]:,.2f}" if v and v[0] is not None else "—",
                              f"{v[1]:,.2f}" if len(v)>1 and v[1] is not None else "—",
                              f"{v[2]:,.2f}" if len(v)>2 and v[2] is not None else "—")
                             for k, v in cv.items()],
                            columns=["Désignation", "Propre N", "Exerc. Préc.", "Total N-1"])
                        st.dataframe(df, use_container_width=True, height=320)

        # ── Comparaison PDF ↔ Excel ───────────────────────────────────────────
        if show_comparison:
            st.markdown("### 🔍 Comparaison PDF ↔ Excel")
            comp_rows = build_comparison(extracted, wb_check)
            render_comparison(comp_rows)

        # ── Téléchargement ────────────────────────────────────────────────────
        st.markdown("### ⬇️ Télécharger")
        raison_slug = (extracted["info"].get("raison_sociale") or "fiscal").replace(" ", "_")[:20]
        date_slug   = (extracted["info"].get("exercice_fin") or "").replace("/", "-")
        fname = f"{raison_slug}_{date_slug}_fiscal.xlsx"

        with open(output_path, "rb") as f:
            st.download_button(
                "📥 Télécharger le fichier Excel",
                data=f, file_name=fname,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

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
            except Exception:
                pass

else:
    st.markdown("""
    <div style="text-align:center;padding:3rem;color:#888;
        border:2px dashed #BDD7EE;border-radius:12px;background:#f8fafd;margin-top:1rem;">
        <div style="font-size:3rem;">📄</div>
        <h3 style="color:#2E75B6;">Importez un PDF pour commencer</h3>
        <p>Pièces annexes à la déclaration IS — Modèle Comptable Normal (loi 9-88)</p>
        <p style="font-size:0.85rem;margin-top:0.5rem;">
            Formats supportés : <strong>DGI 7 pages</strong> · <strong>AMMC/Standard 5 pages</strong>
        </p>
    </div>""", unsafe_allow_html=True)
