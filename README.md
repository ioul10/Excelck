# 📊 FiscalXL — PDF Fiscal → Excel Pro

Convertisseur automatique de **pièces annexes à la déclaration IS** (Modèle Comptable Normal, loi 9-88 Maroc) vers un classeur Excel structuré, coloré et avec formules dynamiques.

---

## ✨ Fonctionnalités

| Étape | Description |
|-------|-------------|
| **1 - Extraction** | Lecture du PDF avec `pdfplumber` |
| **2 - Structuration** | Détection des tableaux : Bilan Actif, Passif, CPC |
| **3 - Relations** | Reconstruction des totaux et formules comptables |
| **4 - Excel Pro** | Classeur 5 feuilles, coloré, avec 100+ formules dynamiques |

### Feuilles générées
- `1 - Infos Générales` : Raison sociale, IF, exercice…
- `2 - Bilan Actif` : Brut / Amort. / **Net N = =Brut-Amort** / Net N-1
- `3 - Bilan Passif` : Capitaux propres, Dettes, Totaux avec formules
- `4 - CPC` : Total N = =Propre+Précédents, Résultats en formules différentielles
- `5 - Tableau de Bord` : KPIs avec **liens inter-feuilles** (`='4 - CPC'!D10`)

---

## 🚀 Installation

```bash
git clone https://github.com/votre-compte/fiscalxl.git
cd fiscalxl
pip install -r requirements.txt
streamlit run app.py
```

---

## 📁 Structure du projet

```
fiscalxl/
├── app.py                    ← Point d'entrée Streamlit
├── requirements.txt
├── README.md
│
├── core/
│   ├── extractor.py          ← Étape 1-2 : PDF → données brutes
│   ├── transformer.py        ← Étape 3 : alignement sur schéma canonique
│   ├── fiscal_schema.py      ← Schéma MCN officiel + définitions formules
│   └── excel_builder.py      ← Étape 4 : génération Excel Pro
│
└── utils/
    ├── validator.py          ← Validation structure PDF
    └── logger.py             ← Logging
```

---

## 🔧 Architecture détaillée

### `core/extractor.py` — PDFExtractor
Lit le PDF avec `pdfplumber`, sépare les pages par type :
- Détection heuristique via mots-clés
- `_split_label_nums()` : sépare le libellé des valeurs numériques (format marocain `1 234,56`)
- Retourne `dict[section → list[tuple]]`

### `core/transformer.py` — FiscalTransformer
Aligne les données extraites sur le schéma officiel :
1. Correspondance exacte (label normalisé)
2. Correspondance fuzzy par mots-clés significatifs (seuil 60%)
3. Fallback sur valeurs du schéma

### `core/fiscal_schema.py` — Schéma MCN
Contient :
- `BILAN_ACTIF_SCHEMA` / `BILAN_PASSIF_SCHEMA` / `CPC_SCHEMA` : structure officielle
- `ACTIF_FORMULAS` / `PASSIF_FORMULAS` / `CPC_FORMULAS` : `{total: [composants]}`
- `CPC_DIFFERENCES` : `{résultat: (positif, négatif)}` pour formules différentielles

### `core/excel_builder.py` — ExcelBuilder
Génère le classeur avec `openpyxl` :
- **Formules NET** : `=Brut-Amort` pour chaque ligne d'actif
- **Formules TOTAUX** : `=A1+A2+A3...` construites depuis `ACTIF_FORMULAS`
- **Formules CPC** : `Total N = =B+C`, Résultats = différences/additions
- **Liens inter-feuilles** : Tableau de Bord référence les 3 autres feuilles

---

## 📋 Format PDF accepté

Le PDF doit être une **pièce annexe à la déclaration IS, Modèle Normal** conforme à la loi 9-88 :

```
Page 1  : Informations générales (raison sociale, IF, exercice)
Page 2  : Bilan Actif (1/2) — Actif immobilisé
Page 3  : Bilan Actif (2/2) + Bilan Passif (début)
Page 4  : Bilan Passif (fin) + CPC (1/2)
Page 5  : CPC (2/2)
```

---

## ⚙️ Options Streamlit

| Option | Effet |
|--------|-------|
| Tableau de Bord | Ajoute la feuille de synthèse avec KPIs |
| Formules dynamiques | `=B-C` au lieu de valeurs figées |
| Mise en forme colorée | Palette bleue/dorée style financier |

---

## 📦 Dépendances

| Package | Usage |
|---------|-------|
| `streamlit` | Interface web |
| `pdfplumber` | Extraction texte/tableaux PDF |
| `openpyxl` | Génération Excel avec formules |
| `pandas` | Prévisualisation des données |
