# 📊 FiscalXL — Convertisseur universel PDF fiscal → Excel

Convertit automatiquement les **pièces annexes IS (Modèle Comptable Normal, loi 9-88 Maroc)**
en classeur Excel structuré avec formules dynamiques.

## ✨ Fonctionnalités

- **Universel** : fonctionne sur n'importe quel PDF MCN, quel que soit le logiciel de génération
- **Algorithme par coordonnées X/Y** : lit les tableaux comme un humain, indépendant de la structure PDF
- **Template Excel** : formules déjà en place, on injecte seulement les valeurs
- **229 formules intactes** : NET = Brut - Amort, Totaux, Résultats, liens inter-feuilles

## 🚀 Installation et lancement

```bash
git clone https://github.com/votre-compte/fiscalxl.git
cd fiscalxl
pip install -r requirements.txt
streamlit run app.py
```

## 📁 Structure

```
fiscalxl/
├── app.py                  ← Interface Streamlit
├── template_fiscal.xlsx    ← Template Excel avec formules
├── requirements.txt
├── core/
│   ├── pdf_parser.py       ← Extraction universelle par coordonnées X/Y
│   └── injector.py         ← Injection dans le template + mapping labels
└── utils/
    ├── validator.py
    └── logger.py
```

## 🔧 Algorithme d'extraction

1. **`extract_words()`** → liste de mots avec position (x0, x1, y)
2. **Regroupement par Y** (tolérance 6pt) → lignes logiques
3. **Détection automatique du seuil** label/nombre par clustering X
4. **Fusion des tokens** numériques adjacents (gap < 6pt) → reconstruit `1 234 567,89` → `1234567.89`
5. **Assignation aux colonnes** par ordre X → [brut, amort, net_n1] ou [val_n, val_n1] ou [propre_n, prec_n, total_n1]

## 📋 Format supporté

Tout PDF **pièces annexes IS Modèle Normal** :
- Page 1 : Infos générales
- Page 2 : Bilan Actif
- Page 3 : Bilan Passif  
- Pages 4-5 : CPC
