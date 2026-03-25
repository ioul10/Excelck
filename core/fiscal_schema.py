"""
core/fiscal_schema.py
Schéma canonique du Modèle Comptable Normal (loi 9-88).
Chaque entrée = valeur par défaut 0 / None selon le type de ligne.

Structure :
  BILAN_ACTIF_SCHEMA  : (label, brut, amort, net_n, net_n1)
  BILAN_PASSIF_SCHEMA : (label, val_n, val_n1)
  CPC_SCHEMA          : (label, propre_n, prec_n, total_n, total_n1)

Les lignes de section/regroupement ont toutes les valeurs à None.
Les lignes de total ont des valeurs à None → remplacées par formules Excel.
"""

# ── Marqueurs de type de ligne ────────────────────────────────────────────────
# Utilisés par ExcelBuilder pour choisir le style et les formules

SECTION_LABELS = {
    # Bilan Actif
    "ACTIF IMMOBILISÉ",
    "ACTIF CIRCULANT",
    "TRÉSORERIE - ACTIF",
    # Bilan Passif
    "FINANCEMENT PERMANENT",
    "CAPITAUX PROPRES",
    "Capitaux propres assimilés (B)",
    "Dettes du passif circulant (F)",
    "PASSIF CIRCULANT",
    "TRÉSORERIE PASSIF",
    # CPC
    "PRODUITS D'EXPLOITATION",
    "CHARGES D'EXPLOITATION",
    "PRODUITS FINANCIERS",
    "CHARGES FINANCIÈRES",
    "PRODUITS NON COURANTS",
    "CHARGES NON COURANTS",
}

TOTAL_LABELS = {
    # Bilan Actif
    "TOTAL I (A+B+C+D+E)",
    "TOTAL II (F+G+H+I)",
    "TOTAL III",
    "TOTAL GÉNÉRAL (I+II+III)",
    # Bilan Passif
    "Total des capitaux propres (A)",
    "TOTAL I (A+B+C+D+E) PASSIF",
    "TOTAL II (F+G+H)",
    "TOTAL III PASSIF",
    "TOTAL GÉNÉRAL PASSIF (I+II+III)",
    # CPC
    "TOTAL I - PRODUITS D'EXPLOITATION",
    "TOTAL II - CHARGES D'EXPLOITATION",
    "RÉSULTAT D'EXPLOITATION (I-II)",
    "TOTAL IV - PRODUITS FINANCIERS",
    "TOTAL V - CHARGES FINANCIÈRES",
    "RÉSULTAT FINANCIER (IV-V)",
    "RÉSULTAT COURANT (III+VI)",
    "TOTAL VIII - PRODUITS NON COURANTS",
    "TOTAL IX - CHARGES NON COURANTS",
    "RÉSULTAT NON COURANT (VIII-IX)",
    "RÉSULTAT AVANT IMPÔTS (VII+X)",
    "IMPÔTS SUR LES BÉNÉFICES",
    "RÉSULTAT NET (XI-XII)",
    "TOTAL DES PRODUITS (I+IV+VIII)",
    "TOTAL DES CHARGES (II+V+IX+XIII)",
    "RÉSULTAT NET (Total produits - Total charges)",
}

RESULT_LABELS = {
    "RÉSULTAT D'EXPLOITATION (I-II)",
    "RÉSULTAT FINANCIER (IV-V)",
    "RÉSULTAT COURANT (III+VI)",
    "RÉSULTAT NON COURANT (VIII-IX)",
    "RÉSULTAT AVANT IMPÔTS (VII+X)",
    "RÉSULTAT NET (XI-XII)",
    "RÉSULTAT NET (Total produits - Total charges)",
}

SUBTOTAL_LABELS = {
    "Immobilisations en non-valeurs [A]",
    "Immobilisations incorporelles [B]",
    "Immobilisations corporelles [C]",
    "Immobilisations financières [D]",
    "Écarts de conversion actif [E]",
    "Stocks [F]",
    "Créances de l'actif circulant [G]",
    "Titres et valeurs de placement [H]",
    "Dettes de financement (C)",
    "Provisions durables pour risques et charges (D)",
    "Autres provisions pour risques et charges (G)",
}

# ── Schéma Bilan Actif ────────────────────────────────────────────────────────
# (label, brut, amort_prov, net_n, net_n1)
BILAN_ACTIF_SCHEMA = [
    # ── Actif Immobilisé ──
    ("ACTIF IMMOBILISÉ",                                         None, None, None, None),
    ("Immobilisations en non-valeurs [A]",                       0,    0,    None, 0),
    ("  Frais préliminaires",                                    0,    0,    None, 0),
    ("  Charges à répartir sur plusieurs exercices",             0,    0,    None, 0),
    ("  Primes de remboursement des obligations",                0,    0,    None, 0),
    ("Immobilisations incorporelles [B]",                        None, None, None, None),
    ("  Immobilisations en Recherche et Développement",          0,    0,    None, 0),
    ("  Brevets, marques, droits et valeurs similaires",         0,    0,    None, 0),
    ("  Fonds commercial",                                       0,    0,    None, 0),
    ("  Autres immobilisations incorporelles",                   0,    0,    None, 0),
    ("Immobilisations corporelles [C]",                          None, None, None, None),
    ("  Terrains",                                               0,    0,    None, 0),
    ("  Constructions",                                          0,    0,    None, 0),
    ("  Installations techniques, matériel et outillage",        0,    0,    None, 0),
    ("  Matériel de transport",                                  0,    0,    None, 0),
    ("  Mobilier, Mat. de bureau, Aménagements divers",          0,    0,    None, 0),
    ("  Autres immobilisations corporelles",                     0,    0,    None, 0),
    ("  Immobilisations corporelles en cours",                   0,    0,    None, 0),
    ("Immobilisations financières [D]",                          None, None, None, None),
    ("  Prêts immobilisés",                                      0,    0,    None, 0),
    ("  Autres créances financières",                            0,    0,    None, 0),
    ("  Titres de participation",                                0,    0,    None, 0),
    ("  Autres titres immobilisés",                              0,    0,    None, 0),
    ("Écarts de conversion actif [E]",                           0,    0,    None, 0),
    ("  Diminution des créances immobilisées",                   0,    0,    None, 0),
    ("  Augmentation des dettes financières",                    0,    0,    None, 0),
    ("TOTAL I (A+B+C+D+E)",                                      None, None, None, None),
    # ── Actif Circulant ──
    ("ACTIF CIRCULANT",                                          None, None, None, None),
    ("Stocks [F]",                                               None, None, None, None),
    ("  Marchandises",                                           0,    0,    None, 0),
    ("  Matières et fournitures consommables",                   0,    0,    None, 0),
    ("  Produits en cours",                                      0,    0,    None, 0),
    ("  Produits intermédiaires et produits résiduels",          0,    0,    None, 0),
    ("  Produits finis",                                         0,    0,    None, 0),
    ("Créances de l'actif circulant [G]",                        None, None, None, None),
    ("  Fournisseurs débiteurs, avances et acomptes",            0,    0,    None, 0),
    ("  Clients et comptes rattachés",                           0,    0,    None, 0),
    ("  Personnel",                                              0,    0,    None, 0),
    ("  État",                                                   0,    0,    None, 0),
    ("  Comptes d'associés",                                     0,    0,    None, 0),
    ("  Autres débiteurs",                                       0,    0,    None, 0),
    ("  Comptes de régularisation - Actif",                      0,    0,    None, 0),
    ("Titres et valeurs de placement [H]",                       0,    0,    None, 0),
    ("Écarts de conversion actif - Éléments circulants [I]",     0,    0,    None, 0),
    ("TOTAL II (F+G+H+I)",                                       None, None, None, None),
    # ── Trésorerie ──
    ("TRÉSORERIE - ACTIF",                                       None, None, None, None),
    ("  Chèques et valeurs à encaisser",                         0,    0,    None, 0),
    ("  Banques, T.G et C.C.P",                                  0,    0,    None, 0),
    ("  Caisse, Régie d'avances et accréditifs",                 0,    0,    None, 0),
    ("TOTAL III",                                                None, None, None, None),
    ("TOTAL GÉNÉRAL (I+II+III)",                                 None, None, None, None),
]

# ── Schéma Bilan Passif ───────────────────────────────────────────────────────
# (label, val_n, val_n1)
BILAN_PASSIF_SCHEMA = [
    # ── Financement Permanent ──
    ("FINANCEMENT PERMANENT",                                    None, None),
    ("CAPITAUX PROPRES",                                         None, None),
    ("Capital social ou personnel",                              0,    0),
    ("Prime d'émission, de fusion, d'apport",                   0,    0),
    ("Écarts de réévaluation",                                   0,    0),
    ("Réserve légale",                                           0,    0),
    ("Autres réserves",                                          0,    0),
    ("Report à nouveau",                                         0,    0),
    ("Résultat en instance d'affectation",                       0,    0),
    ("Résultat net de l'exercice",                               0,    0),
    ("Total des capitaux propres (A)",                           None, None),
    ("Capitaux propres assimilés (B)",                           None, None),
    ("  Subventions d'investissement",                           0,    0),
    ("  Provisions réglementées",                                0,    0),
    ("Dettes de financement (C)",                                0,    0),
    ("  Emprunts obligataires",                                  0,    0),
    ("  Autres dettes de financement",                           0,    0),
    ("Provisions durables pour risques et charges (D)",          0,    0),
    ("  Provisions pour risques",                                0,    0),
    ("  Provisions pour charges",                                0,    0),
    ("Écarts de conversion - passif (E)",                        0,    0),
    ("  Augmentation des créances immobilisées",                 0,    0),
    ("  Diminution des dettes de financement",                   0,    0),
    ("TOTAL I (A+B+C+D+E) PASSIF",                              None, None),
    # ── Passif Circulant ──
    ("PASSIF CIRCULANT",                                         None, None),
    ("Dettes du passif circulant (F)",                           None, None),
    ("  Fournisseurs et comptes rattachés",                      0,    0),
    ("  Clients créditeurs, avances et acomptes",                0,    0),
    ("  Personnel",                                              0,    0),
    ("  Organismes sociaux",                                     0,    0),
    ("  État",                                                   0,    0),
    ("  Comptes d'associés",                                     0,    0),
    ("  Autres créanciers",                                      0,    0),
    ("  Comptes de régularisation passif",                       0,    0),
    ("Autres provisions pour risques et charges (G)",            0,    0),
    ("Écarts de conversion - passif éléments circulants (H)",    0,    0),
    ("TOTAL II (F+G+H)",                                         None, None),
    # ── Trésorerie Passif ──
    ("TRÉSORERIE PASSIF",                                        None, None),
    ("  Crédits d'escompte",                                     0,    0),
    ("  Crédits de trésorerie",                                  0,    0),
    ("  Banques (soldes créditeurs)",                            0,    0),
    ("TOTAL III PASSIF",                                         None, None),
    ("TOTAL GÉNÉRAL PASSIF (I+II+III)",                         None, None),
]

# ── Schéma CPC ────────────────────────────────────────────────────────────────
# (label, propre_n, exerc_prec, total_n, total_n1)
CPC_SCHEMA = [
    # ── Exploitation ──
    ("PRODUITS D'EXPLOITATION",                                  None, None, None, None),
    ("Ventes de marchandises (en l'état)",                       0,    0,    None, 0),
    ("Ventes de biens et services produits",                     0,    0,    None, 0),
    ("  Chiffre d'affaires",                                     0,    0,    None, 0),
    ("  Variation de stocks de produits",                        0,    0,    None, 0),
    ("  Immobilisations produites par l'entreprise",             0,    0,    None, 0),
    ("Subventions d'exploitation",                               0,    0,    None, 0),
    ("Autres produits d'exploitation",                           0,    0,    None, 0),
    ("Reprises d'exploitation : transferts de charges",          0,    0,    None, 0),
    ("TOTAL I - PRODUITS D'EXPLOITATION",                        None, None, None, None),
    ("CHARGES D'EXPLOITATION",                                   None, None, None, None),
    ("Achats revendus de marchandises",                          0,    0,    None, 0),
    ("Achats consommés de matières et fournitures",              0,    0,    None, 0),
    ("Autres charges externes",                                  0,    0,    None, 0),
    ("Impôts et taxes",                                          0,    0,    None, 0),
    ("Charges de personnel",                                     0,    0,    None, 0),
    ("Autres charges d'exploitation",                            0,    0,    None, 0),
    ("Dotations d'exploitation",                                 0,    0,    None, 0),
    ("TOTAL II - CHARGES D'EXPLOITATION",                        None, None, None, None),
    ("RÉSULTAT D'EXPLOITATION (I-II)",                           None, None, None, None),
    # ── Financier ──
    ("PRODUITS FINANCIERS",                                      None, None, None, None),
    ("Produits des titres de participation",                     0,    0,    None, 0),
    ("Gains de change",                                          0,    0,    None, 0),
    ("Intérêts et autres produits financiers",                   0,    0,    None, 0),
    ("Reprises financières : transferts de charges",             0,    0,    None, 0),
    ("TOTAL IV - PRODUITS FINANCIERS",                           None, None, None, None),
    ("CHARGES FINANCIÈRES",                                      None, None, None, None),
    ("Charges d'intérêts",                                       0,    0,    None, 0),
    ("Pertes de change",                                         0,    0,    None, 0),
    ("Autres charges financières",                               0,    0,    None, 0),
    ("Dotations financières",                                    0,    0,    None, 0),
    ("TOTAL V - CHARGES FINANCIÈRES",                            None, None, None, None),
    ("RÉSULTAT FINANCIER (IV-V)",                                None, None, None, None),
    ("RÉSULTAT COURANT (III+VI)",                                None, None, None, None),
    # ── Non courant ──
    ("PRODUITS NON COURANTS",                                    None, None, None, None),
    ("Produits des cessions d'immobilisations",                  0,    0,    None, 0),
    ("Subventions d'équilibre",                                  0,    0,    None, 0),
    ("Reprises sur subventions d'investissement",                0,    0,    None, 0),
    ("Autres produits non courants",                             0,    0,    None, 0),
    ("Reprises non courantes ; transferts de charges",           0,    0,    None, 0),
    ("TOTAL VIII - PRODUITS NON COURANTS",                       None, None, None, None),
    ("CHARGES NON COURANTS",                                     None, None, None, None),
    ("Val. nettes d'amort. des immobilisations cédées",          0,    0,    None, 0),
    ("Subventions accordées",                                    0,    0,    None, 0),
    ("Autres charges non courantes",                             0,    0,    None, 0),
    ("Dotations non courantes aux amort. et provisions",         0,    0,    None, 0),
    ("TOTAL IX - CHARGES NON COURANTS",                          None, None, None, None),
    ("RÉSULTAT NON COURANT (VIII-IX)",                           None, None, None, None),
    ("RÉSULTAT AVANT IMPÔTS (VII+X)",                            None, None, None, None),
    ("IMPÔTS SUR LES BÉNÉFICES",                                 0,    0,    None, 0),
    ("RÉSULTAT NET (XI-XII)",                                    None, None, None, None),
    ("TOTAL DES PRODUITS (I+IV+VIII)",                           None, None, None, None),
    ("TOTAL DES CHARGES (II+V+IX+XIII)",                         None, None, None, None),
    ("RÉSULTAT NET (Total produits - Total charges)",            None, None, None, None),
]

# ── Relations / Formules ──────────────────────────────────────────────────────
# Définit les formules de total : {label_total: [label_composants]}
# Utilisé par ExcelBuilder pour écrire des formules =SUM(ref1+ref2+...)

ACTIF_FORMULAS = {
    "TOTAL I (A+B+C+D+E)": [
        "Immobilisations en non-valeurs [A]",
        "Immobilisations incorporelles [B]",
        "Immobilisations corporelles [C]",
        "Immobilisations financières [D]",
        "Écarts de conversion actif [E]",
    ],
    "TOTAL II (F+G+H+I)": [
        "Stocks [F]",
        "Créances de l'actif circulant [G]",
        "Titres et valeurs de placement [H]",
        "Écarts de conversion actif - Éléments circulants [I]",
    ],
    "TOTAL III": [
        "  Chèques et valeurs à encaisser",
        "  Banques, T.G et C.C.P",
        "  Caisse, Régie d'avances et accréditifs",
    ],
    "TOTAL GÉNÉRAL (I+II+III)": [
        "TOTAL I (A+B+C+D+E)",
        "TOTAL II (F+G+H+I)",
        "TOTAL III",
    ],
    # Sous-totaux
    "Immobilisations en non-valeurs [A]": [
        "  Frais préliminaires",
        "  Charges à répartir sur plusieurs exercices",
        "  Primes de remboursement des obligations",
    ],
    "Immobilisations incorporelles [B]": [
        "  Immobilisations en Recherche et Développement",
        "  Brevets, marques, droits et valeurs similaires",
        "  Fonds commercial",
        "  Autres immobilisations incorporelles",
    ],
    "Immobilisations corporelles [C]": [
        "  Terrains",
        "  Constructions",
        "  Installations techniques, matériel et outillage",
        "  Matériel de transport",
        "  Mobilier, Mat. de bureau, Aménagements divers",
        "  Autres immobilisations corporelles",
        "  Immobilisations corporelles en cours",
    ],
    "Immobilisations financières [D]": [
        "  Prêts immobilisés",
        "  Autres créances financières",
        "  Titres de participation",
        "  Autres titres immobilisés",
    ],
    "Stocks [F]": [
        "  Marchandises",
        "  Matières et fournitures consommables",
        "  Produits en cours",
        "  Produits intermédiaires et produits résiduels",
        "  Produits finis",
    ],
    "Créances de l'actif circulant [G]": [
        "  Fournisseurs débiteurs, avances et acomptes",
        "  Clients et comptes rattachés",
        "  Personnel",
        "  État",
        "  Comptes d'associés",
        "  Autres débiteurs",
        "  Comptes de régularisation - Actif",
    ],
}

PASSIF_FORMULAS = {
    "Total des capitaux propres (A)": [
        "Capital social ou personnel",
        "Prime d'émission, de fusion, d'apport",
        "Écarts de réévaluation",
        "Réserve légale",
        "Autres réserves",
        "Report à nouveau",
        "Résultat en instance d'affectation",
        "Résultat net de l'exercice",
    ],
    "TOTAL I (A+B+C+D+E) PASSIF": [
        "Total des capitaux propres (A)",
        "  Subventions d'investissement",
        "  Provisions réglementées",
        "Dettes de financement (C)",
        "Provisions durables pour risques et charges (D)",
        "Écarts de conversion - passif (E)",
    ],
    "TOTAL II (F+G+H)": [
        "  Fournisseurs et comptes rattachés",
        "  Clients créditeurs, avances et acomptes",
        "  Personnel",
        "  Organismes sociaux",
        "  État",
        "  Comptes d'associés",
        "  Autres créanciers",
        "  Comptes de régularisation passif",
        "Autres provisions pour risques et charges (G)",
        "Écarts de conversion - passif éléments circulants (H)",
    ],
    "TOTAL III PASSIF": [
        "  Crédits d'escompte",
        "  Crédits de trésorerie",
        "  Banques (soldes créditeurs)",
    ],
    "TOTAL GÉNÉRAL PASSIF (I+II+III)": [
        "TOTAL I (A+B+C+D+E) PASSIF",
        "TOTAL II (F+G+H)",
        "TOTAL III PASSIF",
    ],
}

CPC_FORMULAS = {
    "TOTAL I - PRODUITS D'EXPLOITATION": [
        "Ventes de marchandises (en l'état)",
        "Ventes de biens et services produits",
        "Subventions d'exploitation",
        "Autres produits d'exploitation",
        "Reprises d'exploitation : transferts de charges",
    ],
    "TOTAL II - CHARGES D'EXPLOITATION": [
        "Achats revendus de marchandises",
        "Achats consommés de matières et fournitures",
        "Autres charges externes",
        "Impôts et taxes",
        "Charges de personnel",
        "Autres charges d'exploitation",
        "Dotations d'exploitation",
    ],
    "TOTAL IV - PRODUITS FINANCIERS": [
        "Produits des titres de participation",
        "Gains de change",
        "Intérêts et autres produits financiers",
        "Reprises financières : transferts de charges",
    ],
    "TOTAL V - CHARGES FINANCIÈRES": [
        "Charges d'intérêts",
        "Pertes de change",
        "Autres charges financières",
        "Dotations financières",
    ],
    "TOTAL VIII - PRODUITS NON COURANTS": [
        "Produits des cessions d'immobilisations",
        "Subventions d'équilibre",
        "Reprises sur subventions d'investissement",
        "Autres produits non courants",
        "Reprises non courantes ; transferts de charges",
    ],
    "TOTAL IX - CHARGES NON COURANTS": [
        "Val. nettes d'amort. des immobilisations cédées",
        "Subventions accordées",
        "Autres charges non courantes",
        "Dotations non courantes aux amort. et provisions",
    ],
    "TOTAL DES PRODUITS (I+IV+VIII)": [
        "TOTAL I - PRODUITS D'EXPLOITATION",
        "TOTAL IV - PRODUITS FINANCIERS",
        "TOTAL VIII - PRODUITS NON COURANTS",
    ],
    "TOTAL DES CHARGES (II+V+IX+XIII)": [
        "TOTAL II - CHARGES D'EXPLOITATION",
        "TOTAL V - CHARGES FINANCIÈRES",
        "TOTAL IX - CHARGES NON COURANTS",
        "IMPÔTS SUR LES BÉNÉFICES",
    ],
}

# Différences (résultats) : {label: (positif, négatif)}
CPC_DIFFERENCES = {
    "RÉSULTAT D'EXPLOITATION (I-II)": (
        "TOTAL I - PRODUITS D'EXPLOITATION",
        "TOTAL II - CHARGES D'EXPLOITATION",
    ),
    "RÉSULTAT FINANCIER (IV-V)": (
        "TOTAL IV - PRODUITS FINANCIERS",
        "TOTAL V - CHARGES FINANCIÈRES",
    ),
    "RÉSULTAT COURANT (III+VI)": (
        "RÉSULTAT D'EXPLOITATION (I-II)",
        "RÉSULTAT FINANCIER (IV-V)",
        # addition ici, géré dans ExcelBuilder
    ),
    "RÉSULTAT NON COURANT (VIII-IX)": (
        "TOTAL VIII - PRODUITS NON COURANTS",
        "TOTAL IX - CHARGES NON COURANTS",
    ),
    "RÉSULTAT AVANT IMPÔTS (VII+X)": (
        "RÉSULTAT COURANT (III+VI)",
        "RÉSULTAT NON COURANT (VIII-IX)",
        # addition ici
    ),
    "RÉSULTAT NET (XI-XII)": (
        "RÉSULTAT AVANT IMPÔTS (VII+X)",
        "IMPÔTS SUR LES BÉNÉFICES",
    ),
    "RÉSULTAT NET (Total produits - Total charges)": (
        "TOTAL DES PRODUITS (I+IV+VIII)",
        "TOTAL DES CHARGES (II+V+IX+XIII)",
    ),
}

# Résultats construits par addition (pas soustraction)
CPC_ADDITIONS = {
    "RÉSULTAT COURANT (III+VI)",
    "RÉSULTAT AVANT IMPÔTS (VII+X)",
}
