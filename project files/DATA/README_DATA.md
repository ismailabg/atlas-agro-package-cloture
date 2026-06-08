# Données du projet Atlas Agro

## Fichiers échantillons (inclus)

Les 3 fichiers CSV dans ce dossier contiennent un **échantillon de 1 000 lignes** chacun (pour consultation rapide).

| Fichier | Contenu | Lignes (échantillon) | Lignes (complet) |
|---|---|---|---|
| GL_Reel_2026.csv | Écritures réelles Jan-Avr 2026 | 1 000 | 84 910 |
| GL_Budget_2026.csv | Budget annuel 2026 | 1 000 | 258 101 |
| GL_Reel_2025.csv | Écritures réelles 2025 | 1 000 | 256 737 |

**Total complet : 578 340 lignes** (42 clients, 8 centres de coûts, 29 SKU, 5 canaux, 5 BU)

## Régénérer les données complètes

Pour recréer le jeu de données complet (578 000+ lignes), lancez les 2 scripts dans l'ordre :

```bash
python scripts/01b_generate_bigdata.py    # génère 132 667 lignes de base
python scripts/15_exploser_csv.py          # éclate en 578 340 lignes (clients + centres de coûts)
```

Prérequis : `pip install pandas numpy`

Les fichiers sont créés dans le dossier `data/`.

## Structure des colonnes (19 colonnes)

| Colonne | Type | Description |
|---|---|---|
| EcritureID | Texte | Identifiant unique (EC00000001) |
| Scenario | Texte | Réel ou Budget |
| Annee | Entier | 2025 ou 2026 |
| Mois | Entier | 1 à 12 |
| Date | Date | Date de l'écriture |
| CodeEntite | Texte | BU01 à BU05 |
| Entite | Texte | Nom de la Business Unit |
| Site | Texte | Site industriel (8 sites) |
| CodeProduit | Texte | P101 à P529 (29 SKU) |
| Produit | Texte | Nom du produit |
| Canal | Texte | GMS, MDD, B2B Industriel, Export, CHR |
| Compte | Texte | Numéro de compte comptable |
| LibelleCompte | Texte | Libellé du compte |
| PosteP&L | Texte | Poste du compte de résultat (10 postes) |
| Quantite | Entier | Volume (pour le CA) |
| PrixUnitaire_EUR | Décimal | Prix unitaire en EUR |
| Montant_kEUR | Décimal | Montant en milliers d'euros |
| Client | Texte | Nom du client (42 clients, pour les lignes CA) |
| CentreCout | Texte | Centre de coût (8 centres, pour les lignes charges) |
