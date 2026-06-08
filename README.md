# Atlas Agro Holding — Package de clôture automatisé

Modèle de contrôle de gestion couvrant le cycle complet de clôture mensuel : extraction des données, agrégation, analyse des écarts, production du reporting et archivage.

Entreprise fictive : groupe agroalimentaire, 210 M€ de CA, 5 Business Units, 850 collaborateurs.

---

## Tester le modèle

**Test rapide** — Téléchargez le fichier Excel directement :
[Atlas_Agro_Command_Center.xlsm (Google Drive)](https://drive.google.com/drive/folders/1YY35f1PCzimXf0EJDPoDzaiWjpn2y-pb?usp=sharing)
## Tester le modèle

**Activer les macros :**

Si Excel affiche un bandeau jaune « Les macros ont été désactivées », cliquez sur Activer le contenu.

Si Excel affiche un bandeau rouge « Microsoft a bloqué l'exécution des macros », le fichier est marqué comme provenant d'internet. Deux méthodes pour débloquer :

Méthode 1 — Fermez Excel. Clic droit sur le fichier .xlsm dans l'Explorateur, Propriétés, cochez « Débloquer » en bas de l'onglet Général, Appliquer, OK. Rouvrez le fichier.

Méthode 2 — Dans Excel, allez dans Fichier, Options, Centre de gestion de la confidentialité, Paramètres du Centre de gestion de la confidentialité, Emplacements approuvés, Ajouter un nouvel emplacement. Indiquez le dossier où se trouve le fichier téléchargé (par exemple C:\Users\VotreNom\Downloads), cochez « Les sous-dossiers sont également approuvés », OK. Fermez et rouvrez le fichier.

**Explorer le modèle :**

- DASHBOARD — Le Command Center avec les KPI, tendances, faits marquants, heatmap
- PACKAGE_PnL — Le compte de résultat consolidé avec bridge et décomposition Volume/Prix/Mix
- SYNTHESE_BU — L'analyse par Business Unit (9 blocs analytiques)

Le bouton « Pilotage clôture » génère le PDF du package complet.

**Test complet** — Clonez le repository, générez les données (578 000 lignes), importez-les dans le modèle :
```
pip install pandas numpy
python scripts/01b_generate_bigdata.py
python scripts/15_exploser_csv.py
```
Ouvrez le fichier Excel, utilisez le bouton « Importer les données » pour charger les CSV, puis « Pilotage clôture » pour générer le package.

---

## Contenu du projet

**Données** — 578 340 écritures comptables sur 3 exercices (Réel 2026, Budget 2026, Réel 2025). 19 colonnes : entité, site, produit, canal, client (42), centre de coût (8). Générées par script Python.

**Power Query** — 7 requêtes interconnectées. Agrégation BU x Poste, Top Clients, Charges par centre de coût, décomposition mensuelle. Le mois de clôture est paramétré depuis une cellule unique ; l'actualisation rebascule tout le modèle.

**Modèle Excel** — Dashboard (KPI, tendances mensuelles, faits marquants auto-générés, heatmap, Top 10 Clients). Package P&L (compte de résultat, bridge du résultat net, analyse verticale, matérialité, décomposition Volume/Prix/Mix). Synthèse par BU (9 blocs analytiques dont matrice stratégique et bridge EBITDA).

**VBA** — Génération du package en PDF (actualisation, recalcul, export, archivage daté). Formulaire de pilotage guidé (choix du mois, seuil, options). Import dynamique des CSV (détection, validation, copie, actualisation).

**Python + API Claude** — Génération automatique des commentaires de gestion sur les écarts matériels. Sorties : PDF, Word, injection Excel. Fonctionne aussi sans clé API (générateur de secours par règles).

---

## Structure

```
Archives_Cloture/    PDF générés par la macro (archivage daté)
Commentaire/         Sorties du script Python (PDF, Word)
DATA/                Échantillons CSV (1 000 lignes)
Scripts/             Génération des données + commentaires IA
import/              CSV d'import pour tester le module d'import dynamique
vba/                 Modules VBA (.bas)
```

---

## Technologies

Excel, Power Query (M), VBA, Python, pandas, reportlab, python-docx, API Claude (Anthropic)

---

## Auteur

Ismail Abgar — M1 Financial & Risk Management, ISC Paris
En recherche d'alternance en contrôle de gestion.
