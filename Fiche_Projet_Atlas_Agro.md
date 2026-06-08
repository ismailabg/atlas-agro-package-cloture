# Atlas Agro Holding — Package de clôture automatisé

## Fiche projet | Ismail Abgar

---

### Le projet en une phrase

Un modèle complet de contrôle de gestion qui automatise le cycle de clôture mensuel : de l'extraction ERP (578 000 lignes) jusqu'au package PDF diffusable, en passant par un dashboard interactif, des commentaires de gestion générés par intelligence artificielle, et une industrialisation complète en VBA et Python.

---

### Le contexte

Atlas Agro Holding est un groupe agroalimentaire fictif (210 M€ de CA, 5 Business Units, 850 collaborateurs, 8 sites industriels). Le projet simule le travail d'un contrôleur de gestion rattaché à la Direction Financière, responsable du reporting mensuel de clôture.

---

### Ce que le modèle fait

**Données**
- 3 extractions ERP (Réel 2026, Budget 2026, Réel 2025) totalisant 578 340 lignes
- 19 colonnes analytiques : entité, site, produit, canal, client (42 clients), centre de coût (8 centres)
- Granularité : BU × Site × Produit × Canal × Client / Centre de coût × Mois

**Power Query (ETL)**
- 7 requêtes interconnectées : 3 imports, 1 agrégation BU×Poste, 1 mensuelle, 1 Top Clients, 1 Charges par centre de coût
- Paramètre centralisé : le mois de clôture est lu depuis une cellule ; un changement + actualisation rebascule tout le modèle
- Volume/Prix/Mix : décomposition automatique des écarts de CA en 3 effets

**Modèle Excel**
- Dashboard "Command Center" : 3 KPI vedettes + jauge d'atterrissage, 5 KPI secondaires, 11 faits marquants auto-générés, 2 courbes de tendance mensuelle, KPI financiers, CA par BU, donut contribution, Top 10 Clients, charges par centre de coût, heatmap des écarts, repères groupe
- Package P&L : compte de résultat consolidé, analyse verticale (% du CA), indicateurs de matérialité, bridge du résultat net, décomposition Volume/Prix/Mix
- Synthèse BU : 9 blocs analytiques (performance YTD, mois, contribution, structure de coûts, commentaires de gestion, bridge EBITDA, matrice stratégique, heatmap, Top/Flop)

**VBA (automatisation)**
- Macro "Générer le package" : actualise les données, recalcule, horodate, exporte Dashboard + Package + Synthèse en PDF, archive avec nommage daté
- Formulaire de pilotage : interface guidée en 5 étapes (mois, seuil, options, confirmation, exécution) — utilisable par un non-technicien
- Import dynamique : sélecteur de dossier, détection des CSV attendus, comptage des lignes, copie et actualisation automatique

**Python + API Claude (IA)**
- Script qui lit les CSV, reproduit l'agrégation, détecte les écarts matériels, et génère des commentaires de gestion via l'API Claude
- 3 sorties : rapport PDF, document Word éditable, injection dans le bloc commentaires de la synthèse Excel
- Générateur de secours intégré (fonctionne sans clé API)

---

### Technologies utilisées

| Domaine | Outils |
|---|---|
| Modélisation | Excel (formules avancées, graphiques, mise en forme conditionnelle) |
| ETL | Power Query (code M, requêtes paramétrées, agrégations) |
| Automatisation | VBA (macros, formulaire, import, gestion d'erreurs) |
| Data / IA | Python (pandas, reportlab, python-docx, openpyxl, API Claude) |
| Volume | 578 340 lignes, 42 clients, 8 centres de coûts, 29 SKU, 5 canaux |

---

### Compétences démontrées

- Construction d'un modèle de reporting structuré (P&L, écarts, bridge, matérialité)
- Maîtrise de Power Query pour l'agrégation de gros volumes
- Décomposition Volume/Prix/Mix des écarts de CA
- Automatisation VBA de bout en bout (industrialisation du processus)
- Intégration d'une API LLM dans un workflow financier
- Construction d'un dashboard analytique riche et lisible

---

### Comment l'utiliser

1. Placer les 3 CSV dans le dossier `data/`
2. Ouvrir le fichier `.xlsm`, activer les macros
3. Cliquer sur "Pilotage clôture" : choisir le mois, le seuil, actualiser et exporter
4. Le PDF est généré dans `Archives_Cloture/`
5. (Optionnel) Lancer le script Python pour les commentaires IA

---

### À propos

**Ismail Abgar**
Master 1 Financial & Risk Management — ISC Paris
Expérience : stage en contrôle de gestion (MARJANE, Maroc)
Recherche : alternance en contrôle de gestion

LinkedIn : [votre lien]
Email : [votre email]
GitHub : [votre lien]
