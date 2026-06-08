"""Enrichissement des CSV : ajout dimensions Client (CA) et Centre de Coût (charges).
   Préserve TOUS les totaux par BU × Poste × Mois."""
import pandas as pd
import numpy as np
import os
 
np.random.seed(42)
DATA_DIR = "./data"
 
CLIENTS_PAR_CANAL = {
    "GMS":           ["Carrefour","Auchan","Leclerc","Intermarché","Système U","Casino","Monoprix","Cora","Géant Casino"],
    "MDD":           ["MR Leclerc","Carrefour Bio","Carrefour Classic","Auchan Mmm!","Auchan Bio","Casino Bio","U Bio","Lidl Deluxe","Aldi Premium"],
    "B2B Industriel":["Sodexo","Compass France","Elior","Metro Cash","Brake France","Pomona","Transgourmet","Promocash"],
    "Export":        ["Carrefour Espagne","Aldi Allemagne","Tesco UK","Migros CH","Esselunga IT","Albert Heijn NL","El Corte Inglés","Conad IT"],
    "CHR":           ["Accor","Marriott","Eurest","Sodexo HR","Korian","Elior Resto","Disneyland Paris","Club Med"],
}
 
# Pour chaque poste de charge : liste de (CentreCout, poids cible)
CENTRES_PAR_POSTE = {
    "Coût des matières":     [("Production",0.85),("Qualité",0.10),("R&D",0.05)],
    "Charges de personnel":  [("Production",0.55),("Maintenance",0.10),("Qualité",0.07),("Logistique",0.10),("Commercial",0.08),("Administratif",0.10)],
    "Charges externes":      [("Maintenance",0.40),("Qualité",0.20),("Administratif",0.40)],
    "Transport & logistique":[("Logistique",0.95),("Production",0.05)],
    "Marketing & publicité": [("Commercial",1.0)],
    "Autres charges":        [("Administratif",0.6),("Direction",0.4)],
    "Amortissements":        [("Production",0.7),("Maintenance",0.3)],
    "Charges financières":   [("Direction",1.0)],
    "Impôt":                 [("Direction",1.0)],
}
 
N_CLIENTS_PAR_LIGNE = 5  # nombre de clients tirés par ligne de CA
 
def exploser_ca(ca_df):
    """Pour chaque ligne CA, tire 5 clients du canal et répartit montant + quantité."""
    blocs = []
    for canal, pool in CLIENTS_PAR_CANAL.items():
        sub = ca_df[ca_df["Canal"] == canal]
        if len(sub) == 0:
            continue
        ns = len(sub); pool_arr = np.array(pool); n_pool = len(pool_arr)
        n_pick = min(N_CLIENTS_PAR_LIGNE, n_pool)
        rand = np.random.random((ns, n_pool))
        idx = np.argsort(rand, axis=1)[:, :n_pick]
        clients = pool_arr[idx]
        parts = np.random.dirichlet(np.ones(n_pick) * 2.5, size=ns)
        # Conservation EXACTE : arrondir les n-1 premières colonnes, la dernière absorbe le reste
        m = sub["Montant_kEUR"].values[:, None] * parts
        m[:, :-1] = m[:, :-1].round(2)
        m[:, -1] = sub["Montant_kEUR"].values - m[:, :-1].sum(axis=1)
        q = sub["Quantite"].values[:, None] * parts
        q[:, :-1] = q[:, :-1].round(0)
        q[:, -1] = sub["Quantite"].values - q[:, :-1].sum(axis=1)
        q = np.nan_to_num(q, nan=0).round(0).astype(int)
        rep = sub.loc[sub.index.repeat(n_pick)].reset_index(drop=True)
        rep["Client"] = clients.flatten()
        rep["CentreCout"] = ""
        rep["Montant_kEUR"] = m.flatten()
        rep["Quantite"] = q.flatten()
        rep["PrixUnitaire_EUR"] = (rep["Montant_kEUR"]*1000 / rep["Quantite"].replace(0,1)).round(2)
        blocs.append(rep)
    return pd.concat(blocs, ignore_index=True)
 
def exploser_charges(ch_df):
    """Pour chaque ligne charge, éclate selon les centres de coût du poste."""
    blocs = []
    for poste, centres in CENTRES_PAR_POSTE.items():
        sub = ch_df[ch_df["PosteP&L"] == poste]
        if len(sub) == 0:
            continue
        ns = len(sub); n_cc = len(centres)
        poids_cible = np.array([p for _, p in centres])
        poids = poids_cible[None, :] * np.random.uniform(0.85, 1.15, size=(ns, n_cc))
        poids = poids / poids.sum(axis=1, keepdims=True)
        # Conservation EXACTE : la dernière colonne absorbe le reste sans arrondi
        m = sub["Montant_kEUR"].values[:, None] * poids
        m[:, :-1] = m[:, :-1].round(2)
        m[:, -1] = sub["Montant_kEUR"].values - m[:, :-1].sum(axis=1)
        noms_cc = np.array([cc for cc, _ in centres])
        rep = sub.loc[sub.index.repeat(n_cc)].reset_index(drop=True)
        rep["Client"] = ""
        rep["CentreCout"] = np.tile(noms_cc, ns)
        rep["Montant_kEUR"] = m.flatten()
        blocs.append(rep)
    return pd.concat(blocs, ignore_index=True)
 
def traiter_fichier(chemin):
    df = pd.read_csv(chemin, sep=";", encoding="utf-8-sig")
    is_ca = df["PosteP&L"] == "Chiffre d'affaires"
    ca_res = exploser_ca(df[is_ca].copy())
    ch_res = exploser_charges(df[~is_ca].copy())
    res = pd.concat([ca_res, ch_res], ignore_index=True)
    # numéroter EcritureID
    res["EcritureID"] = ["EC" + str(i).zfill(8) for i in range(1, len(res)+1)]
    # ordre des colonnes
    cols = ["EcritureID","Scenario","Annee","Mois","Date","CodeEntite","Entite","Site",
            "CodeProduit","Produit","Canal","Compte","LibelleCompte",
            "PosteP&L","Quantite","PrixUnitaire_EUR","Montant_kEUR","Client","CentreCout"]
    res = res[cols]
    # vérification : totaux par BU × Poste identiques
    avant = df.groupby(["CodeEntite","PosteP&L"])["Montant_kEUR"].sum().round(2)
    apres = res.groupby(["CodeEntite","PosteP&L"])["Montant_kEUR"].sum().round(2)
    diff = (avant - apres).abs().max()
    return df, res, diff
 
for f in ["GL_Reel_2026.csv","GL_Budget_2026.csv","GL_Reel_2025.csv"]:
    chemin = os.path.join(DATA_DIR, f)
    df_in, df_out, diff = traiter_fichier(chemin)
    df_out.to_csv(chemin, sep=";", index=False, float_format="%.2f", encoding="utf-8-sig")
    print(f"{f:<22} {len(df_in):>7,} -> {len(df_out):>7,} lignes  |  écart max BU×Poste : {diff:.2f} k€")
 
# total
total = sum(len(pd.read_csv(os.path.join(DATA_DIR,f),sep=";")) for f in ["GL_Reel_2026.csv","GL_Budget_2026.csv","GL_Reel_2025.csv"])
print(f"\nTOTAL : {total:,} lignes sur les 3 fichiers")