"""
ATLAS AGRO HOLDING — Générateur d'extractions ERP volumineuses (v2)
Dimensions ajoutées : Site, Produit/SKU, Canal de vente, Quantité, Prix unitaire.
Plan de comptes étendu (~36 comptes). Objectif : ~150 000 écritures.
Sortie : 3 fichiers CSV (Réel 2025, Budget 2026, Réel 2026 Jan-Mar).
"""
import csv, random, os
random.seed(42)

OUT = "./DATA"
os.makedirs(OUT, exist_ok=True)

# ---------------- Référentiels ----------------
ENTITES = {
    "BU01": ("Céréales & Meunerie",      ["Roubaix","Reims"]),
    "BU02": ("Huiles & Corps Gras",      ["Lille","Casablanca"]),
    "BU03": ("Lait & Produits Laitiers", ["Arras","Lille"]),
    "BU04": ("Sucre & Conserves",        ["Amiens","Compiègne"]),
    "BU05": ("Boissons & Jus",           ["Béthune"]),
}

# Catalogue produits : code, libellé, poids dans le CA de la BU, prix unitaire (€)
PRODUITS = {
 "BU01":[("P101","Farine T55 1kg",0.22,0.95),("P102","Farine T65 1kg",0.15,1.05),
         ("P103","Semoule fine 1kg",0.12,1.40),("P104","Flocons avoine 500g",0.13,1.80),
         ("P105","Céréales petit-déj 375g",0.20,2.60),("P106","Mix boulanger 25kg",0.18,18.50)],
 "BU02":[("P201","Huile tournesol 1L",0.30,2.20),("P202","Huile colza 1L",0.20,2.40),
         ("P203","Huile olive 75cl",0.22,6.90),("P204","Margarine 250g",0.16,1.30),
         ("P205","Margarine pro 2kg",0.12,7.50)],
 "BU03":[("P301","Lait UHT entier 1L",0.24,0.89),("P302","Lait demi-écrémé 1L",0.20,0.85),
         ("P303","Yaourt nature x4",0.16,1.60),("P304","Yaourt fruits x8",0.18,2.95),
         ("P305","Fromage frais 200g",0.12,1.75),("P306","Crème fraîche 20cl",0.10,1.15)],
 "BU04":[("P401","Sucre blanc 1kg",0.26,1.05),("P402","Sucre poudre 500g",0.14,0.75),
         ("P403","Confiture fraise 370g",0.18,2.40),("P404","Confiture abricot 370g",0.14,2.40),
         ("P405","Conserve haricots 800g",0.16,1.30),("P406","Conserve fruits 500g",0.12,1.95)],
 "BU05":[("P501","Jus orange 1L",0.28,1.85),("P502","Jus pomme 1L",0.20,1.65),
         ("P503","Nectar multifruits 1L",0.18,1.55),("P504","Eau aromatisée citron 1.5L",0.18,0.95),
         ("P505","Eau aromatisée menthe 1.5L",0.16,0.95)],
}

# Canaux de vente : libellé, poids moyen, indice de prix (remise/surcote)
CANAUX = [("GMS",0.45,1.00),("MDD",0.25,0.82),("B2B Industriel",0.15,0.90),
          ("Export",0.07,0.95),("CHR",0.08,1.12)]

# Plan de comptes étendu : (compte, libellé, poste P&L, type P/C, clé structure coût)
COMPTES_PRODUITS = [
    ("701000","Ventes produits finis","Chiffre d'affaires"),
    ("707000","Ventes de marchandises","Chiffre d'affaires"),
    ("706000","Prestations de services","Chiffre d'affaires"),
    ("709000","RRR accordés","Chiffre d'affaires"),
]
# charges : (compte, libellé, poste, clé)
COMPTES_CHARGES = [
    ("601000","Achats matières premières","Coût des matières","mat"),
    ("601100","Achats ingrédients","Coût des matières","mat"),
    ("602000","Achats emballages","Coût des matières","mat"),
    ("604000","Achats sous-traitance","Coût des matières","mat"),
    ("603000","Variation de stocks","Coût des matières","mat"),
    ("641000","Salaires & traitements","Charges de personnel","pers"),
    ("641100","Primes & bonus","Charges de personnel","pers"),
    ("645000","Charges sociales","Charges de personnel","pers"),
    ("647000","Autres charges sociales","Charges de personnel","pers"),
    ("621000","Personnel intérimaire","Charges de personnel","pers"),
    ("606000","Achats non stockés (énergie)","Charges externes","ext"),
    ("613000","Locations","Charges externes","ext"),
    ("615000","Entretien & réparations","Charges externes","ext"),
    ("616000","Primes d'assurance","Charges externes","ext"),
    ("618000","Documentation & divers","Charges externes","ext"),
    ("624000","Transport sur ventes","Transport & logistique","transp"),
    ("624100","Transport sur achats","Transport & logistique","transp"),
    ("625000","Déplacements & missions","Transport & logistique","transp"),
    ("623000","Publicité","Marketing & publicité","mkt"),
    ("623100","Promotions","Marketing & publicité","mkt"),
    ("626000","Frais postaux & télécom","Marketing & publicité","mkt"),
    ("627000","Services bancaires","Marketing & publicité","mkt"),
    ("681000","Dotations amortissements","Amortissements","amort"),
    ("681100","Dotations provisions","Amortissements","amort"),
    ("658000","Autres charges de gestion","Autres charges","autres"),
    ("635000","Impôts & taxes","Autres charges","autres"),
    ("661000","Charges d'intérêts","Charges financières","fin"),
    ("668000","Autres charges financières","Charges financières","fin"),
    ("695000","Impôt sur les sociétés","Impôt","is"),
]

# ---------------- Hypothèses économiques (identiques v1 pour cohérence) ----------------
CA_BASE_2025 = {"BU01":4200,"BU02":3100,"BU03":5400,"BU04":2300,"BU05":1900}
SAISON = {
 "BU01":[1.02,0.98,1.00,1.01,1.03,1.05,0.95,0.92,1.00,1.04,1.05,0.95],
 "BU02":[0.95,0.93,0.98,1.02,1.05,1.08,1.10,1.08,1.02,0.98,0.92,0.89],
 "BU03":[1.00,0.97,1.01,1.03,1.06,1.08,1.05,1.00,0.99,0.98,0.96,0.97],
 "BU04":[0.88,0.85,0.92,0.95,1.00,1.05,1.10,1.12,1.08,1.05,1.10,1.30],
 "BU05":[0.80,0.82,0.95,1.05,1.20,1.40,1.55,1.50,1.10,0.85,0.75,0.78],
}
CROISSANCE_BUDGET = {"BU01":0.05,"BU02":0.08,"BU03":0.04,"BU04":0.10,"BU05":0.12}
STRUCT = {
 "BU01":dict(mat=0.62,pers=0.11,ext=0.05,transp=0.04,mkt=0.03,amort=0.04,autres=0.02,fin=0.015),
 "BU02":dict(mat=0.68,pers=0.08,ext=0.04,transp=0.05,mkt=0.02,amort=0.03,autres=0.02,fin=0.018),
 "BU03":dict(mat=0.58,pers=0.13,ext=0.06,transp=0.06,mkt=0.04,amort=0.05,autres=0.02,fin=0.012),
 "BU04":dict(mat=0.55,pers=0.12,ext=0.05,transp=0.04,mkt=0.05,amort=0.04,autres=0.03,fin=0.020),
 "BU05":dict(mat=0.48,pers=0.14,ext=0.07,transp=0.05,mkt=0.09,amort=0.06,autres=0.03,fin=0.015),
}
# Répartition d'un poste de charge sur ses comptes (poids)
def repartir(comptes_cle, total):
    n=len(comptes_cle)
    poids=[random.uniform(0.6,1.4) for _ in range(n)]
    s=sum(poids)
    return [total*p/s for p in poids]

HEADER=["EcritureID","Scenario","Annee","Mois","Date","CodeEntite","Entite","Site",
        "CodeProduit","Produit","Canal","Compte","LibelleCompte","PosteP&L",
        "Quantite","PrixUnitaire_EUR","Montant_kEUR"]

def jit(x,p): return x*(1+random.uniform(-p,p))

def gen(scenario, annee, mois_list, ca_ref, ecart=False):
    rows=[]; eid=1
    for bu,(libbu,sites) in ENTITES.items():
        prods=PRODUITS[bu]; struct=STRUCT[bu]
        for m in mois_list:
            ca_mois = ca_ref[bu]*SAISON[bu][m-1]
            if ecart:
                d=random.uniform(-0.08,0.06)
                if bu=="BU05" and m in (2,3): d-=0.10
                if bu=="BU04" and m==3: d+=0.09
                ca_mois*=(1+d)
            # ---- PRODUITS (ventes) ventilées par produit × canal ----
            for (pc,pl,pw,pu) in prods:
                for (canal,cw,cidx) in CANAUX:
                    ca_pc = ca_mois*pw*cw            # CA de la combinaison (k€)
                    if ca_pc<=0: continue
                    pu_eff = pu*cidx                  # prix effectif selon canal
                    # nombre de transactions (factures) sur le mois
                    ntx = random.randint(14,34)
                    for _ in range(ntx):
                        montant = jit(ca_pc/ntx,0.18)         # k€
                        qte = round(montant*1000/pu_eff,0)    # unités
                        jour=random.randint(1,28)
                        site=random.choice(sites)
                        # 92% ventes produits finis, sinon marchandises/prestations
                        r=random.random()
                        if r<0.90: cpt=("701000","Ventes produits finis")
                        elif r<0.96: cpt=("707000","Ventes de marchandises")
                        else: cpt=("706000","Prestations de services")
                        rows.append([eid,scenario,annee,m,f"{annee}-{m:02d}-{jour:02d}",
                                     bu,libbu,site,pc,pl,canal,cpt[0],cpt[1],
                                     "Chiffre d'affaires",int(qte),round(pu_eff,3),
                                     round(montant,2)])
                        eid+=1
                    # RRR (remises) : quelques lignes négatives par combo
                    if random.random()<0.5:
                        rrr=-jit(ca_pc*0.02,0.3)
                        rows.append([eid,scenario,annee,m,f"{annee}-{m:02d}-28",
                                     bu,libbu,random.choice(sites),pc,pl,canal,
                                     "709000","RRR accordés","Chiffre d'affaires",
                                     "","",round(rrr,2)])
                        eid+=1
            # ---- CHARGES ventilées par compte × site ----
            for poste_cle in ["mat","pers","ext","transp","mkt","amort","autres","fin","is"]:
                if poste_cle=="is":
                    rows.append([eid,scenario,annee,m,f"{annee}-{m:02d}-28",
                                 bu,libbu,sites[0],"P000","Commun","N/A",
                                 "695000","Impôt sur les sociétés","Impôt",
                                 "","",-round(ca_mois*0.03,2)]); eid+=1
                    continue
                taux=struct[poste_cle]
                if ecart:
                    if poste_cle=="mat": taux*=(1+random.uniform(-0.02,0.07))
                    elif poste_cle=="pers": taux*=(1+random.uniform(-0.01,0.06))
                    else: taux*=(1+random.uniform(-0.05,0.05))
                total_poste=ca_mois*taux
                comptes=[c for c in COMPTES_CHARGES if c[3]==poste_cle]
                montants=repartir(comptes,total_poste)
                for (cpt,montant) in zip(comptes,montants):
                    # éclater chaque compte en plusieurs écritures par site
                    ntx=random.randint(6,15)
                    for _ in range(ntx):
                        site=random.choice(sites)
                        part=-jit(montant/ntx,0.22)
                        jour=random.randint(1,28)
                        rows.append([eid,scenario,annee,m,f"{annee}-{m:02d}-{jour:02d}",
                                     bu,libbu,site,"P000","Commun","N/A",
                                     cpt[0],cpt[1],cpt[2],"","",round(part,2)])
                        eid+=1
    return rows

ca_budget={bu:CA_BASE_2025[bu]*(1+CROISSANCE_BUDGET[bu]) for bu in ENTITES}

budget   = gen("Budget",2026,list(range(1,13)),ca_budget,ecart=False)
reel2026 = gen("Réel",  2026,[1,2,3],          ca_budget,ecart=True)
reel2025 = gen("Réel",  2025,list(range(1,13)),CA_BASE_2025,ecart=True)

def write(path,rows):
    with open(path,"w",newline="",encoding="utf-8-sig") as f:
        w=csv.writer(f,delimiter=";"); w.writerow(HEADER); w.writerows(rows)

write(f"{OUT}/GL_Reel_2025.csv",reel2025)
write(f"{OUT}/GL_Budget_2026.csv",budget)
write(f"{OUT}/GL_Reel_2026.csv",reel2026)

tot=len(budget)+len(reel2026)+len(reel2025)
print(f"Réel 2025   : {len(reel2025):>7,} écritures")
print(f"Budget 2026 : {len(budget):>7,} écritures")
print(f"Réel 2026   : {len(reel2026):>7,} écritures  (Jan-Mar)")
print(f"{'TOTAL':<12}: {tot:>7,} écritures")