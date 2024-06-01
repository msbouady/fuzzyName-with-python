import pandas as pd
from fuzzywuzzy import process, fuzz
path1 = ''
path2 =''


ss1 = pd.read_excel(path2)
ss2 = pd.read_excel(path2)

# Convertir les colonnes 'ncomplet' en listes
ncomplet_ss1 = ss1['ncomplet'].tolist()
ncomplet_ss2 = ss2['ncomplet'].tolist()

# Assurer que la colonne 'etat' est de type object (string)
ss1['etat'] = ''
ss2['etat'] = ''

def nettoyer_nom(nom):
    # Supprimer les points et les espaces superflus
    nom = nom.replace('.', '').strip()
    # Diviser le nom en parties pour permettre des comparaisons plus flexibles
    parties = nom.split()
    # Créer des variations du nom pour inclure les permutations
    variations = [' '.join(parties)]
    if len(parties) > 1:
        variations.append(' '.join(parties[::-1]))  # Ajouter l'inversion du nom
    return variations

def comparer_noms(nom, liste_noms):
    # Nettoyer le nom et créer ses variations
    variations_nom = nettoyer_nom(nom)
    best_score = 0
    best_match = None
    for variation in variations_nom:
        match = process.extractOne(variation, liste_noms, scorer=fuzz.ratio)
        if match and match[1] > best_score:
            best_score = match[1]
            best_match = match
    return best_match if best_score >= 70 else None

# Parcourir les 'ncomplet' de SS1 et vérifier dans SS2
for index, nom in enumerate(ncomplet_ss1):
    match = comparer_noms(nom, ncomplet_ss2)
    ss1.at[index, 'etat'] = 'admis' if match else 'sortant'

# Parcourir les 'ncomplet' de SS2 et vérifier dans SS1
for index, nom in enumerate(ncomplet_ss2):
    match = comparer_noms(nom, ncomplet_ss1)
    ss2.at[index, 'etat'] = 'admis' if not match else 'entrant'

# Enregistrer les résultats dans de nouveaux fichiers Excel
ss1.to_excel('SS1_etat.xlsx', index=False)
ss2.to_excel('SS2_etat.xlsx', index=False)
