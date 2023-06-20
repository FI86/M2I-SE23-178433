# coding:utf-8
import os
import pandas as p

chemin = os.path.dirname(__file__)

# Utilisons un dictionnaire pour remplir un DataFrame:
# Les clés dans notre dictionnaire servira de nom de colonne.
# De même, le valeurs devenir les lignes contenant les informations.
df = p.DataFrame({"Pays": ["France", "Canada", "Belgique"], 
                  "Capitale": ["Paris", "Ottawa", "Bruxelles"]})

# Nous pouvons utiliser le to_excel() pour écrire le contenu dans un fichier.
# df.to_excel(chemin + "/pays.xlsx")

# sheet name
# df.to_excel(chemin + "/pays.xlsx", sheet_name="Pays")

# index = False --> Ne creer pas de premiere colonne avec un numero d'index par defaut 
df.to_excel(chemin + "/pays.xlsx", sheet_name="Pays", index = False)

# Ecriture de plusieurs DataFrames dans un fichier Excel
income1 = p.DataFrame({"Noms": ["Stephen", "Camilla", "Tom"], "Salaire": [100000, 70000, 60000]}) 
income2 = p.DataFrame({"Noms": ["Pete", "April", "Marty"], "Salaire": [120000, 110000, 50000]})
income3 = p.DataFrame({"Noms": ["Victor", "Victoria", "Jennifer"], "Salaire": [75000, 90000, 40000]})

income_sheets = {"Groupe1": income1, "Groupe2": income2, "Groupe3": income3}
writer = p.ExcelWriter(chemin + "/income.xlsx", engine="xlsxwriter")

for sheet_name in income_sheets.keys(): 
    income_sheets[sheet_name].to_excel(writer, sheet_name=sheet_name, index=False)

writer.save()

# Previsualisation
students_grades = p.read_excel(chemin + "/mics.xlsx")
print(students_grades.head())

cols = [0, 1]
mics_result = p.read_excel(chemin + "/mics.xlsx", usecols=cols)
print(mics_result.head())

# Parcourir
xl = p.ExcelFile(chemin + "/mics.xlsx")
df1 = p.read_excel(chemin + "/mics.xlsx", sheet_name = xl.sheet_names[0], header = None, usecols=cols) # data

for j in df1:
    print(eval("df1[1][j]")) # Lecture de toute la colonne 1
    # Stop à la 6eme ligne pour ne pas perdre de temps à afficher
    # des milliers de ligne contenu dans le fichier.
    if int(j) > 6:
        break