import pandas as pd
import matplotlib.pyplot as plt


# Charger les 3 feuilles du fichier Excel dans 3 DataFrames différents
df1 = pd.read_excel("Statapp/Bases/parc_vp_region_2024_1.xlsx", sheet_name="Parc Régionaux", engine="openpyxl", header=3)
df2 = pd.read_excel("Statapp/Bases/parc_vp_region_2024_1.xlsx", sheet_name="Distance", engine="openpyxl", header=4)
df3 = pd.read_excel("Statapp/Bases/parc_vp_region_2024_1.xlsx", sheet_name="Parcours annuel moyen", engine="openpyxl", header=4)

# Propager les valeurs non-NaN dans les colonnes VP, Unnamed: 1 et Unnamed: 2
df2['VP'] = df2['VP'].ffill()
df2['Unnamed: 1'] = df2['Unnamed: 1'].ffill()
df2['Unnamed: 2'] = df2['Unnamed: 2'].ffill()
df3['VP'] = df3['VP'].ffill()
df3['Unnamed: 1'] = df3['Unnamed: 1'].ffill()
df3['Unnamed: 2'] = df3['Unnamed: 2'].ffill()


df1.to_csv("Statapp/Bases/parc_vp_region_2024_1_PR.xlsx", index=False)
df2.to_csv("Statapp/Bases/parc_vp_region_2024_1_Dis.xlsx", index=False)
df3.to_csv("Statapp/Bases/parc_vp_region_2024_1_PAM.xlsx", index=False)


'''# Afficher les premières lignes de chaque DataFrame pour vérifier
print("DF1:")
print(df1.head())


print("\nDF2:")
print(df2.head(15))


print("\nDF3:")
print(df3.head(15))
'''

# Extraire les lignes de 0 à 29
lignes_to_copy = df2.iloc[0:30]

# Insérer ces lignes après la ligne 29, donc dans les lignes 30 à 59
df_with_copied_rows = pd.concat([df2.iloc[:30], lignes_to_copy, df2.iloc[30:]]).reset_index(drop=True)

# Extraire les lignes de 30 à 59
df_lines_30_59 = df_with_copied_rows.iloc[30:60]

# Remplacer "Total" par "Moyennes régionales" dans la colonne "VP" (colonne 0) pour les lignes 30 à 59
df_lines_30_59.loc[:, 'VP'] = df_lines_30_59['VP'].replace('Total', 'Moyennes régionales')

# Diviser les valeurs des colonnes à partir de la 4ème colonne (incluse) par 20
df_lines_30_59 = df_lines_30_59.copy()
df_lines_30_59.loc[:, df_lines_30_59.columns[3:]] = df_lines_30_59.iloc[:, 3:] / 20



# Réinsérer ces lignes modifiées dans le DataFrame
df_with_modifications = pd.concat([df_with_copied_rows.iloc[:30], df_lines_30_59, df_with_copied_rows.iloc[60:]]).reset_index(drop=True)

# Afficher les premières lignes pour vérifier
#print(df_with_modifications.head(60))  # Affiche les 60 premières lignes pour vérifier

#df_with_modifications.to_csv("Statapp/Bases/parc_vp_region_2024_1_Dis_V2.xlsx", index=False)




'''
Tracer les distances parcourues par an selon les années selon les carburants. 
D'abord en total puis séparés selon professionels ou particuliers

'''
# Charger les données dans un DataFrame
df_distance_tot = df_with_modifications.iloc[:10].reset_index(drop=True)

# Renommer les colonnes pour éviter les "Unnamed" (à changer un peu parce que c'est pas sur que ce soit le bon nom)
df_distance_tot.rename(columns={"Unnamed: 1": "Statut", "Unnamed: 2": "Carburant"}, inplace=True)

# Filtrer pour exclure la ligne "Total" dans la colonne "Carburant"
df_distance_tot = df_distance_tot[df_distance_tot["Carburant"] != "Total"]

# Sélectionner les années (colonnes à partir de la 4ᵉ colonne)
years = df_distance_tot.columns[3:]

import matplotlib.pyplot as plt

# Crée un simple graphique pour tester
plt.plot([0, 1, 2], [0, 1, 4])
plt.title('Test Graphique')
plt.xlabel('X')
plt.ylabel('Y')

plt.show()
'''
# Tracer les courbes
plt.figure(figsize=(12, 6))

for _, row in df_distance_tot.iterrows():
    plt.plot(years, row[3:].values, label=row["Carburant"])  # Tracer chaque ligne

# Personnalisation du graphique
plt.xlabel("Année")
plt.ylabel("Distance parcourue")
plt.title("Distance parcourue par type de carburant selon les années (particuliers et professionels)")
plt.legend(loc="upper left", bbox_to_anchor=(1, 1))  # Mettre la légende en dehors du graphe

# Afficher le graphique
plt.xticks(rotation=45)  # Rotation des années pour une meilleure lisibilité
plt.grid(True)
plt.show()
'''