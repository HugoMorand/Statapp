import pandas as pd

# Charger les 3 feuilles du fichier Excel dans 3 DataFrames différents
df1 = pd.read_excel("Statapp/Bases/parc_vp_region_2024_1.xlsx", sheet_name="Parc Régionaux", engine="openpyxl", header=3)
df2 = pd.read_excel("Statapp/Bases/parc_vp_region_2024_1.xlsx", sheet_name="Distance", engine="openpyxl", header=4)
df3 = pd.read_excel("Statapp/Bases/parc_vp_region_2024_1.xlsx", sheet_name="Parcours annuel moyen", engine="openpyxl", header=4)

# Propager les valeurs non-NaN dans les colonnes VP, Unnamed: 1 et Unnamed: 2
df2['VP'] = df2['VP'].fillna(method='ffill')
df2['Unnamed: 1'] = df2['Unnamed: 1'].fillna(method='ffill')
df2['Unnamed: 2'] = df2['Unnamed: 2'].fillna(method='ffill')

df3['VP'] = df3['VP'].fillna(method='ffill')
df3['Unnamed: 1'] = df3['Unnamed: 1'].fillna(method='ffill')
df3['Unnamed: 2'] = df3['Unnamed: 2'].fillna(method='ffill')

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
df_lines_30_59['VP'] = df_lines_30_59['VP'].replace('Total', 'Moyennes régionales')

# Diviser les valeurs des colonnes à partir de la 4ème colonne (incluse) par 20
df_lines_30_59.iloc[:, 3:] = df_lines_30_59.iloc[:, 3:] / 20

# Réinsérer ces lignes modifiées dans le DataFrame
df_with_modifications = pd.concat([df_with_copied_rows.iloc[:30], df_lines_30_59, df_with_copied_rows.iloc[60:]]).reset_index(drop=True)

# Afficher les premières lignes pour vérifier
print(df_with_modifications.head(60))  # Affiche les 60 premières lignes pour vérifier

df_with_modifications.to_csv("Statapp/Bases/parc_vp_region_2024_1_Dis_V2.xlsx", index=False)


