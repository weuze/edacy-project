import pandas as pd
import os

eleves = pd.read_excel(os.getcwd()+'/eleve_data.xlsx')

# Extraire et créer un fichier EXCEL contenant les liste des élèves ayant la moyenne
eleves_moyenne = eleves[eleves.Moyenne >= 10]
eleves_moyenne.to_excel(os.getcwd()+'/eleve_moyenne_data.xlsx', index=False)

# Extraire et créer un fichier EXCEL contenant les élèves ayant plus de 20 ans
eleves_age = eleves[eleves.Age > 20 ]
eleves_age.to_excel(os.getcwd()+'/eleve_plus_20_data.xlsx', index=False)

# Extraire et créer un fichier EXCEL contenant les statistiques globales de l’école
moyenne = eleves['Moyenne'].mean()
pcFille = (eleves[eleves.Sexe == 'F'].shape[0]/eleves.shape[0])*100
pcGarcon = (eleves[eleves.Sexe == 'M'].shape[0]/eleves.shape[0])*100
meuilleur_region = eleves.groupby('Region')['Moyenne'].mean().sort_values(ascending=False).index[0]

df = pd.DataFrame([[moyenne, pcFille, pcGarcon,meuilleur_region]])
df.columns = ['Moyenne Ecole', 'Percent. Fille', 'Percent. Garcon', 'Meuilleur Region']
df.to_excel(os.getcwd()+'/stat_ecole.xlsx', index=False)
