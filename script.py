import pandas as pd
import os

eleves = pd.read_excel(os.getcwd()+'/eleve_data.xlsx')
eleves_moyenne = eleves[eleves.Moyenne >= 10]
eleves_moyenne.to_excel(os.getcwd()+'/eleve_moyenne_data.xlsx', index=False)

eleves_age = eleves[eleves.Age > 20 ]
eleves_age.to_excel(os.getcwd()+'/eleve_plus_20_data.xlsx', index=False)

