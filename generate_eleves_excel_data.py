import xlsxwriter

# Create a workbook and add a worksheet.
workbook = xlsxwriter.Workbook('eleve_data.xlsx')
worksheet = workbook.add_worksheet()

# Some data we want to write to the worksheet.
Prenom=['Marie', 'Sara', 'Allen', 'Marc', 'Virginia', 'Anthony', 'Hilda', 'Warren', 'Todd', 'Terry', 'Leila', 'Rosa', 'Nathaniel', 'Adam', 'Pauline', 'Darrell', 'Elva', 'Landon', 'Paul', 'Alvin', 'Lena', 'Caroline', 'Cordelia', 'Jesus', 'Hester', 'Max', 'Mable', 'Callie', 'Lloyd', 'Chester', 'Samuel', 'Leroy', 'Melvin', 'Evelyn', 'Mabel', 'Lela', 'Elijah', 'Craig', 'Bruce', 'Elmer', 'Gertrude', 'Laura', 'Delia', 'Stanley', 'Esther', 'Ellen', 'Travis', 'Isaac', 'Lou', 'Margaret', 'Nina', 'Owen', 'Lewis', 'Sally', 'Cecelia', 'George', 'Cecilia', 'Glen', 'Jason', 'Birdie', 'Seth', 'Dustin', 'Clayton', 'Agnes', 'Bessie', 'Lola', 'Martin', 'Myrtie', 'Milton', 'Bess']
Nom=['Oliver', 'Ferrante', 'Baronti', 'Palmer', 'Savage', 'Fabbrini', 'Meucci', 'Black', 'Metcalfe', 'Porter', 'Arrighi', 'Carpenter', 'Tomlinson', 'Ridolfi', 'Cianferoni', 'Rolland', 'Paoli', 'Braun', 'Benoît', 'Mazzini', 'Baldwin', 'Giannini', 'Naylor', 'Carli', 'King', 'Cecchi', 'Dumas', 'van der Veen', 'Petrucci', 'Robson', 'Giannetti', 'Dondoli', 'van Beek', 'Uchida', 'Markus', 'Poirier', 'Hammond', 'Powell', 'Parsons', 'Alfani', 'Thomas', 'Wilson', 'Blanco', 'Ranfagni', 'Griffith', 'Otsuka', 'Herrera', 'Garofalo', 'Talbot', 'Warner', 'Rastrelli', 'Fanfani', 'Harris', 'Elliott', 'Lucchesi', 'Reese', 'Ciolli', 'Simoncini', 'Schmitt', 'Harrison', 'Frati', 'Buti', 'Fabbrucci', 'Carrasco', 'Morton', 'Mustafa', 'Del Lungo', 'Patterson', 'Rinaldi', 'Schmidt']
Adresse=['332 Zigme Terrace', '436 Olwa Boulevard', '1932 Jomaw Key', '776 Rovu Grove', '487 Kabosi Lane', '1544 Kupgos Court', '1797 Cubzun Mill', '1951 Hicwis Point', '1865 Cotje Ridge', '109 Afofa Mill', '1781 Jihe Avenue', '793 Dizo Street', '1823 Ifgeh Avenue', '1517 Cijo Junction', '1292 Vusar Heights', '753 Meru Plaza', '680 Ruwim Plaza', '1608 Nuzzij Circle', '51 Wezrug Drive', '1129 Jotvu Parkway', '504 Vihme Place', '1541 Kuhub Court', '1510 Piwok Avenue', '1722 Sovot Park', '949 Zoziw Pass', '1181 Duco Extension', '1712 Luoz Drive', '70 Kotij Heights', '152 Suvwo Way', '916 Kuvi Place', '633 Geklaw Junction', '1684 Abjel Path', '441 Liplal Grove', '503 Eviwo Heights', '1700 Ipegi River', '276 Waad Glen', '472 Odisu Lane', '354 Onecal Path', '504 Vikkaj Turnpike', '1872 Lizno Path', '621 Nine Boulevard', '636 Abzug Extension', '798 Masup River', '843 Toeke Pike', '1467 Bawav Path', '41 Bijon Road', '340 Vebaz Loop', '1469 Ozicu Highway', '274 Luik Highway', '534 Jouf Center', '1782 Vuhuh Lane', '1315 Dorzi Place', '431 Evaod Heights', '699 Ebhe River', '568 Icema Street', '1903 Mijum Avenue', '1772 Waswo Street', '868 Deshiw Mill', '493 Zokuv Parkway', '1925 Hidtuz Manor', '1956 Ezwir Park', '508 Vubrec Extension', '1068 Kodap Turnpike', '313 Laplar Loop', '1718 Hospev Court', '1265 Noku Terrace', '583 Tede Trail', '1959 Wabo Parkway', '529 Kimfi Pike', '1137 Tumav Lane']
Age=[19, 15, 19, 14, 15, 15, 14, 17, 18, 20, 17, 16, 14, 17, 17, 17, 20, 15, 16, 13, 20, 18, 16, 18, 17, 13, 17, 16, 16, 13, 17, 20, 14, 13, 15, 18, 14, 15, 16, 13, 14, 16, 17, 18, 20, 14, 15, 18, 15, 14, 18, 17, 18, 20, 15, 13, 19, 15, 16, 14, 14, 15, 14, 18, 17, 19, 15, 15, 17, 16]
Moyenne=[1.2, 19.8, 4.3, 19.8, 16.6, 2.9, 1.4, 6.7, 5.1, 8.2, 7.5, 17.6, 19.6, 5.3, 0.1, 11.1, 18, 19, 5.6, 9, 0.8, 7, 16, 12, 0.5, 2.5, 1.3, 0.5, 4, 18.3, 8.9, 11, 2.9, 0.7, 0.7, 10.6, 8.5, 6.9, 4.6, 11.3, 4.1, 6.4, 17.9, 4, 5, 12.8, 9.6, 16.7, 3.1, 16.8, 13, 3, 15.1, 19.5, 10.4, 17.8, 6.1, 13.3, 17.6, 10.5, 1.5, 1.6, 16.2, 19.5, 2.4, 7.1, 12.4, 9.2, 19.9, 8.1]
Region=['Saint - Louis', 'Fatick', 'Tambacounda', 'Kaffrine', 'Sedhiou', 'Kolda', 'Sedhiou', 'Kaffrine', 'Sedhiou', 'Diourbel', 'Dakar', 'Thies', 'Kaolack', 'Matam', 'Kaolack', 'Sedhiou', 'Kolda', 'Saint - Louis', 'Kaolack', 'Sedhiou', 'Kolda', 'Thies', 'Tambacounda', 'Diourbel', 'Saint - Louis', 'Saint - Louis', 'Kaffrine', 'Diourbel', 'Matam', 'Kedougou', 'Ziguinchor', 'Diourbel', 'Fatick', 'Kolda', 'Kolda', 'Kolda', 'Thies', 'Saint - Louis', 'Kaolack', 'Kaolack', 'Ziguinchor', 'Tambacounda', 'Louga', 'Thies', 'Saint - Louis', 'Sedhiou', 'Ziguinchor', 'Diourbel', 'Fatick', 'Sedhiou', 'Thies', 'Kedougou', 'Diourbel', 'Matam', 'Kolda', 'Fatick', 'Tambacounda', 'Kaffrine', 'Matam', 'Sedhiou', 'Kaolack', 'Kedougou', 'Ziguinchor', 'Kaffrine', 'Diourbel', 'Thies', 'Dakar', 'Kedougou', 'Thies', 'Matam']
Sexe=['F','F','M','M','F','M','F','M','M','M','F','F','M','M','F','M','F','M','M','M','F','F','F','M','M','M','M','F','M','M','M','M','F','F','F','F','F','M','M','M','M','F','F','M','F','F','M','M','F','F','F','M','M','F','F','M','F','M','M','F','F','M','M','F','F','F','F','F','M','F']
Specialite=['Francais', 'Francais', 'Francais', 'Francais', 'Histoire', 'Geographie', 'Francais', 'Chimie', 'Geographie', 'Francais', 'Francais', 'Anglais', 'Geographie', 'Geographie', 'Histoire', 'Physique', 'Geographie', 'Chimie', 'Chimie', 'Physique', 'Mathematique', 'Mathematique', 'Chimie', 'Francais', 'Mathematique', 'Geographie', 'Physique', 'Physique', 'Mathematique', 'Mathematique', 'Chimie', 'Francais', 'Anglais', 'Anglais', 'Mathematique', 'Chimie', 'Anglais', 'Chimie', 'Mathematique', 'Mathematique', 'Francais', 'Anglais', 'Physique', 'Physique', 'Anglais', 'Physique', 'Mathematique', 'Francais', 'Mathematique', 'Geographie', 'Francais', 'Physique', 'Francais', 'Geographie', 'Physique', 'Histoire', 'Geographie', 'Histoire', 'Anglais', 'Chimie', 'Geographie', 'Physique', 'Anglais', 'Chimie', 'Histoire', 'Geographie', 'Mathematique', 'Chimie', 'Geographie', 'Chimie']

tup = ()
for i in range(70):
    l = [Prenom[i], Nom[i], Age[i], Adresse[i], Moyenne[i], Region[i], Specialite[i], Sexe[i]]
    tup = tup + (l,)

# Start from the first cell. Rows and columns are zero indexed.
row = 0
col = 0

# Iterate over the data and write it out row by row.
for prenom, nom, age, adresse, moyenne, region, specialite, sexe in (tup):
    worksheet.write(row, col,     prenom)
    worksheet.write(row, col + 1, nom)
    worksheet.write(row, col + 2, age)
    worksheet.write(row, col + 3, adresse)
    worksheet.write(row, col + 4, moyenne)
    worksheet.write(row, col + 5, region)
    worksheet.write(row, col + 6, specialite)
    worksheet.write(row, col + 7, sexe)
    row += 1


workbook.close()