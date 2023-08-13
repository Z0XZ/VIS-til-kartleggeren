import pandas as pd
import numpy as np
import xlrd

#Påkrevde kolonner ansatte
kolonner_ansatte = "Fornavn, Etternavn, Fødselsnummer, Epost (brukernavn)".split(", ")


#Lager en tom liste for ansatte
ansatt_tabell = {}
for kolonner in kolonner_ansatte:
    ansatt_tabell[kolonner] = [np.nan]

ansatt_tabell = pd.DataFrame(ansatt_tabell)

#ansatt_tabell = pd.read_excel("Ansatte_kartleggeren.xlsx", dtype=str) # Kommenter denne linjen ut om du ikke skal importere med ansatte.

#Påkrevde kolonner elever
kolonner = "Fornavn Etternavn Fødselsnummer Epost Trinn Klassenavn".split(" ")

#Leser inn Excel-filen
df = pd.read_excel("Elevoversikt_vis.xls")

#Fjerner rader som ikke har med klasser
df = df.dropna(subset=["Klasse"])

#Fjerner irrelevante kolonner
index = 0

for heading in df:
    if index > 2:
        del df[heading]
    index += 1

for kolonne in kolonner:
    if kolonne not in df:
        df[kolonne] = np.zeros(len(df.index))

trinn = "1. VGS, 2. VGS, 3. VGS, 7. trinn, 10. trinn".split(", ")

df = df.reset_index(drop=True)

index = 0
klasser = []

df["Trinn"] = df["Trinn"].astype(str)

for klassen in df["Klasse"]:
    if klassen[0] == "1":
        klasse = trinn[0]
    elif klassen[0] == "2":
        klasse = trinn[1]
    elif klassen[0] == "3" or klassen[0] == "4":
        klasse = trinn[2]
    elif klassen[-2] + klassen[-1] == "HT":
        klasse = trinn[0]
    elif klassen[0:3] == "GSM":
        klasse = trinn[3]
    elif klassen[0] + klassen[1] == "L1":
        klasse = trinn[0]
    elif klassen[0] + klassen[1] == "L2":
        klasse = trinn[1]
    
    df.at[index, 'Trinn'] = klasse
    #df["Trinn"][index] = klasse #Tydeligvis dårlig kode..
    index += 1

df = df.replace(0,np.nan)

#Sortere etter ulike kolonner
#df = df.sort_values("Klasse")

#Sortere datakolonnene i riktig rekkefølge
tabell = {}

for kolonne in kolonner:
    if kolonne == "Klassenavn":
        tabell[kolonne] = df["Klasse"].tolist()
    else:
        tabell[kolonne] = df[kolonne].tolist()

tabell = pd.DataFrame(tabell)

# Sjekke om arket inneholder noe nytt:
df = pd.read_excel("ElevExportKartleggeren.xlsx", sheet_name=['Lærere', 'Elever'])
df1 = df['Elever']

# Om ikke tabellen inneholder nye/oppdaterte elever lagres ikke filen på nytt.
if len(tabell) == len(df1):
    if df1.compare(tabell).empty:
        print("Tabellene er like!")
        print("Avbryter lagring :)")
    else:
        print("Tabellene er ulike, lagrer ny fil som: ElevExportKartleggeren.xlsx")
        #Eksporterer dataen til excel-arket
        writer = pd.ExcelWriter('ElevExportKartleggeren.xlsx')
        ansatt_tabell.to_excel(writer,'Lærere',index=False)
        tabell.to_excel(writer,'Elever',index=False)
        writer.save()
        print("Ny fil er lagret!")
else:
    print("Tabellene er ulike, lagrer ny fil som: ElevExportKartleggeren.xlsx")
    #Eksporterer dataen til excel-arket
    writer = pd.ExcelWriter('ElevExportKartleggeren.xlsx')
    ansatt_tabell.to_excel(writer,'Lærere',index=False)
    tabell.to_excel(writer,'Elever',index=False)
    writer.save()
    print("Ny fil er lagret!")
