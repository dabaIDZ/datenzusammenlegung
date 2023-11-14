# This is a sample Python script.

# Press Umschalt+F10 to execute it or replace it with your code.
# Press Double Shift to search everywhere for classes, files, tool windows, actions, and settings.
import pandas as pd
import zipfile
import os
import pyreadstat

SteffenTest = True
pfad = r"C:\Users\steff\Downloads\tmp071101"
beratungsstellen = ["a", "b", "c"]
tabellen_univariat = ["AGG-Relevanz_häufigkeit"]
tabellen_bivariat = ["Mehrfachnennung nach AGG-Relevanz_kreuztabelle"]
spalten_loeschen = ["ignorieren - AGG-RelevanzN", "keine Angabe - AGG-RelevanzN"]
spalten_umbenennen = {'AGG-relevant': 'A1', 'B': 'Neu_B'}

if SteffenTest == True:
    pfad = r"C:\Users\steff\Downloads\tmp071101"
    beratungsstellen = ["a", "b", "c"]
    tabellen_univariat = ["AGG-Relevanz_häufigkeit"]
    tabellen_bivariat = ["Mehrfachnennung nach AGG-Relevanz_kreuztabelle"]
    spalten_loeschen = ["ignorieren - AGG-RelevanzN", "keine Angabe - AGG-RelevanzN"]
    spalten_umbenennen = {'AGG-relevant': 'A1', 'B': 'Neu_B'}

for beratungsstelle in beratungsstellen:
    zip_file_path = pfad + "\\" + beratungsstelle + "\\" + "An IDZ senden.zip"
    zip_extraktion_path = pfad + "\\" + beratungsstelle
    with zipfile.ZipFile(zip_file_path, "r") as zip_file:
        # Extrahiere alle Dateien in den aktuellen Ordner
        zip_file.extractall(zip_extraktion_path)

def beratungsstellen_zusammenfuegen(df, df_zusammen, erster_durchgang_tabelle):
    # Füge eine neue Spalte "Beratungsstelle" hinzu und trage den Wert ein
    df['Beratungsstelle'] = beratungsstelle

    # Überprüfe, ob es sich um den ersten Durchgang handelt
    if erster_durchgang_tabelle:
        df_zusammen = df.copy()  # Kopiere das DataFrame
        erster_durchgang_tabelle = False
    else:
        # Füge die Zeile zum vorhandenen DataFrame df_zusammen hinzu
        df_zusammen = pd.concat([df_zusammen, df], ignore_index=True)

    return df_zusammen, erster_durchgang_tabelle

def tabellen_zusammenfuegen(df_zusammen, df_gesamt, erster_durchgang):
    # Überprüfe, ob es sich um den ersten Durchgang handelt
    if erster_durchgang:
        df_gesamt = df_zusammen.copy()  # Kopiere das DataFrame
        erster_durchgang = False
    else:
        # Füge die Zeile zum vorhandenen DataFrame df_zusammen hinzu
        df_gesamt = pd.merge(df_gesamt, df_zusammen, how='outer', on='Beratungsstelle')

    return df_gesamt, erster_durchgang

erster_durchgang = True
df_gesamt = pd.DataFrame()
for tabelle in tabellen_univariat:
    erster_durchgang_tabelle = True
    df_zusammen = pd.DataFrame()
    # Gehe durch die Liste der Beratungsstellen
    for beratungsstelle in beratungsstellen:
        # Öffne das Excel-Dokument
        dateipfad = pfad + "\\" + beratungsstelle + "\\" + tabelle + ".xlsx"
        if os.path.exists(dateipfad):
            df = pd.read_excel(pfad + "\\" + beratungsstelle + "\\" + tabelle + ".xlsx")

            # Definiere die Spaltennamen
            spalten_zu_loeschen = ["nicht genannt", "keine Angabe", "trifft nicht zu"]
            # Lösche die Spalten
            df = df.drop(spalten_zu_loeschen, axis=1)

            # Transponiere den Datensatz
            df = df.T
            df = df.reset_index(drop=True)
            df = df.drop(df.index[0])
            df = df.reset_index(drop=True)
            df.columns = df.iloc[0]
            df = df.drop(df.index[0])
            df = df.reset_index(drop=True)
            df_zusammen, erster_durchgang_tabelle = beratungsstellen_zusammenfuegen(df, df_zusammen, erster_durchgang_tabelle)
        else:
            print(f"Von {beratungsstelle} liegt {tabelle} nicht vor.")
    df_gesamt, erster_durchgang = tabellen_zusammenfuegen(df_zusammen, df_gesamt, erster_durchgang)

for tabelle in tabellen_bivariat:
    # Gehe durch die Liste der Beratungsstellen
    erster_durchgang_tabelle = True
    df_zusammen = pd.DataFrame()
    for beratungsstelle in beratungsstellen:
        # Öffne das Excel-Dokument
        dateipfad = pfad + "\\" + beratungsstelle + "\\" + tabelle + ".xlsx"
        if os.path.exists(dateipfad):
            df = pd.read_excel(pfad + "\\" + beratungsstelle + "\\" + tabelle + ".xlsx")

            # DataFrame in Zeichenketten umwandeln und dann schmelzen
            df_str = df.astype(str)
            df_str = df_str.set_index('Unnamed: 0')
            #df_transformed = df_str.melt(var_name='original_column', value_name='value').set_index('original_column')['value'].reset_index()
            # Spalte 'original_column' mit ursprünglichen Zeilennummern hinzufügen
            df_transformed = df_str.reset_index().melt(id_vars=df_str.index.name, var_name='original_column', value_name='value')

            # Optional: Spalte 'index' in 'original_column' umbenennen
            #df_transformed = df_transformed.rename(columns={'index': 'originalzeilennummer'})

            df_transformed["titel"] = df_transformed["Unnamed: 0"] + df_transformed["original_column"]
            df_transformed = df_transformed.drop(['Unnamed: 0', 'original_column'], axis=1, errors='ignore')
            df = df_transformed.T
            df = df.reset_index(drop=True)
            df.columns = df.iloc[1]
            df = df.drop(df.index[1])
            df = df.reset_index(drop=True)

            df_zusammen, erster_durchgang_tabelle = beratungsstellen_zusammenfuegen(df, df_zusammen, erster_durchgang_tabelle)
        else:
            print(f"Von {beratungsstelle} liegt {tabelle} nicht vor.")

    df_gesamt, erster_durchgang = tabellen_zusammenfuegen(df_zusammen, df_gesamt, erster_durchgang)

# Überprüfe, ob die Spalten im DataFrame vorhanden sind, bevor du sie löschst
spalten_zum_loeschen = [col for col in spalten_loeschen if col in df_gesamt.columns]
# Lösche die gefundenen Spalten
df_gesamt = df_gesamt.drop(columns=spalten_zum_loeschen, errors='ignore')

# Überprüfe, ob die alten Spaltennamen im DataFrame vorhanden sind, bevor du sie umbenennst
spalten_zum_umbenennen = [col for col in spalten_umbenennen.keys() if col in df_gesamt.columns]
# Ändere die Spaltennamen
df_gesamt = df_gesamt.rename(columns=spalten_umbenennen, errors='ignore')

print(df_gesamt)

df_gesamt.to_excel(pfad + "\\" + "ausgabe.xlsx")
print("Ende erreicht")
# Exportiere den DataFrame als SPSS-Datei
#pyreadstat.write_sav(df_gesamt, pfad + "\\" + "ausgabe.sav")