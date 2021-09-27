# Input: Reporter-Abfrage "Stimmberechtigte Mitgliederversammlung"
#        in Zwischenablage, dann dieses Skript starten

import csv
import io
import win32clipboard
import datetime

heute = datetime.datetime.strftime(datetime.datetime.today(), "%Y-%m-%d_%H%M%S")

print("Bitte Reporter-Abfrage 'Stimmberechtigte Mitgliederversammlung'")
print("durchführen und Daten in Zwischenablage ablegen.")
input("Bitte ENTER drücken, wenn dies geschehen ist!")

win32clipboard.OpenClipboard()
data = win32clipboard.GetClipboardData()
win32clipboard.CloseClipboard()

if not data.startswith("Person (Nr)\t"):
    print("Fehler: Unerwarteter Inhalt der Zwischenablage!")
    exit()

eltern_nummern = set()
daten = []

# Erster Durchlauf: Schnupperer entfernen und Nummern aller Eltern sammeln
# Außerdem die Adressanrede durch den korrekten Text ersetzen und Zahlen in integer wandeln

with io.StringIO(data) as infile:
    for person in csv.DictReader(infile, delimiter="\t"):
        if "Schnupperer" in person["Bereich"]: # Schnupperer werden nicht berücksichtigt
            continue
        person["Alter am Stichtag"] = int(person["Alter am Stichtag"])
        person["Person (Nr)"] = int(person["Person (Nr)"])
        person["Nummer HZ"] = int(person["Nummer HZ"])

        if 0 < person["Alter am Stichtag"] < 18: # bei Minderjährigen: Elternteil ist stimmberechtigt
            eltern_nummern.add(person["Nummer HZ"])

        if person["Adressanrede (Nr) HZ"] == "1":
            person["Adressanrede (Nr) HZ"] = "Frau"
        elif person["Adressanrede (Nr) HZ"] == "2":
            person["Adressanrede (Nr) HZ"] = "Herrn"
        elif person["Adressanrede (Nr) HZ"] == "4":
            person["Adressanrede (Nr) HZ"] = "Familie"

        daten.append(person)




stimmberechtigte = []

for person in daten: # Fördermitglieder, die schon als Eltern oder Mitglieder erfasst sind, löschen    
    if person["Person (Nr)"] in eltern_nummern and person["Status"] == "Förderndes Mitglied":
        print(f"{person['Vorname']} {person['Name']} ({person['Person (Nr)']}) nicht berücksichtigt")
    else:
        sbp = {}
        sbp["Nummer"] = person["Person (Nr)"]
        sbp["Status (Mitglied)"] = person["Status"]
        if 0 < person["Alter am Stichtag"] < 18: # bei Minderjährigen: Elternteil ist stimmberechtigt 
            # (Alter 0 = nicht angegeben, nur bei Erwachsenen zu erwarten)
            sbp["Adressanrede"] = person["Adressanrede (Nr) HZ"]
            sbp["Titel"] = person["Titel HZ"]
            sbp["Vorname"] = person["Vorname HZ"]
            sbp["Name"] = person["Name HZ"]
            sbp["Adresse"] = person["Straße/Postfach HZ"]
            sbp["PLZ"] = person["PLZ HZ"]
            sbp["Ort"] = person["Ort HZ"]
            sbp["Mail"] = person["Serienbrief E-Mail Adresse HZ"]
            sbp["stimmberechtigt für"] = f'{person["Vorname"]} {person["Name"]} ({person["Alter am Stichtag"]})'
        else:
            sbp["Adressanrede"] = person["Adressanrede"]
            sbp["Titel"] = person["Titel"]
            sbp["Vorname"] = person["Vorname"]
            sbp["Name"] = person["Name"]
            sbp["Adresse"] = person["Straße/Postfach"]
            sbp["PLZ"] = person["PLZ"]
            sbp["Ort"] = person["Ort"]
            sbp["Mail"] = person["Serienbrief E-Mail Adresse"]
            sbp["stimmberechtigt für"] = "sich selbst"

    
    stimmberechtigte.append(sbp)

feldnamen = ["Nummer", "Status (Mitglied)", "Adressanrede", "Titel", "Vorname", "Name", "Adresse", "PLZ", "Ort", "Mail", "stimmberechtigt für"]   

with open(f"Stimmberechtigte_{heute}.csv", mode="w", newline="", encoding="cp1252") as outfile:
    output = csv.DictWriter(outfile, feldnamen, delimiter=";")
    output.writeheader()
    output.writerows(stimmberechtigte)

print(f"Fertig! Die Datei Stimmberechtigte_{heute}.xlsx wurde im aktuellen Ordner abgelegt.")
