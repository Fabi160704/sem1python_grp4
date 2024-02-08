# Importieren der notwendigen Bibliotheken
import matplotlib.pyplot as plt  # Importiert die Matplotlib-Bibliothek für Diagramme
import pandas as pd

# Globale Konstanten für Debugging und Grafikmodus
DEBUG_INFO = False  # Schaltet Debug-Informationen ein oder aus

#Lesen der Exceldatei
def xfile_read(filename):
    # Lesen der Daten aus der Excel-Datei in ein DataFrame.
    df = pd.read_excel(filename, dtype=str)

    # Die Spaltennamen (Kopfzeile) in eine Liste umwandeln und in header_a speichern.
    header_a = df.columns.tolist()

    # Die Datenzeilen in eine Liste von Listen umwandeln und in alldata speichern.
    alldata = df.values.tolist()

    # Die extrahierte Kopfzeile und die Datenzeilen zurückgeben.
    return header_a, alldata

#Ergebenisse in eine Exceldatei schreiben
def xfile_write(data):
    filename = "100_Pivot_Output.xlsx"
    
    # Daten in ein DataFrame umwandeln
    df = pd.DataFrame(data)
    
    # Excel-Datei erstellen oder überschreiben und DataFrame speichern
    df.to_excel(filename, index=False)

    print(f"Daten wurden in die Datei {filename} geschrieben.")

# Diese Funktion erstellt ein Wörterbuch aus der übergebenen Kopfzeilen-Liste.
def create_header_dict(header_a):
    header_dict = {}  # Initialisiere ein leeres Wörterbuch

    # Durchlaufe alle Elemente in der Kopfzeilen-Liste
    for i in range(len(header_a)):
        # Füge jedes Element der Liste als Schlüssel in das Wörterbuch ein.
        # Der Wert ist der Index des Elements in der Liste.
        header_dict[header_a[i]] = i

    # Gib das fertige Wörterbuch zurück
    return header_dict

# Funktion zum Berechnen des Durchschnitts basierend auf einem bestimmten Spaltenindex
def calc_mean_by_index(alldata, search_term='Liefermenge', header = 'Liefermenge'):
    # Bestimme die Anzahl der Zeilen in den Daten
    num_rows = len(alldata)

    # Bestimme den Spaltenindex für den gesuchten Begriff mithilfe des zuvor erstellten header_dict
    index = header[search_term]

    total = 0
    # Summiere alle Werte in der ausgewählten Spalte
    for row in alldata:
    # Konvertiere den Wert in der Spalte mit dem Index "index" in einen Dezimalwert (float) und addiere ihn zu "total".
    # Dabei wird die Gesamtsumme der Werte berechnet.
        total += float(row[index])

    # Berechne den Durchschnitt, wenn Daten vorhanden sind
    mean = total / num_rows if num_rows > 0 else 0

    # Gib den berechneten Durchschnitt zurück
    return mean

# Funktion zum Konvertieren von deutschen Float-Zahlen ins Englische Format
def german_to_english_float(germfloat_string):
    germfloat_string = germfloat_string.replace(",", ".")
    return germfloat_string

# Funktion zum Berechnen des gewichteten Durchschnitts basierend auf einem bestimmten Spaltenindex
def calc_weighted_mean_by_index(min_value, alldata, search_term='Liefermenge', header = 'Liefermenge'):
    # Initialisiere die gesamte gewichtete Summe und die Anzahl der Werte über dem Mindestwert
    total_weighted_sum = 0
    counter = 0

    # Bestimme den Spaltenindex für den gesuchten Begriff mithilfe des zuvor erstellten header_dict
    index = header[search_term]

    # Iteriere über alle Zeilen in den Daten
    for row in alldata:
        # Extrahiere den aktuellen Wert in der ausgewählten Spalte und konvertiere ihn in Float
        current_value = float(row[index])

        # Überprüfe, ob der aktuelle Wert größer als das Mindestwert ist
        if current_value > min_value:
            # Addiere den aktuellen Wert zur gewichteten Summe
            total_weighted_sum += current_value
            # Inkrementiere die Anzahl der Werte über dem Mindestwert
            counter += 1

    # Berechne den gewichteten Durchschnitt, wenn die Anzahl der Werte über dem Mindestwert nicht null ist
    if counter != 0:
        weighted_mean = total_weighted_sum / counter
    elif counter == 0:
        weighted_mean = 0

    # Gib den berechneten gewichteten Durchschnitt zurück
    return weighted_mean

#Funktion zur Berechnung der Summe von Werten nach Gruppe und Parameter
def weighted_sum(alldata, parameter='101', search_term='Produktgruppe', header = 'Produktgruppe'):
    # Initialisiere die Summe mit 0.0
    summe = 0.0
    index = header[search_term]
    # Durchlaufe jede Zeile in der Datenliste
    for row in alldata:
        # Prüfe, ob der Wert im angegebenen Suchfeld (`search_term`) dem `parameter` entspricht
        if row[index] == parameter:
            # Konvertiere den Wert im Feld 'Wert' zu einer Fließkommazahl und addiere ihn zur Summe
            wert = float(row[header['Wert']])
            summe += wert

    # Gib die berechnete Summe zurück
    return summe

# Funktion zum Zeichnen eines Matplotlib-Balkendiagramms basierend auf den Daten
def draw_graph(alldata, search_term='Liefermenge'):
    index = header_dict[search_term]  # Der Spaltenindex für den gesuchten Begriff
    x_values = range(1, len(alldata) + 1)  # X-Werte für die Datenpunkte
    y_values = [float(row[index]) for row in alldata]  # Y-Werte für die Datenpunkte
    
    plt.grid(True)
    plt.bar(x_values, y_values)  # Erstellt ein Matplotlib-Balkendiagramm
    plt.xlabel('Datennummer')  # Setzt das Label für die X-Achse
    plt.ylabel(search_term)  # Setzt das Label für die Y-Achse
    plt.title(f'{search_term} Verteilung')  # Setzt den Titel des Diagramms
    plt.show()  # Zeigt das Diagramm an

# __main__
filename = "100_Pivot_Grunddaten.xlsx"
header_a, alldata = xfile_read(filename) #Excel-Datei einlesen
header_dict = create_header_dict(header_a)  # Erstellen eines Dictionarys für die Kopfzeile

if DEBUG_INFO:
    print("--- Start debug infos ---")
    print("Kopfzeile (header_a): ", header_a)
    print("... header_dict: ", header_dict)
    print("Datenzeilen (alldata): ", alldata)
    print("--- End debug infos ---")

while True:
    # Nutzerabfrage welche Aktion
    while True:
        try:
            search_nr = int(input("""Welchen der folgenden Terme wollen Sie untersuchen?
            1: (gew.) Durchschnitt von Bestellmenge
            2: (gew.) Durchschnitt von Liefermenge 
            3: (gew.) Durchschnitt von Wert 
            4: Summe der Werte nach Kategorie
            9: Exit 
            >>> """))
            break
        except ValueError:
            print("Bitte korrekte Zahl eingeben!")
            print("-----------------------------\n")

### Programm beenden ###
    if search_nr == 9: 
        break
    
    
### Durchschnitt ###
    elif 1 <= search_nr <= 3:
        search_terms = ['Bestellmenge', 'Liefermenge', 'Wert']
        # search_term das Wort zuweisen, welches in der Liste den selben Wert hat, wie Nummer bei der Eingabe
        search_term = search_terms[search_nr - 1]
        
        # Durchschnitt berechnen
        mean = calc_mean_by_index(alldata, search_term, header_dict)
        print(f"Durchschnittliche {search_term}: {mean:.2f}")

        # Gewichteter Durchschnittwert abfragen, Eingabe (wenn nötig) ins englische Format umwandeln
        while True:
            try:
                min_value = float(german_to_english_float(input("Geben Sie ein Minimum zur Berechnung des gew. MW ein: ")))
                break
            except ValueError:
                # Fehlerbehandlung, falls keine gültige Zahl eingegeben wird
                print("Bitte eine Zahl eingeben!")
                print("-------------------------\n")
        # Wenn die Eingabe kleiner gleich Null ist, Standardwert setzen
        if min_value <= 0:
            print("Das Minimum muss größer als 0 sein. Setze Minimum auf 200!")
            min_value = 200

        # Gewichteten Durchschnitt berechnen
        weighted_mean = calc_weighted_mean_by_index(min_value, alldata, search_term, header_dict)
        print(f"Gewichtete durchschnittliche {search_term} für min: {min_value:.2f} = {weighted_mean:.2f}")

        #Ergebnisse werden an Funktion übergeben, um in eine Excel zu schreiben
        data_out = [["Suchkriterium", "Durchschnitt", "Gewichteter Durchschnitt"], [search_term, mean, weighted_mean]]
        xfile_write(data_out)

        # Abfrage, ob ein Graph gezeichnet werden soll
        while True:
            # Benutzer wird aufgefordert, 'y' für Ja oder 'n' für Nein einzugeben
            user_input = input("Zugehörigen Graph zeichnen (y/n)? ").lower()

            if user_input == "y":
                # Meldung, dass das Zeichnen des Graphen beginnt, und Aufruf der Funktion
                print(f"Beginn der Erstellung eines Matplot-Diagramms ...")
                draw_graph(alldata, search_term)
                break
            elif user_input == 'n':
                # Schleife beenden, da der Benutzer keine Zeichnung wünscht
                break
            else:
                # Benutzer hat ungültige Eingabe gemacht, daher wird eine Meldung ausgegeben; Schleife wird fortgesetzt, um eine erneute Eingabe zu ermöglichen
                print("Ungültige Eingabe. Bitte geben Sie 'y' oder 'n' ein.")
                
             
### Summe der Werte ###           
    elif search_nr == 4:
        while True:
            try:
                # Eine Liste von Suchkriterien für den Benutzer anzeigen
                search_terms = ['Produktgruppe', 'Artikel', 'Bestellnummer', 'Kunde']
                category = int(input("""Welchen der folgenden Kategorien wollen Sie untersuchen?
                1: Produktgruppe
                2: Artikel
                3: Bestellnummer
                4: Kunde
                >>> """))
                
                # Überprüfen, ob die ausgewählte Kategorie gültig ist (zwischen 1 und 4)
                if 1 <= category <= 4:
                    search_term = search_terms[category - 1]
                    break
                else:
                    print("Bitte korrekte Zahl eingeben!")
                    print("-----------------------------\n")
            except ValueError:
                # Fehler behandeln, falls keine gültige Zahl eingegeben wird
                print("Bitte korrekte Zahl eingeben!")
                print("-----------------------------\n")

        value = input("Geben Sie Ihre Suchkriterium ein\n>>>")
        
        # Berechnung der gewichteten Summe und Anzeige des Ergebnisses
        sum = weighted_sum(alldata, value, search_term, header_dict)
        print(f"Die Summe der Werte in {search_term} {value} beträgt: {sum}")
        print("-----------------------------\n")
        
        # Ergebnisse werden an Funktion übergeben, um in eine Excel-Tabelle zu schreiben
        data_out = [["Suchkriterium", "Suchkriterium", "Summe der Werte"], [search_term, value, sum]]
        xfile_write(data_out)


### Wenn falsche zahl eingegeb wurde ###
    else:
        print("Bitte korrekte Zahl eingeben!")
        print("-----------------------------\n")