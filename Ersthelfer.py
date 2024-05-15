import pandas as pd
from bs4 import BeautifulSoup

path = 'G:\\Abteilungen\\Datenverarbeitung\\Programme_und_Features\\ErsteHilfe\\ErstHelfer.xlsx'
htmlPath = "G:\\Abteilungen\\Datenverarbeitung\\Programme_und_Features\\ErsteHilfe\\mainScreen.html"
#["AV","Auzubi","Betriebsratvorsitzender","Brandschutzhelfer","CNC-Dreherei","EK","Gear Saver","GF","Getriebemontage","Hausmeister","Immobilien","IT","Instandhaltung","LW","Materialwesen","Mechanische Fertigung","Produktion","PM","QS","RW","Service","Technik","VA","VI","Werkstatt","Sekreteriat","Personalwesen","Vertrieb","Qualitätswesen"]
divId_list =  ["AV","AZ","DM","EK","GF","GIG","IN","IT","L","LW","MF","P","PM","PW","QS","QW","RM","RW","S","T","V","VZ","Arzt"]


idcounter = 0
counter = 0

# Excel-Datei einlesen
df = pd.read_excel(path,sheet_name="Tabelle1")

# DataFrame anzeigen

# Öffnen Sie die HTML-Datei zum Schreiben

    # Schreiben Sie den HTML-Header
with open(htmlPath, 'r') as f:
    html_content = f.read()


# Erstellen eines BeautifulSoup-Objekts
soup = BeautifulSoup(html_content, "html.parser")

# Iteriere über die Div-Elemente und lösche den Inhalt
for div_id in divId_list:
    element = soup.find(id=div_id)
    if element:
        element.clear()

# Funktion zum Umwandeln von Umlauten in Umschreibungen
def umlaute_umwandeln(text):
    umlaute_dict = {'ä': 'ae', 'ö': 'oe', 'ü': 'ue', 'Ä': 'Ae', 'Ö': 'Oe', 'Ü': 'Ue', 'ß': 'ss'}
    for umlaut, umschreibung in umlaute_dict.items():
        text = text.replace(umlaut, umschreibung)
    return text

for  name in divId_list:
    div_id = name
    element = soup.find(id=div_id)
    h1_tag = soup.new_tag("h1")
    h1_tag.string = div_id
    if element:
        element.append(h1_tag)

# Iteriere über die Excel-Daten und füge sie zu den Div-Elementen hinzu
for idx, name in enumerate(df["Name"]):
    div_id = divId_list[df['Kenner'][idx]]
    element = soup.find(id=div_id)
    
    if element:
        if df["Ersthelfer"][idx] == 1:
            li_tag = soup.new_tag("li")
            li_tag.string = ""
            
            # Füge ein <span> Element innerhalb des <li> Elements hinzu
            span_tag = soup.new_tag("span")
            span_tag.string = umlaute_umwandeln(name) + " "  # Umlaute umwandeln
            li_tag.append(span_tag)
            element.append(li_tag)
        
        else:
            li_tag = soup.new_tag("li")
            li_tag.string = ""
            
            # Füge ein <span> Element innerhalb des <li> Elements hinzu
            span_tag = soup.new_tag("abbr")
            span_tag.string = umlaute_umwandeln(name) + " "  # Umlaute umwandeln
            li_tag.append(span_tag)
            element.append(li_tag)
        

        p_tag = soup.new_tag("p")
        p_tag.string = str(df["Nummer"][idx])
        element.append(p_tag)

# Speichern der geänderten HTML-Datei
with open(htmlPath, "w") as f:
    f.write(str(soup))

