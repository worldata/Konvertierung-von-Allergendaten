import pandas as pd
import numpy as np
import os.path, time
from datetime import datetime

#Katalogkonvertierung:

# Entscheidung Defaultwert Bestelleinheit, Einzeleinheit
Bestelleinheit = "Stück"

# None = Keine Kopfzeile in Outputdatei; True = mit Kopfzeile in Outputdatei
Kopfzeile = None

def konvertieren(Dateiname):
    # Dateinamen erstellen
    Erstellungsdatum = datetime.now().strftime("%d_%m_%Y")
    Erstellungsdatum_Y_M_D = datetime.now().strftime("%Y%m%d")

    Input_Datei = Dateiname

    df = pd.read_excel(
        "C:\\Users\Maximilian.Rasch\\Desktop\\Projekt 400\\Input\\" + Input_Datei,
        sheet_name="Allergene-Zusatzstoffe",
    )

    # Weizen Allergene

    df = df.replace(np.nan, "", regex=True)

    df["Keine Allergene"] = df["Keine Allergene"].map({1: 0, "": 1})
    df["… Weizen [A1]"] = df["… Weizen [A1]"].map({1: "A1"})
    df["… Roggen [A2]"] = df["… Roggen [A2]"].map({1: "A2"})
    df["… Gerste [A3]"] = df["… Gerste [A3]"].map({1: "A3"})
    df["… Hafer [A4]"] = df["… Hafer [A4]"].map({1: "A4"})
    df["… Dinkel [A5]"] = df["… Dinkel [A5]"].map({1: "A5"})
    df["… Kamut [A6]"] = df["… Kamut [A6]"].map({1: "A6"})

    # Normale Allergene
    df["Krebstiere / Krebstiererzeugnisse [B]"] = df[
        "Krebstiere / Krebstiererzeugnisse [B]"
    ].map({1: 2, "": 0})
    df["Eier / Eierzeugnisse [C] "] = df["Eier / Eierzeugnisse [C] "].map({1: 2, "": 0})
    df["Fisch / Fischerzeugnisse [D] "] = df["Fisch / Fischerzeugnisse [D]"].map(
        {1: 2, "": 0}
    )
    df["Erdnüsse / Erdnusserzeugnisse [E] "] = df["Erdnüsse / Erdnusserzeugnisse [E]"].map(
        {1: 2, "": 0}
    )
    df["Soja / Sojaerzeugnisse [F]"] = df["Soja / Sojaerzeugnisse [F]"].map({1: 2, "": 0})
    df["Milch / Milcherzeugnisse einschl. Lactose [G]"] = df[
        "Milch / Milcherzeugnisse einschl. Lactose [G]"
    ].map({1: 2, "": 0})

    # Nüsse
    df["Schalenfrüchte (Nüsse) und Erzeugnisse [H]"] = df[
        "Schalenfrüchte (Nüsse) und Erzeugnisse [H]"
    ].map({1: 2, "": 0})
    df["… Mandel [H1]"] = df["… Mandel [H1]"].map({1: "H1"})
    df["… Haselnüsse  [H2]"] = df["… Haselnüsse  [H2]"].map({1: "H2"})
    df["… Walnüsse  [H3]"] = df["… Walnüsse  [H3]"].map({1: "H3"})
    df["… Cashewnüsse / Kaschunüsse  [H4]"] = df["… Cashewnüsse / Kaschunüsse  [H4]"].map(
        {1: "H4"}
    )
    df["… Pecannüsse [H5]"] = df["… Pecannüsse [H5]"].map({1: "H5"})
    df["… Paranüsse [H6]"] = df["… Paranüsse [H6]"].map({1: "H6"})
    df["… Pistazien  [H7]"] = df["… Pistazien  [H7]"].map({1: "H7"})
    df["… Macadamianüsse / Queenslandnüsse [H8]"] = df[
        "… Macadamianüsse / Queenslandnüsse [H8]"
    ].map({1: "H8"})

    # Normale Allergene
    df["Sellerie / Sellerieerzeugnisse [I]"] = df["Sellerie / Sellerieerzeugnisse [I]"].map(
        {1: 2, "": 0}
    )
    df["Senf / Senferzeugnisse [J]"] = df["Senf / Senferzeugnisse [J]"].map({1: 2, "": 0})
    df["Sesam / Sesamerzeugnisse [K]"] = df["Sesam / Sesamerzeugnisse [K]"].map(
        {1: 2, "": 0}
    )
    df["Schwefeldioxid und Sulfite [L]"] = df["Schwefeldioxid und Sulfite [L]"].map(
        {1: 2, "": 0}
    )
    df["Lupinen / Lupinenerzeugnisse [M]"] = df["Lupinen / Lupinenerzeugnisse [M]"].map(
        {1: 2, "": 0}
    )
    df["Weichtiere / Weichtiererzeugnisse [N]"] = df[
        "Weichtiere / Weichtiererzeugnisse [N]"
    ].map({1: 2, "": 0})

    # Zusatzstoffe
    df["Keine Zusatzstoffe"] = df["Keine Zusatzstoffe"].map({1: 0, "": 1})
    df["Antioxidationsmittel [1]"] = df["Antioxidationsmittel [1]"].map({1: "1"})
    df["mit Konservierungsstoffen [2]"] = df["mit Konservierungsstoffen [2]"].map({1: "2"})
    df["mit Farbstoffen [3]"] = df["mit Farbstoffen [3]"].map({1: "3"})
    df["mit Süßungsmitteln [4]"] = df["mit Süßungsmitteln [4]"].map({1: "4"})
    df["enthält eine Phenylalaninquelle [5]"] = df[
        "enthält eine Phenylalaninquelle [5]"
    ].map({1: "5"})
    df["mit Geschmacksverstärker [6]"] = df["mit Geschmacksverstärker [6]"].map({1: "6"})
    df["mit Phosphat [7]"] = df["mit Phosphat [7]"].map({1: "7"})
    df["geschwefelt [8]"] = df["geschwefelt [8]"].map({1: "8"})
    df["gewachst [9]"] = df["gewachst [9]"].map({1: "9"})
    df["geschwärzt [10]"] = df["geschwärzt [10]"].map({1: "10"})
    df["Oberfläche mit Natamycin behandelt [11]"] = df[
        "Oberfläche mit Natamycin behandelt [11]"
    ].map({1: "11"})
    df["chininhaltig [12]"] = df["chininhaltig [12]"].map({1: "12"})
    df["koffeinhaltig [13]"] = df["koffeinhaltig [13]"].map({1: "13"})
    df["mit Alkohol [14]"] = df["mit Alkohol [14]"].map({1: "14"})


    df = df[df.Artikelnummer != ""]
    df = df.fillna("")

    df["Allergenangaben"] = (
        df["… Weizen [A1]"].map(str)
        + ", "
        + df["… Roggen [A2]"].map(str)
        + ", "
        + df["… Gerste [A3]"].map(str)
        + ", "
        + df["… Hafer [A4]"].map(str)
        + ", "
        + df["… Dinkel [A5]"].map(str)
        + ", "
        + df["… Kamut [A6]"].map(str)
        + ", "
        + df["… Mandel [H1]"].map(str)
        + ", "
        + df["… Haselnüsse  [H2]"].map(str)
        + ", "
        + df["… Walnüsse  [H3]"].map(str)
        + ", "
        + df["… Cashewnüsse / Kaschunüsse  [H4]"].map(str)
        + ", "
        + df["… Pecannüsse [H5]"].map(str)
        + ", "
        + df["… Paranüsse [H6]"].map(str)
        + ", "
        + df["… Pistazien  [H7]"].map(str)
        + ", "
        + df["… Macadamianüsse / Queenslandnüsse [H8]"].map(str)
    )

    df["Allergenangaben"] = df["Allergenangaben"].str.replace(r" ", "", regex=True)
    df["Allergenangaben"] = df["Allergenangaben"].str.lstrip(",")
    df["Allergenangaben"] = df["Allergenangaben"].str.rstrip(",")

    df["Allergenangaben"] = df["Allergenangaben"].str.replace(r"\,,", ",", regex=True)
    df["Allergenangaben"] = df["Allergenangaben"].str.replace(r"\,,,", ",", regex=True)
    df["Allergenangaben"] = df["Allergenangaben"].str.replace(r"\,,,,", ",", regex=True)
    df["Allergenangaben"] = df["Allergenangaben"].str.replace(r"\,,,,,", ",", regex=True)
    df["Allergenangaben"] = df["Allergenangaben"].str.replace(r"\,,,,,,", ",", regex=True)
    df["Allergenangaben"] = df["Allergenangaben"].str.replace(r"\,,,,,,,", ",", regex=True)
    df["Allergenangaben"] = df["Allergenangaben"].str.replace(r"\,,,,,,,,", ",", regex=True)
    df["Allergenangaben"] = df["Allergenangaben"].str.replace(
        r"\,,,,,,,,,", ",", regex=True
    )
    df["Allergenangaben"] = df["Allergenangaben"].str.replace(r"\,,,", ",", regex=True)


    df["Allergenangaben"] = df["Allergenangaben"].str.replace(r",", ", ", regex=True)

    df["Antioxidationsmittel [1]"] = df["Antioxidationsmittel [1]"].astype("str")

    df["Zusatzstoffangaben"] = (
        df["Antioxidationsmittel [1]"].map(str)
        + ", "
        + df["mit Konservierungsstoffen [2]"].map(str)
        + ", "
        + df["mit Farbstoffen [3]"].map(str)
        + ", "
        + df["mit Süßungsmitteln [4]"].map(str)
        + ", "
        + df["enthält eine Phenylalaninquelle [5]"].map(str)
        + ", "
        + df["mit Geschmacksverstärker [6]"].map(str)
        + ", "
        + df["mit Phosphat [7]"].map(str)
        + ", "
        + df["geschwefelt [8]"].map(str)
        + ", "
        + df["gewachst [9]"].map(str)
        + ", "
        + df["geschwärzt [10]"].map(str)
        + ", "
        + df["Oberfläche mit Natamycin behandelt [11]"].map(str)
        + ", "
        + df["chininhaltig [12]"].map(str)
        + ", "
        + df["koffeinhaltig [13]"].map(str)
        + ", "
        + df["mit Alkohol [14]"].map(str)
    )

    df["Zusatzstoffangaben"] = df["Zusatzstoffangaben"].str.replace(r" ", "", regex=True)
    df["Zusatzstoffangaben"] = df["Zusatzstoffangaben"].str.lstrip(",")
    df["Zusatzstoffangaben"] = df["Zusatzstoffangaben"].str.rstrip(",")


    df["Zusatzstoffangaben"] = df["Zusatzstoffangaben"].str.replace(r" ", "", regex=True)
    df["Zusatzstoffangaben"] = df["Zusatzstoffangaben"].str.lstrip(",")
    df["Zusatzstoffangaben"] = df["Zusatzstoffangaben"].str.rstrip(",")

    df["Zusatzstoffangaben"] = df["Zusatzstoffangaben"].str.replace(r"\,,", ",", regex=True)
    df["Zusatzstoffangaben"] = df["Zusatzstoffangaben"].str.replace(
        r"\,,,", ",", regex=True
    )
    df["Zusatzstoffangaben"] = df["Zusatzstoffangaben"].str.replace(
        r"\,,,,", ",", regex=True
    )
    df["Zusatzstoffangaben"] = df["Zusatzstoffangaben"].str.replace(
        r"\,,,,,", ",", regex=True
    )
    df["Zusatzstoffangaben"] = df["Zusatzstoffangaben"].str.replace(
        r"\,,,,,,", ",", regex=True
    )
    df["Zusatzstoffangaben"] = df["Zusatzstoffangaben"].str.replace(
        r"\,,,,,,,", ",", regex=True
    )
    df["Zusatzstoffangaben"] = df["Zusatzstoffangaben"].str.replace(
        r"\,,,,,,,,", ",", regex=True
    )
    df["Zusatzstoffangaben"] = df["Zusatzstoffangaben"].str.replace(
        r"\,,,,,,,,,", ",", regex=True
    )
    df["Zusatzstoffangaben"] = df["Zusatzstoffangaben"].str.replace(r"\,,", ",", regex=True)

    df["Zusatzstoffangaben"] = df["Zusatzstoffangaben"].str.replace(r",", ", ", regex=True)

    # Füllt komplette column Lieferanten Name mit ersten Wert der column
    lf_name = df["Lieferant Name"].iloc[0]
    df["Lieferant Name"] = lf_name

    # Füllt komplette column Aramark LieferantNr mit ersten Wert der column
    lf_nr = df["Aramark LieferantNr"].iloc[0]
    df["Aramark LieferantNr"] = lf_nr

    # Columns umbenennen & Ausgabe bestimmen
    df = df.rename(
        columns={
            "Aramark LieferantNr": "Lieferanten Nr.",
            "Lieferant Name": "Lieferantenname",
            "Artikelnummer": "Artikel.Nr.",
            "Artikelname": "Artikelname",
            "Keine Allergene": "Allergene Kennzeichnung",
            "Keine Zusatzstoffe": "Zusatzstoff Kennzeichnung",
            "Krebstiere / Krebstiererzeugnisse [B]": "Krebstiere (AC)",
            "Eier / Eierzeugnisse [C] ": "Eier (AE)",
            "Fisch / Fischerzeugnisse [D]": "Fisch (AF)",
            "Milch / Milcherzeugnisse einschl. Lactose [G]": "Milch (AM)",
            "Schalenfrüchte (Nüsse) und Erzeugnisse [H]": "Nüsse     (AN)",
            "Erdnüsse / Erdnusserzeugnisse [E]": "Erdnüsse (AP)",
            "Sesam / Sesamerzeugnisse [K]": "Sesamsamen (AS)",
            "Schwefeldioxid und Sulfite [L]": "Schwefeldioxid/Sulphite (AU)",
            "Glutenhaltiges Getreide und Erzeugnisse [A]": "Glutenhaltiges Getreide     (AW)",
            "Soja / Sojaerzeugnisse [F]": "Sojabohnen (AY)",
            "Sellerie / Sellerieerzeugnisse [I]": "Sellerie (BC)",
            "Senf / Senferzeugnisse [J]": "Senf (BM)",
            "Lupinen / Lupinenerzeugnisse [M]": "Lupine (NL)",
            "Weichtiere / Weichtiererzeugnisse [N]": "Weichtiere (UM)",
            "Allergenangaben": "Codeliste für Allergene",
            "Zusatzstoffangaben": "Codeliste für Zusatzstoffe",
        }
    )

    #Restliche Columns hinzufügen inkl. default Wert
    df["Zus. Bezeichnung"] = ""
    df["BLS-Schlüssel"] = ""
    df["Alternative Artikelnummer"] = ""
    df["Fixe Bestellmenge"] = ""
    df["Kurzinfo"] = ""
    df["Textdatei"] = ""
    df["Bilddatei"] = ""
    df["Artikel-kennzeichen"] = ""
    df["ILN (Hersteller)"] = ""
    df["Hersteller-Artikelnummer"] = ""
    df["Hersteller"] = ""
    df["Marke"] = ""
    df["ArtikelgruppenWarengruppen"] = ""
    df["Oberwaren-gruppe"] = ""
    df["Klassifikation"] = ""
    df["Bestelleinheit"] = ""
    df["Info zur Bestelleinheit"] = ""
    df["Einzeleinheit"] = ""
    df["Menge der Einzeleinheit in der Bestelleinheit"] = ""
    df["Gewichtsartikel"] = ""
    df["Anbruch"] = ""
    df["Preis-kennzeichen"] = ""
    df["Preis gültig von(jjjjmmtt)"] = ""
    df["Preis gültig bis(jjjjmmtt)"] = ""
    df["Tagespreis- Kennzeichen"] = ""
    df["Preis"] = ""
    df["Preis pro Einzeleinheit"] = ""
    df["Rabatt"] = ""
    df["Preismenge"] = ""
    df["Verkaufs-preis"] = ""
    df["Preiswährung"] = ""
    df["Steuersatz"] = ""
    df["Handlungs-anforderung/-benachritigung"] = ""
    df["Anforderung/Benachrichtigung Bestellbar ab(jjjjmmtt)"] = ""
    df["Artikel bestellbar?"] = ""
    df["Bestellbar ab(jjjjmmtt)"] = ""
    df["Bestellbar bis (jjjjmmtt)"] = ""
    df["EAN Bestelleinheit"] = ""
    df["EAN Packungsart"] = ""
    df["Nettogewicht Abtropfgewicht"] = ""
    df["Basiseinheit"] = ""
    df["Umrechnungs-faktor für BLS"] = ""

    # Output definieren und Columns sortieren
    df = df[
        [
            "Lieferanten Nr.",
            "Lieferantenname",
            "Artikel.Nr.",
            "Artikelname",
            "Zus. Bezeichnung",
            "BLS-Schlüssel",
            "Alternative Artikelnummer",
            "Fixe Bestellmenge",
            "Kurzinfo",
            "Textdatei",
            "Bilddatei",
            "Artikel-kennzeichen",
            "ILN (Hersteller)",
            "Hersteller-Artikelnummer",
            "Hersteller",
            "Marke",
            "ArtikelgruppenWarengruppen",
            "Oberwaren-gruppe",
            "Klassifikation",
            "Bestelleinheit",
            "Info zur Bestelleinheit",
            "Einzeleinheit",
            "Menge der Einzeleinheit in der Bestelleinheit",
            "Gewichtsartikel",
            "Anbruch",
            "Preis-kennzeichen",
            "Preis gültig von(jjjjmmtt)",
            "Preis gültig bis(jjjjmmtt)",
            "Tagespreis- Kennzeichen",
            "Preis",
            "Preis pro Einzeleinheit",
            "Rabatt",
            "Preismenge",
            "Verkaufs-preis",
            "Preiswährung",
            "Steuersatz",
            "Handlungs-anforderung/-benachritigung",
            "Anforderung/Benachrichtigung Bestellbar ab(jjjjmmtt)",
            "Artikel bestellbar?",
            "Bestellbar ab(jjjjmmtt)",
            "Bestellbar bis (jjjjmmtt)",
            "EAN Bestelleinheit",
            "EAN Packungsart",
            "Nettogewicht Abtropfgewicht",
            "Basiseinheit",
            "Umrechnungs-faktor für BLS",
            "Allergene Kennzeichnung",
            "Zusatzstoff Kennzeichnung",
            "Krebstiere (AC)",
            "Eier (AE)",
            "Fisch (AF)",
            "Milch (AM)",
            "Nüsse     (AN)",
            "Erdnüsse (AP)",
            "Sesamsamen (AS)",
            "Schwefeldioxid/Sulphite (AU)",
            "Glutenhaltiges Getreide     (AW)",
            "Sojabohnen (AY)",
            "Sellerie (BC)",
            "Senf (BM)",
            "Lupine (NL)",
            "Weichtiere (UM)",
            "Codeliste für Allergene",
            "Codeliste für Zusatzstoffe",
        ]
    ]
    # Verbesserung, hat oben nicht funktioniert
    df["Fisch (AF)"] = df["Fisch (AF)"].map({1: 2, "": 0})
    df["Erdnüsse (AP)"] = df["Erdnüsse (AP)"].map({1: 2, "": 0})
    df["Glutenhaltiges Getreide     (AW)"] = df["Glutenhaltiges Getreide     (AW)"].map(
        {1: 2, "": 0}
    )
    df["Codeliste für Allergene"] = df["Codeliste für Allergene"].str.replace(r"\, , ", ", ", regex=True)
    df["Codeliste für Zusatzstoffe"] = df["Codeliste für Zusatzstoffe"].str.replace(r"\, , ", ", ", regex=True)

    df.to_excel(
        "C:\\Users\Maximilian.Rasch\\Desktop\\Projekt 400\\Output\\"
        + str(lf_name)
        + "_"
        + str(lf_nr)[:-2]
        + "_"
        + Erstellungsdatum
        + ".xlsx",
        index=False,
        header=Kopfzeile,
    )

u = "C:\\Users\\Maximilian.Rasch\\Desktop\\Projekt 400\\Input\\"
os.chdir(u)




for f in os.listdir():
    print(f)
    konvertieren(f)
    



