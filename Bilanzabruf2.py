#C:/Users/TimFe/AppData/Local/Programs/Python/Python313/python.exe -m pip install -r requirements.txt
#requests
#beautifulsoup4
#lxml
#html5lib
#pandas
#openpyxl
import requests
from bs4 import BeautifulSoup
import pandas as pd
import time

def scrape_gemeinde(gemeinde_name, url):
    try:
        response = requests.get(url, timeout=10)
        response.raise_for_status()
    except Exception as e:
        print(f"Fehler bei {gemeinde_name}: {e}")
        return None

    soup = BeautifulSoup(response.text, "html.parser")
    cards = soup.find_all("div", class_="card-body")

    rows = []
    for card in cards:
        cardtitle_tag = card.find("h5", class_="card-title")
        if not cardtitle_tag:
            continue
        cardtitle = cardtitle_tag.text.strip()

        # Beende das Parsing, wenn "Mehrjahresansicht" kommt (nicht relevant)
        if "Mehrjahresansicht" in cardtitle:
            break

        rows.append([cardtitle])  # Card-Überschrift

        # Alle Label-Wert-Paare extrahieren
        saldo_labels = card.find_all("span", class_="saldo-label")
        saldo_values = card.find_all("span", class_="saldo-value")
        for label, value in zip(saldo_labels, saldo_values):
            rows.append([label.text.strip(), value.text.strip()])

        rows.append([])  # Leerzeile zur Trennung

        # Falls dieser Block der Ergebnishaushalt ist → danach stoppen
        if cardtitle.lower().startswith("ergebnishaushalt"):
            break

    return rows

def main():
    # 🔁 Deine Liste mit Gemeindenamen und URLs
    gemeinden = [
        ("Wien", "https://www.offenerhaushalt.at/gemeinde/wien"),
        ("Eisenstadt", "https://www.offenerhaushalt.at/gemeinde/eisenstadt"),
        ("Linz", "https://www.offenerhaushalt.at/gemeinde/linz"),
        ("Klagenfurt", "https://www.offenerhaushalt.at/gemeinde/klagenfurt"),
        ("Salzburg", "https://www.offenerhaushalt.at/gemeinde/salzburg"),
        ("St. Pölten", "https://www.offenerhaushalt.at/gemeinde/sankt-pölten"),
        ("Bregenz", "https://www.offenerhaushalt.at/gemeinde/bregenz"),
        ("Innsbruck", "https://www.offenerhaushalt.at/gemeinde/innsbruck"),
        ("Graz", "https://www.offenerhaushalt.at/gemeinde/graz"),
        # ⚠️ ...füge hier alle weiteren Gemeinden hinzu ...
    ]

    print(f"Starte Scraping für {len(gemeinden)} Gemeinden ...")

    with pd.ExcelWriter("Bilanz_Alle_Gemeinden.xlsx", engine="openpyxl") as writer:
        for i, (name, url) in enumerate(gemeinden, 1):
            print(f"[{i}/{len(gemeinden)}] Verarbeite {name} ...")
            data = scrape_gemeinde(name, url)

            if not data:
                print(f"⚠️ Keine Daten für {name}, übersprungen.")
                continue

            df = pd.DataFrame(data)
            sheet_name = name[:31]  # Excel-Sheetnamen dürfen max. 31 Zeichen haben

            try:
                df.to_excel(writer, sheet_name=sheet_name, index=False, header=False)
            except Exception as e:
                print(f"Fehler beim Schreiben von {sheet_name}: {e}")

            time.sleep(1)  # Kurze Pause, um Server nicht zu überlasten

    print("✅ Fertig! Datei 'Bilanz_Alle_Gemeinden.xlsx' wurde erstellt.")

if __name__ == "__main__":
    main()

