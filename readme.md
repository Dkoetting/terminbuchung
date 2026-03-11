\# Terminbuchung



Einfache Terminbuchungsseite mit HTML/JavaScript und Google Apps Script als Backend.



\## Funktionen



\- Anzeige freier Zeitfenster

\- Blockierung bereits gebuchter Termine aus Google Sheets

\- Buchung über Google Apps Script Web App

\- Synchronisation mit Google Sheets

\- ICS-Download für Outlook/Kalender



\## Projektstruktur



\- `index.html` – Frontend der Terminbuchung

\- `assets/` – Bilder, Logos, statische Dateien

\- `data/` – optionale lokale Daten

\- `Code.gs` – Google Apps Script Backend (nicht im Frontend-Hosting enthalten)



\## Lokaler Test



```bash

python -m http.server 8001 --bind 127.0.0.1

