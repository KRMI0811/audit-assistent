# Audit-Assistent Qualitaetsmanagement (Webversion)

Internes Audit nach ISO 9001 planen, durchfuehren und auswerten.
Dialektisches Verfahren: ein Agententeam baut den Konformitaetsnachweis,
ein zweites greift ihn an, der menschliche Auditor entscheidet als Richter.

## Dateien
- app.py                 Die Anwendung
- requirements.txt       Benoetigte Pakete
- normen/ISO_9001.json   Normkatalog, fachlich zu pruefen und zu ergaenzen
- secrets_beispiel.toml  Vorlage fuer die Zugangsdaten, NICHT auf GitHub mit echtem Schluessel
- logo.png               Optional, erscheint oben links

## Einrichtung in Streamlit Community Cloud
1. Dieses Verzeichnis in ein GitHub-Repository legen (privat moeglich).
2. Auf share.streamlit.io anmelden, "Create app", Repository waehlen, Hauptdatei app.py.
3. Unter Settings > Secrets den Inhalt von secrets_beispiel.toml einfuegen
   und API_KEY sowie APP_PASSWORD ersetzen.
4. Unter Settings > Sharing die App auf privat stellen und Personen einladen.

## Hinweise
- Der Arbeitsstand liegt nur in der Sitzung. Ueber die Seitenleiste sichern und wieder laden.
- Keine vertraulichen Originalunterlagen hochladen, solange Datenschutz und IT nicht zugestimmt haben.
- Die Anforderungen im Normkatalog sind Zusammenfassungen in eigenen Worten
  und ersetzen den lizenzierten Normtext nicht.
