# Synthetische Daten Generator

**Für Wirtschaftsprüfer, Steuerberater und Buchhalter, die KI-Tools nutzen wollen – ohne echte Mandantendaten preiszugeben.**

> Excel-Datei rein → synthetische Kopie raus → mit KI arbeiten, DSGVO-konform.

---

## Das Problem

KI-Tools wie Claude, ChatGPT oder Copilot können die Arbeit von Wirtschaftsprüfern erheblich
beschleunigen – Berechnungen, Auswertungen, Plausibilitätsprüfungen. Aber:

**Echte Mandantendaten dürfen nicht in externe KI-Systeme hochgeladen werden.**

Dieses Tool löst das Problem: Es ersetzt alle sensiblen Werte durch realistische, aber erfundene
Alternativen – in Sekunden, vollständig lokal auf Ihrem Rechner.

---

## Features

- **100% lokal** – keine Internetverbindung, keine Cloud, kein API-Aufruf
- **Formatierung bleibt erhalten** – Farben, Fettschrift, Spaltenbreiten, Rahmen identisch
- **Deutsche Datentypen** – IBAN, Steuernummer, USt-IdNr, Handelsregisternummer
- **SAP-Kürzel** – KdNr, BuchDat, LfdNr, DebNr, BelNr, Re-Nr, BtrgNetto, ...
- **Konsistentes Mapping** – gleiche Kundennummer → immer gleicher Fake-Wert
- **Beträge im Originalbereich** – Verteilung und Größenordnung bleiben plausibel
- **Vorschau vor Generierung** – erkannte Typen prüfen und korrigieren
- **Auditbericht** – HTML-Dokument als Nachweis der datenschutzkonformen Verarbeitung
- **Kein Python nötig** – einfach `.exe` herunterladen und starten

---

## Unterstützte Datentypen

| Kategorie | Felder |
|-----------|--------|
| Personen | Vorname, Nachname, Vollständiger Name |
| Organisation | Firma, Unternehmen, Handelsregisternummer |
| Adresse | Straße, PLZ, Ort, Land |
| Kontakt | E-Mail, Telefon, Fax |
| Finanzen | IBAN, BIC/SWIFT, Betrag, Netto, Brutto |
| Steuer | Steuernummer, USt-IdNr |
| Nummern | Kunden-Nr, Rechnungs-Nr, Belegnummer (inkl. SAP-Kürzel) |
| Zeit | Datum, Buchungsdatum, Fälligkeitsdatum |

Nicht erkannte Spalten (z. B. Kostenstelle, Buchungstyp) bleiben **unverändert**.

---

## Installation

### Option A – EXE (empfohlen, kein Python nötig)

1. Neueste `SynthetischeDatenGenerator.exe` aus [Releases](../../releases) herunterladen
2. Doppelklick – fertig

### Option B – Python

```bash
pip install openpyxl faker
python app.py
```

---

## Verwendung

1. **Datei auswählen** – Original-Excel öffnen
2. **Vorschau prüfen** – erkannte Spaltentypen kontrollieren, bei Bedarf korrigieren
3. **Erstellen klicken** – synthetische Datei wird generiert
4. **Zwei Dateien erhalten:**
   - `dateiname_synthetisch.xlsx` – zum Hochladen in KI-Tools
   - `dateiname_synthetisch_Auditbericht.html` – Nachweis für die Dokumentation

---

## Auditbericht

Nach jeder Verarbeitung wird automatisch ein HTML-Bericht erstellt:

- Zeitstempel der Verarbeitung
- Gerätename (Hostname)
- Welche Spalten anonymisiert wurden (und wie)
- Bestätigungsvermerk für interne Dokumentation
- Druckbar als PDF (Browser → Drucken → Als PDF speichern)

---

## Selbst bauen (EXE)

```bash
pip install openpyxl faker pyinstaller
build_exe.bat
```

Die EXE erscheint in `dist/SynthetischeDatenGenerator.exe` (~16 MB, keine weiteren Dateien nötig).

---

## Technisches

- **Sprache:** Python 3.10+
- **Abhängigkeiten:** `openpyxl`, `faker` – kein pandas, kein numpy
- **GUI:** tkinter (Python-Standardbibliothek)
- **Fake-Daten:** Faker mit `de_DE` Locale
- **Konsistenz:** Gleicher Originalwert → immer gleicher Fake-Wert (innerhalb einer Datei)
- **Beträge:** Gleichmäßig verteilt im Bereich [Min, Max] der Originalspalte

---

## Lizenz

MIT – kostenlos nutzbar, auch kommerziell. Siehe [LICENSE](LICENSE).

---

## Mithelfen

Pull Requests willkommen – besonders für:
- Weitere SAP-Feldnamen / Kürzel
- Österreichische und Schweizer Datentypen (UID-Nummer, AHV-Nr, ...)
- Unterstützung für CSV-Dateien
