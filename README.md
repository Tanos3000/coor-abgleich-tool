# Coor-Abgleich-Tool

Vergleicht automatisch die C16-Buchungsexport-Datei (Coor) mit der eigenen
Eingangsrechnungen-Excel, markiert Abweichungen gelb und erzeugt zusätzlich
eine Version mit automatisch behobenen eindeutigen Fehlern.

## Was das Tool prüft (Annahmen)

- **C16-Datei**: flache Tabelle mit Spalten `Belegnr.`, `Belegnr.2`, `Istbetrag`.
- **Eigene Datei**: pro Rechnung ein Kopfblock mit `Ext. Re. Nr 1` / `Ext. Re. Nr 2`,
  darunter eine oder mehrere Zahlungszeilen mit `Zahlung Netto` / `Zahlung Brutto`.
- **Verknüpfung**: `Ext. Re. Nr 1/2` wird mit `Belegnr.`/`Belegnr.2` abgeglichen
  (exakt sowie über den Präfix vor dem ersten `/`, da Ratenzahlungen in der
  C16 oft als `BASISNR/laufende-Nr` auftauchen).
- **Vergleich**: jede einzelne Zahlungszeile wird geprüft, ob ihr Betrag
  (Toleranz 0,02 €) unter dem zugehörigen Beleg in der C16 vorkommt.
- **Gelb markiert** werden nur die Betrag-Zellen (Netto/Brutto) abweichender
  Zeilen, mit Kommentar zur Begründung.
- **Automatisch korrigiert** wird nur der eindeutige Fall: genau eine offene
  Zahlung im Block + genau ein noch nicht zugeordneter Betrag in der C16.
  Alles andere (mehrere Zahlungen, kein passender Beleg gefunden, mehrdeutig)
  bleibt gelb markiert zur manuellen Prüfung – es wird nichts geraten.

Getestet mit den echten Beispieldateien (Villa Sophia / CS63): 148 Rechnungsblöcke,
110 Zahlungszeilen, 56 Treffer, 1 eindeutige Korrektur (echter Zahlendreher
6.970 € statt korrekt 300 €), 53 Fälle zur manuellen Prüfung.

**Wichtig:** Stichprobenartig gegenprüfen, bevor die korrigierte Datei
weiterverwendet wird – das Tool ersetzt keine fachliche Prüfung.

## Bedienung

1. Tool starten (Doppelklick auf `CoorAbgleichTool.exe`).
2. Links die C16-Datei hineinziehen (oder anklicken zum Auswählen).
3. Rechts die eigene Eingangsrechnungen-Excel hineinziehen.
4. Auf „Abgleichen" klicken, Zielordner für die Ergebnisse wählen.
5. Ausgabe: `<dateiname>_markiert.xlsx` (Original + gelbe Markierungen) und
   `<dateiname>_korrigiert.xlsx` (zusätzlich mit den eindeutigen Korrekturen).

## Wie du aus dem Code eine Windows-.exe bekommst

Empfohlener Weg: GitHub Actions baut die .exe automatisch, sobald der Code
in ein Repo gepusht wird – du brauchst dafür **keinen** eigenen Windows-Rechner.

1. Auf github.com ein neues (privates) Repository anlegen, z. B. `coor-abgleich-tool`.
2. Diesen gesamten Ordner (`app.py`, `core.py`, `requirements.txt`,
   `.github/workflows/build-windows-exe.yml`, `README.md`) in das Repo pushen
   – entweder per `git push` oder über „Add file → Upload files" im Browser.
3. Im Repo auf den Reiter **Actions** gehen. Der Workflow „Build Windows EXE"
   startet automatisch nach dem Push (dauert ca. 2–3 Minuten).
4. Ist der Lauf grün, unten auf der Lauf-Seite unter **Artifacts** die Datei
   `CoorAbgleichTool-windows` herunterladen – darin liegt `CoorAbgleichTool.exe`.
5. Diese .exe kannst du beliebig an Kolleg:innen weitergeben (z. B. per Mail,
   Teams, gemeinsamem Laufwerk). Kein Python, keine Installation nötig –
   einfach doppelklicken.

**Alternative ohne GitHub:** Wenn du kurzfristig Zugriff auf einen
Windows-Rechner mit Python hast, reicht auch:
```
pip install -r requirements.txt
pyinstaller --onefile --windowed --name CoorAbgleichTool --collect-all tkinterdnd2 app.py
```
Die fertige `CoorAbgleichTool.exe` liegt danach im Ordner `dist`.

## Bekannte Grenzen / mögliche Weiterentwicklung

- Die Spaltenpositionen (F/G für Netto/Brutto, B für Gewerke-Nr etc.) sind
  fest im Code hinterlegt, weil die Vorlage laut Aufgabenstellung immer
  gleich aussieht. Ändert sich das Excel-Layout, muss `core.py` angepasst werden.
- Bei „Kein passender Beleg in C16 gefunden" kann das auch bedeuten, dass die
  Buchung in der C16 schlicht noch nicht erfolgt ist (z. B. weil sie erst
  später gebucht wird) – nicht zwangsläufig ein Fehler in der eigenen Datei.
