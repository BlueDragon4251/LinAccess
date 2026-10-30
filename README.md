# LinAccess Native (C++20 / Qt 6) — Version 0.7.0, Entwicklungsstand

Native Linux-Desktopanwendung als Grundlage für einen Access-ähnlichen
Datenbankdesigner. **Kein Python-Runtimebedarf.**

> **Wichtig:** Diese Version interpretiert nur eine **experimentelle,
> eingeschränkte VBA-Teilmenge** und kann Access-2003-Formulare aus dem
> Bridge-Export mit **annähernder Originalanordnung der Standardcontrols**
> öffnen. Sie ist
> kein Ersatz für Microsoft Access in der Produktion. Lies vor einem Import
> [COMPATIBILITY.md](docs/COMPATIBILITY.md).

## Build unter Ubuntu / Debian

Abhängigkeiten (je nach Distribution können die Paketnamen abweichen):

```bash
sudo apt update
sudo apt install build-essential cmake qt6-base-dev libqt6sql6-sqlite mdbtools

cmake -S . -B build -DCMAKE_BUILD_TYPE=Release
cmake --build build -j
./build/linaccess
```

Eine `.deb`-Datei lässt sich nach dem Build erzeugen mit:

```bash
cpack --config build/CPackConfig.cmake -G DEB
```

Die Datei `.github/workflows/build-linux.yml` enthält zusätzlich einen
GitHub-Actions-Build für Ubuntu 24.04. Er wurde hier noch nicht ausgeführt.

`mdbtools` ist optional, aber zum Lesen von `.mdb` und versuchsweise `.accdb`
im Access-Inspektor erforderlich. Die Unterstützung ist je nach MDBTools-Version
und Dateiformat unterschiedlich. Access-Dateien werden dabei nie überschrieben.

Nur die C++-Parser-Tests (ohne Qt) mit GCC:

```bash
g++ -std=c++20 -Wall -Wextra -pedantic -Isrc \
    tests/core_tests.cpp src/vba_analyzer.cpp src/vba_runtime.cpp \
    src/vba_expression.cpp src/csv.cpp -o core_tests
./core_tests
```

Oder, wenn CMake verfügbar ist, ohne Qt:

```bash
cmake -S . -B build-core -DLINACCESS_BUILD_GUI=OFF
cmake --build build-core
ctest --test-dir build-core --output-on-failure
```

## Funktionen in dieser Entwicklungsstufe

- Neue `.linaccess`-Datei als SQLite-Projekt, oder bestehendes SQLite-Projekt öffnen
- Tabellen mit einfachen Datentypen erstellen, Datenansicht mit manueller Speicherung
- Gespeicherte SQL-SELECT-Abfragen
- Einfache tabellengebundene Formulare definieren und Datensätze damit bearbeiten
- Pro Formular VBA-Quelltext speichern; `.bas`/`.cls`-Textdateien laden und auf
  Prozeduren bzw. Abhängigkeiten statisch untersuchen. Originalbytes importierter
  Module werden zusätzlich als Base64 gesichert; die Editoransicht geht noch
  von UTF-8 aus.
- **VBA-Teilinterpretation** für native Formulare und importierte VBIDE-Module:
  `Form_Load`, `btnSpeichern_Click`, `btnAktion_Click`, `MsgBox`,
  `If/Else/End If`, `Dim`, Variablen, Rechenausdrücke und Vergleiche,
  `For/Next`, `Exit Sub/Function`, `Me!Feld = ...`, `&`-Verkettung,
  einige Standardfunktionen (`IsNull`, `Nz`, `Len`, `CStr`, `CLng`,
  `CDbl`, `UCase`, `LCase`, `Abs`), `DoCmd.RunCommand acCmdSaveRecord`,
  `DoCmd.Maximize` und `DoCmd.GoToRecord , , acLast`.
  Das Ereignis `Form_Open(Cancel As Integer)` wird unterstützt und kann
  durch Setzen von `Cancel` die Formularöffnung abbrechen.
  **Keine vollständige Microsoft-VBA-Runtime, keine DAO/COM-API.**
- SQL-basierte Tabellenberichte als HTML-Vorschau und PDF
- Microsoft Access-Dateien über MDBTools **lesend** untersuchen: Tabellen,
  gespeicherte SQL-Abfragen, sichtbare Formular-/Bericht-/Modulnamen und Beziehungen
  aus dem Jet-Katalog `MSysObjects`. Abfragen werden auf mögliche fehlende
  Tabellenreferenzen geprüft.
- Optionaler *verlustbehafteter* Tabellenimport mit `INTEGER`, `REAL`, `TEXT`
  und echten `NULL`-Werten. `BLOB`-Spalten können übernommen werden, solange
  ihre Zellen NULL sind; echte OLE-Nutzdaten erfordern einen gesonderten
  Binäradapter. Schlüssel, Indizes, Beziehungen und Standardwerte bleiben
  **nicht** originalgetreu.
- Im MDB-Inspektor können **alle Tabellen auf einmal** ins LinAccess-Projekt
  kopiert werden. Vorhandene Tabellen bleiben stehen. Pro Tabellenimport wird
  ein individuelles Ergebnis angezeigt, nicht lesbare Tabellen werden gemeldet.
- Objekt-Metadaten als `metadata_only` mit Ursprungsdatei-Hash im eigenen
  Projekt katalogisieren (ohne ausführbaren VBA-/Formulardesigner-Import).
- Importprotokoll `_la_import_log` für nachträgliche Nachvollziehbarkeit
- Access-Bridge-Import für Quellartefakte aus einer separaten Windows/Access-
  Exportbrücke mit SHA-256-Prüfung und transaktionaler SQLite-Speicherung
- Importierte VBA-Module können im Objektkatalog auf Prozeduren untersucht
  und per Klick mit dem eingeschränkten C++-Interpreter
  ausgeführt werden. Der Import eines VBIDE-Moduls und anschließende Ausführung
  seiner `Form_Load`-Prozedur sind automatisiert getestet.
- Die Bridge zeigt nach dem Import eine statische VBA-Zusammenfassung mit
  Prozeduranzahl und erkannten Access-/Windows-Abhängigkeiten. Deutsche
  VBA-Ereignisnamen mit Umlauten werden dabei erkannt; einfache Prozeduren
  mit solchen Namen können in der eingeschränkten Laufzeit starten.
- VBA aus einem Bridge-Paket kann in ein natives LinAccess-Formular übernommen,
  dem Formularereignis zugeordnet und dort mit Feld-/Speicher-Host ausgeführt werden.
  UTF-8/UTF-16LE und Windows-1252 werden in der Editoransicht unterstützt;
  Originalbytes bleiben separat gespeichert.
- Access-2003-`SaveAsText`-Formulare können jetzt im ursprünglichen
  **TWIP-Koordinatensystem** als Qt-Layout geöffnet werden. Unterstützte
  Controlarten: Textfelder, Labels, Buttons, Kombinationslisten, Checkboxen,
  Unterformular-Startschaltflächen und einfache Rahmen. Gebundene direkte
  SQLite-Tabellen können über diese Ansicht gelesen, bearbeitet und gespeichert
  werden. Für direkt gebundene Tabellen dient SQLite-`rowid` als Schlüssel.
- Klicks auf importierte Schaltflächen rufen `Name_Click` im originalen
  `Form_...`-VBA-Modul auf. `Form_Open`, `Form_Load` und `Form_Current` werden
  beim Öffnen versucht, ohne zusätzliches VBA-Zustimmungsfenster. Neben den
  bisherigen Befehlen kennt der Interpreter nun `DoCmd.RunSQL`,
  `CurrentDb.Execute`, `DoCmd.OpenForm`, `GoToRecord acNext/acNewRec`,
  `DoCmd.Requery`, `Me.Control.Visible`, `.SetFocus`, `.Requery`,
  parameterloses `Call` und VBA-Zeilenfortsetzung `_`.
  **SQL-Aktionen wirken direkt auf die geöffnete SQLite-Projektkopie**, auch
  bei `DELETE`, `UPDATE`, `DROP TABLE` und `INSERT`. Es erscheint keine
  LinAccess-spezifische Sicherheitsabfrage vor jeder VBA-Anweisung.
- Importierte Access-Abfragen lassen sich in einer SQLite-Abfrageansicht
  ausführen, soweit ihre Jet-SQL-Syntax kompatibel ist. Importierte Berichte
  haben eine grobe Layoutvorschau (noch keine Layout-/Druckkompatibilität).
- Für ein per Windows-Bridge importiertes Formular gibt es die Aktion
  **Als natives Formular übernehmen (Teilkonvertierung)**. Sie liest
  `RecordSource` und gebundene `ControlSource`-Felder aus `SaveAsText`, gleicht
  sie mit einer bereits in SQLite importierten Tabelle ab und verknüpft das
  gleichnamige importierte `Form_...`-VBA-Modul mit dem neuen Formular.
  Die ältere **Teilkonvertierung** übernimmt weiterhin nicht das
  Access-Layout; hierfür die neue gesonderte Access-Layoutansicht benutzen.
- Native Formulare können auch auf importierten SQLite-Tabellen ohne `id`-Spalte
  arbeiten (SQLite-`rowid`-Fallback).

**Nicht implementiert:** vollständige Original-VBA-Ausführung, wirklich
pixelidentische Form-/Report-Rekonstruktion, eingebettete Access-Makros,
DAO-Recordsets/ADO, ActiveX, Windows-COM, Jet/ACE-SQL und 1:1-Konvertierung.

## Reale Access-2003-Regression

Mit einer echten Access-2003-Datei `Keksrezepte.mdb` (JET4) wurde der neue
Adapter getestet: **2 Tabellen mit 4 + 35 Zeilen, 4 Formulare, 1 Bericht,
3 Benutzerabfragen und 1 Beziehung**. Der typisierte Import, NULL-Erhaltung,
ein fehlgeschlagener Doppelimport und der unveränderte SHA-256-Fingerabdruck
der Originaldatei sind automatisiert geprüft. Weitere 16 versteckte
Access-Systemabfragen wurden katalogisiert.

Details: [Testbericht](docs/KEKSREZEPTE_2003_TEST.md).

Der vom Nutzer mit Access 365 erzeugte **echte** Bridge-Export dieser Access-2003-
MDB wurde zusätzlich erfolgreich auf Linux eingelesen: **26 Quellobjekte**, davon
**4 Formulare, 1 Bericht, 19 Abfragen und 2 VBA-Klassenmodule**. Der Quelltext
enthält **8 Ereignisprozeduren**. Eine der beiden Prozeduren mit ausführbaren
Anweisungen (`Form_Open`) wurde in der C++-Laufzeit mit ihren zwei realen
`DoCmd`-Hostaktionen getestet. Die andere (`Menge__g__AfterUpdate`) verwendet
noch nicht implementierte Access-Formreferenzen, `SetFocus`, `Close` und
`OpenForm`. Sie läuft **nicht** originalgetreu. Siehe
[Bridge-Regression](docs/KEKSREZEPTE_BRIDGE_TEST.md).

Ein zweiter umfangreicher Realtest aus **Access 2003 selbst** liegt ebenfalls
vor: `Kassa_CSE_V2.0.mdb`. Die Windows-Bridge hat **192 Objekte** fehlerfrei
exportiert, alle Hashes wurden überprüft. Auf Linux wurden 47 Formulare,
9 Berichte, 100 Abfragen, 4 Access-Module und 32 VBA-Dateien importiert.
Der statische Analyzer erkennt **187 VBA-Prozeduren**, führt beim Import
jedoch **keinen** fremden Code aus. 23 Formulare nutzen direktes SQL als
Datenquelle, lediglich 10 sind unmittelbar tabellengebunden.
Die Bestandsanwendung ist **noch nicht lauffähig**, insbesondere wegen
DAO/Recordsets, `CurrentDb`, Formularobjekten und diversen `DoCmd`-Befehlen.
In Version 0.7 wurden alle **47 realen Access-2003-Formdefinitionen** mit
dem Layout-Parser geprüft: 685 Platzierungsdaten und 43 Schaltflächen wurden
erkannt. Ein echter Qt-Klick mit VBA-`CurrentDb.Execute` wird gegen eine
separate SQLite-Testtabelle ausgeführt. Das belegt **nicht**, dass die
Original-Kassenaktionen bereits laufen.
Auch der tabellarische MDB-Import wurde direkt an der originalen Kassen-MDB
getestet: **29 von 29 Tabellen** konnten als typisierte SQLite-Kopien angelegt
werden. Die vorhandene `Foto`-BLOB-Spalte enthielt im Test keine nicht-NULL-
Werte. Das ist kein Nachweis für PK-/FK-/Index- oder DAO-Kompatibilität.
Details: [Kassen-Regression](docs/KASSA_ACCESS2003_BRIDGE_TEST.md).

Die vom Anwender bereitgestellte MDB-Datei wird **nicht** mit dem Code
veröffentlicht. Für eigene lokale Regressionstests:

```bash
cmake -S . -B build -DLINACCESS_REAL_MDB_FILE=/pfad/zu/Keksrezepte.mdb
cmake --build build -j2
ctest --test-dir build --output-on-failure
```

Der echte Windows-Bridge-Regressionstest lässt sich mit
`-DLINACCESS_REAL_BRIDGE_DIR=/pfad/zu/Keksrezepte-Bridge` aktivieren.
Für das komplexe Projekt gibt es den weiteren optionalen Pfad
`-DLINACCESS_COMPLEX_BRIDGE_DIR=/pfad/zu/Kassa_CSE_V2.0-Bridge-2003-Retry`.
Die Original-MDB lässt sich über
`-DLINACCESS_COMPLEX_MDB_FILE=/pfad/zu/Kassa_CSE_V2.0.mdb` zusätzlich als
vollständiger Tabellenkopie-Regressionstest verwenden.

## Access Bridge: Metadaten auf Windows exportieren

Auf einem Windows-PC **mit Microsoft Access**:

```powershell
powershell.exe -NoProfile -File .\bridge\Export-LinAccessBridge.ps1 `
    -DatabasePath "C:\Data\Kunden.accdb" `
    -OutputDirectory "C:\Data\LinAccessExport"
```

Mit zwei installierten Access-Versionen verwendet `Access.Application` den
als Standard registrierten COM-Server, auf dem Nutzerrechner war das **Access
365**. Für einen neuen, tatsächlich mit **Access 2003** durchgeführten Export:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\bridge\Export-LinAccessBridge.ps1 `
    -DatabasePath "C:\Data\Keksrezepte.mdb" `
    -OutputDirectory "C:\Data\LinAccessExport-2003" `
    -AccessProgId "Access.Application.11" `
    -RequiredAccessMajorVersion 11
```

Die Version wird nach COM-Aktivierung geprüft, **vor** dem Öffnen der
temporären MDB-Kopie. Bei fehlender Registrierung oder Rückgabe von Version
`16.x` bricht das Skript ab, statt behauptete Access-2003-Daten auszugeben.
Eine Zweitinstallation registriert ihre COM-Version nicht zwangsläufig
separat. `-ExecutionPolicy Bypass` gilt nur für diesen Prozess; Skripte vor
Ausführung prüfen und ausschließlich aus vertrauenswürdiger Quelle starten.

- Der Zielordner muss leer sein. Das Skript öffnet eine temporäre **Kopie** der
  Access-Datei und prüft den SHA-256-Fingerabdruck des Originals vor und nachher.
- Es versucht, Automatisierungsmakros vor dem Öffnen zu deaktivieren und bricht
  ab, wenn diese Sicherheitsbedingung nicht bestätigt werden kann. Trotzdem
  ausschließlich eigene oder vertrauenswürdige Access-Dateien verwenden.
- `SaveAsText` ist keine zugesicherte, stabile Access-Konvertierungs-API.
  Formulare und Berichte können je nach Version fehlschlagen. Fehler werden
  im `manifest.json` sichtbar, nicht stillschweigend ignoriert.
- VBIDE-Modul-Export benötigt gegebenenfalls vertrauenswürdigen Zugriff auf
  das VBA-Projekt. Das Skript verändert keine Trust-Center-Einstellungen.
- Wenn `VBComponents.Export` für ein Modul fehlschlägt, versucht die Bridge
  als Fallback die tatsächlichen `CodeModule.Lines` als UTF-8-Quelltext zu sichern.
  Das erhält den ausführbaren Quelltext, aber keine Designer-/Klassen-Binärdaten.
- In Linux: `.linaccess`-Projekt erstellen → **Access untersuchen** →
  `.mdb` auswählen → **Alle Tabellen in Projektkopie übernehmen** →
  **Bridge-Export öffnen** → Exportordner wählen.
  Artefakte werden zunächst nur kopiert und geprüft. Anschließend können
  importierte Formulare über **Access-Formular öffnen** mit den positionsgetreu
  gelesenen Standardcontrols aufgerufen und unterstützte VBA-Events ausgeführt
  werden; nicht unterstützte Access-Funktionen melden weiterhin Fehler.
- Für formgebundene Ereignisse: Formulardesigner öffnen →
  **VBA aus Bridge übernehmen** → Modul auswählen → speichern → Formular öffnen.
  Bei einfachen tabellengebundenen Formularen alternativ das importierte
  Formular öffnen und **Als natives Formular übernehmen** wählen. Ein
  vorheriger Import seiner Tabelle ist nötig; der Export reproduziert
  ausdrücklich nicht dessen ursprüngliche UI.
- Die Bridge wurde mit einem **vom Anwender unter Microsoft Access 365**
  exportierten echten Access-2003-Projekt auf der Linux-Importseite getestet.
  Der Windows-seitige Export wurde dabei nicht selbst instrumentiert; aus
  einem fehlerfreien Manifest folgt keine allgemeine VBA-Kompatibilität.
- Für die Linux-Importoberfläche liegt unter `examples/bridge_sample` ein
  **synthetisches**, SHA-256-konsistentes Testpaket. Es enthält **keine echte**
  Access-Datei und testet deshalb keine Microsoft-Access-Kompatibilität.

## Experimentelle VBA-Ausführung
