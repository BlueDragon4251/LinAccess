# LinAccess Native (C++20 / Qt 6) — Version 0.7.0, Entwicklungsstand

Native Linux-Desktopanwendung als Grundlage für einen Access-ähnlichen
Datenbankdesigner. **Kein Python-Runtimebedarf.**

> **Wichtig:** Diese Version interpretiert nur eine **experimentelle,
> eingeschränkte VBA-Teilmenge** und kann Access-2003-Formulare aus dem
> Bridge-Export mit **annähernder Originalanordnung der Standardcontrols**
> öffnen. Sie ist
> kein Ersatz für Microsoft Access in der Produktion. Lies vor einem Import
> [COMPATIBILITY.md](docs/COMPATIBILITY.md).

Der [feste Arbeitskatalog mit 456 Punkten](docs/FIXED_WORK_CATALOG.md) ordnet
die Entwicklung in fünf Kategorien und 14 Unterkategorien. IDs und Gesamtzahl
bleiben fest; bestätigte Teilumfänge sind keine Access-2003-Paritätsquote.

## Build unter Ubuntu / Debian

Dieses Repository enthält zwei Access-2003-Testdatenbanken samt originalen
Bridge-Exporten in einem komprimierten, verlustfrei wiederherstellbaren
Git-Testpaket. Vor den vollständigen Regressionstests ausführen:

```bash
sudo apt install xz-utils unzip
./scripts/restore-fixtures.sh
```

Die Original-MDBs werden anhand von SHA-256 überprüft; nähere Hinweise zur
Vertraulichkeit und Pfadstruktur stehen unter [fixtures](fixtures/README.md).
Die vorhandene App ist **noch nicht Access-kompatibel genug, um das komplexe
Kassensystem vollständig auszuführen**. Der konkrete Entwicklungsauftrag an
Codex steht in [CODEX_HANDOFF.md](CODEX_HANDOFF.md).
Der Gesamtimport ist jetzt atomar und übernimmt exportierte Schlüssel, Indizes
und Beziehungen. Die Kassen-Regression fordert 29/29 Tabellen und prüft sämtliche
exportierten Zellwerte. Der Schema-Parserfehler mit MDBTools 1.0.1 ist behoben;
die früheren intermittierenden MDBTools-1.0.0-/EOF-Fehler sind hier bislang nicht
reproduziert. Details und verbleibende Grenzen:
[Import-Untersuchung](docs/MDB_IMPORT_REGRESSION.md).

Abhängigkeiten (je nach Distribution können die Paketnamen abweichen):

```bash
sudo apt update
sudo apt install build-essential cmake qt6-base-dev libqt6sql6-sqlite mdbtools dpkg-dev file

cmake -S . -B build -DCMAKE_BUILD_TYPE=Release
cmake --build build -j
./build/linaccess
```

Der GUI-Build kompiliert SQLite 3.54.1 aus dem mitgelieferten, per SHA3-256
geprüften offiziellen Quellarchiv. Dafür ist kein Download beim Build nötig.
Die private Bibliothek ersetzt kein Systempaket. Qt-SQL-Treiber und native
Adapter werden vor Verwendung auf dieselbe Backend-Identität geprüft;
Hintergrund und Aktualisierung: [SQLite-Backend](third_party/sqlite/README.md).
Auch ohne grafische Sitzung prüfbar: `./build/linaccess --check-sql-backend`.

VBA-Eingabedialoge `InputBox(...)`/`InputBox$(...)` verwenden einen echten
Qt-Dialog mit Text-/Abbruchrückgabe. Geprüfter Umfang und verbleibende Grenzen:
[InputBox-Laufzeit](docs/INPUT_BOX_RUNTIME.md).

DAO-Clone-/Bookmark-/Find-Suche kann den nativen Formulardatensatz wechseln,
einschließlich Validierung und zweier originaler generierter Suchhandler auf
vorbereiteten Testdaten. Ein weiterer Originalhandler enthält einen fehlenden
Controlverweis. Umfang und verbleibende Grenzen:
[DAO-Cursor und Formularsuche](docs/DAO_CURSOR_RUNTIME.md).

Kombofelder unterstützen zusätzlich native Wertlisten und Feldlisten aus
Abfragemetadaten, einschließlich mehrspaltiger Auswahl und gebundener Speicherung:
[dynamische Quelltypen](docs/DYNAMIC_COMBO_RUNTIME.md).

`Eval` und `Application.Eval` führen begrenzte native Access-Ausdrücke und
öffentliche VBA-Funktionen aus, einschließlich echter Formularverweise:
[Eval-Laufzeit und Grenzen](docs/EVAL_RUNTIME.md).

Öffentliche Standardmodulfunktionen können außerdem in tatsächlichen nativen
SQL-SELECTs und UPDATE-Berechnungen ausgeführt werden, mit geteiltem Modul-/Static-
Zustand. Deklarierte Geld-/Datumsrückgaben werden in unterstützten Projektionen
typisiert und gegen veraltete Funktionsdeklarationen geprüft. GUI-/DAO-Kontext,
allgemeiner Objekt-/Varianttransport und vollständige Jet-Typweitergabe bleiben offen:
[SQL-/VBA-Ausführung und Grenzen](docs/NATIVE_VBA_SQL.md).


Native Datensatzauswahl und begrenztes Löschen verwenden echte Transaktionen,
Bestätigungs-/Abbruchereignisse und `AllowDeletions`. Ein unveränderter
Original-Löschhandler ist auf synthetischen Zeilen direkt und per Button geprüft:
[Datensatz-Löschlaufzeit](docs/RECORD_DELETE_RUNTIME.md).

Der Import erhält AutoNumber-Markierungen und verwendet für einzelne
INTEGER-Primärschlüssel native Generierung ohne Wiederverwendung gelöschter
Höchstwerte. Auch die fünf Zähler ohne Primärschlüssel haben eine native
Generierung. Genaue Grenzen und verbleibende Jet-Abweichungen:
[AutoNumber-Import](docs/AUTONUMBER_IMPORT.md).

Begrenztes natives Datensatz-Kopieren/Anfügen verwendet die echte Zwischenablage
und erzeugt neue Schlüssel nur für tatsächliche AutoNumber-Felder. Der originale
Kopierhandler behandelt seinen manuellen Schlüsselfehler im Test:
[Clipboard-Laufzeit und Grenzen](docs/RECORD_CLIPBOARD_RUNTIME.md).

Eine `.deb`-Datei lässt sich nach dem Build erzeugen mit:

```bash
cpack --config build/CPackConfig.cmake -G DEB
./scripts/check-deb.sh ./linaccess-0.7.0-Linux.deb build/linaccess
```

Die Datei `.github/workflows/build-linux.yml` prüft Ubuntu 24.04 sowie Debian
12/13 und installiert die erzeugten Pakete. Die zuletzt bestätigten Läufe sind
erfolgreich; die frühere Kontozahlungs-/Ausgabenlimit-Sperre trat dabei nicht
mehr auf. Commitbezogene Nachweise stehen im [Fortschrittsprotokoll](docs/PROGRESS.md).

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

Selbst angelegte Berichte öffnen jetzt dieselbe native Report-Laufzeit mit SQL,
VBA-Ereignissen, Seitenansicht und PDF. Der automatische Spaltenentwurf bleibt
begrenzt; siehe [eigene Berichte](docs/NATIVE_STORED_REPORTS.md).

Die native klassische Oberfläche ist unter **Ansicht → Klassische Oberfläche**
umschaltbar; die Auswahl wird gespeichert. Sie bildet eine klassische Qt-Grundlage,
keine nachgewiesene pixelgenaue Access-2003-Oberfläche. Details: [Darstellung](docs/CLASSIC_INTERFACE.md).

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
  `For/Next`, `For Each`, `Exit For`, `Exit Sub/Function`, `Me!Feld = ...`, `&`-Verkettung,
  einige Standardfunktionen (`IsNull`, `Nz`, `Len`, `CStr`, `CLng`,
  `CDbl`, `UCase`, `LCase`, `Abs`), `DoCmd.RunCommand acCmdSaveRecord`,
  `DoCmd.Maximize` und `DoCmd.GoToRecord , , acLast`.
  Das Ereignis `Form_Open(Cancel As Integer)` wird unterstützt und kann
  durch Setzen von `Cancel` die Formularöffnung abbrechen.
  **Keine vollständige Microsoft-VBA-Runtime; DAO nur als Teilmenge, COM nicht implementiert.**
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
  native `Filter`/`FilterOn`-Eigenschaften mit Filterleiste ([Nachweise](docs/NATIVE_FILTERS.md)),
  parameterisierte Sub-/Function-Aufrufe, ByVal/ByRef, lokale Fehlerhandler
  mit Labels/Resume und VBA-Zeilenfortsetzung `_`. Gebundene Update-/Dirty-/
  Navigations-/Schließereignisse samt Cancel werden getestet; siehe
  [Ereignisse](docs/FORM_EVENTS.md).
  **SQL-Aktionen wirken direkt auf die geöffnete SQLite-Projektkopie**, auch
  bei `DELETE`, `UPDATE`, `DROP TABLE` und `INSERT`. Es erscheint keine
  LinAccess-spezifische Sicherheitsabfrage vor jeder VBA-Anweisung.
- Begrenzte native Date-Variants mit typisierten VBA-/DAO-/Formularwerten,
  CDate, DateSerial, TimeSerial und DateAdd sowie geprüften Datumsprojektionen
  in SQL; regionale Formate und vollständige Datumsparität bleiben offen.
  Siehe [Datumslaufzeit](docs/DATE_RUNTIME.md).
- Importierte Access-Abfragen laufen im Navigator über einen begrenzten Jet-SQL-
  Adapter mit Parameterdialogen, DAO-QueryDefs, TOP-Sortiergleichständen, eager
  IIf und begrenzter nativer Datumsformatierung für SQL/VBA-Format$. Gespeicherte Abfragen können schreibgeschützte Formular-Datenquellen sein.
  78/100 Kassenabfragen werden ausgeführt (73 ohne zusätzliche Parameter, fünf mit
  Datenparametern und geprüften positiven Ergebnissen), bisher ohne Access-Referenzergebnis;
  SQL-Verweise auf geöffnete Forms/Formulare und Reports/Berichte lesen echte
  Controlwerte, auch in gespeicherten Abfragen; [Umfang und Tests](docs/SQL_OBJECT_VALUES.md).
  siehe [Abfrageadapter](docs/QUERY_ADAPTER.md). Der native Berichtspfad bietet
  begrenzte Mehrdatensatz-Paginierung, PDF-Ausgabe, Abschnittsseitenwechsel und
  Aggregate; vollständige Access-Berichts-/Druckkompatibilität fehlt weiterhin.
  Grenzen: [Aggregate](docs/REPORT_AGGREGATES.md),
  [Textwachstum](docs/REPORT_CANGROW.md) und [Fußbereichsentwurf](docs/NATIVE_REPORT_FOOTER_DESIGN.md).
  Drei originale Open-/Close-Paare laufen
  automatisch. Sechs von neun Originalquellen binden 54 nicht-NULL-Textfelder;
  siehe [Berichtsereignisse](docs/REPORT_EVENTS.md). Der native CommandBars-
  Adapter ist in [Druckvorschau](docs/PRINT_PREVIEW_ADAPTER.md) beschrieben.
- Importierte Formulare bieten native `Forms`-/`Controls`-Sammlungen einschließlich
  `Me.Controls`, tatsächlicher formulierübergreifender Datenänderungen und
  geprüfter Qt-Objektlebensdauer. Die Implementierung bleibt eine Teilmenge;
  siehe [Formular-Objektmodell](docs/FORM_OBJECT_MODEL.md).
- Für ein per Windows-Bridge importiertes Formular gibt es die Aktion
  **Als natives Formular übernehmen (Teilkonvertierung)**. Sie liest
  `RecordSource` und gebundene `ControlSource`-Felder aus `SaveAsText`, gleicht
  sie mit einer bereits in SQLite importierten Tabelle ab und verknüpft das
  gleichnamige importierte `Form_...`-VBA-Modul mit dem neuen Formular.
  Die ältere **Teilkonvertierung** übernimmt weiterhin nicht das
  Access-Layout; hierfür die neue gesonderte Access-Layoutansicht benutzen.
- Native Formulare können auch auf importierten SQLite-Tabellen ohne `id`-Spalte
  arbeiten (SQLite-`rowid`-Fallback).
  Der Formulardesigner unterstützt außerdem gespeicherte Abfragen und einzeilige,
  parameterfreie SELECT-Datenquellen als schreibgeschützte Formulare; Feldlisten
  stammen aus vorbereiteten Abfragen. [Umfang](docs/NATIVE_QUERY_FORM_DESIGN.md).

**Nicht implementiert:** vollständige Original-VBA-Ausführung, wirklich
pixelidentische Form-/Report-Rekonstruktion, eingebettete Access-Makros,
vollständige DAO-/ADO-Unterstützung, ActiveX, Windows-COM, Jet/ACE-SQL und 1:1-Konvertierung.

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
der weiterhin unvollständigen VBA-/Objekt-/SQL-Semantik. Elf originale
Tastatur-Schaltflächen samt Parameterhelfer sowie ein DAO-Einfügehandler
werden inzwischen separat gegen Testkopien ausgeführt; siehe
[DAO-Laufzeit](docs/DAO_RUNTIME.md).
In Version 0.7 wurden alle **47 realen Access-2003-Formdefinitionen** mit
dem Layout-Parser geprüft: Nach Korrektur der Blockzuordnung werden 879 Controls
und 86 sichtbare CommandButtons erkannt. Unterstützte Widgets aller 47 Formulare
werden in der Strukturregression instanziiert. Ein echter Qt-Klick mit VBA-`CurrentDb.Execute` wird gegen eine
separate SQLite-Testtabelle ausgeführt. Das belegt **nicht**, dass die
Original-Kassenaktionen bereits laufen.
Auch der tabellarische MDB-Import wurde direkt an der originalen Kassen-MDB
getestet: **29 von 29 Tabellen** konnten als typisierte SQLite-Kopien angelegt
werden. Die vorhandene `Foto`-BLOB-Spalte enthielt im Test keine nicht-NULL-
Werte. Die Regression prüft exportierte Schlüssel/Indizes, Beziehungen und
sämtliche Zellen einschließlich NULL. Exakte Jet-Typen, Standardwerte und
OLE-Nutzdaten sind weiterhin nicht allgemein erhalten.
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

## Native DAO-Teilmenge

Tabellenrecordsets können jetzt per VBA geöffnet, durchlaufen und mit
`Edit`/`AddNew`/`Update`/`Delete` verändert werden. Ein echter Qt-Klick führt den
unveränderten Kassenhandler `SF_Ende_Kassa_Click` gegen eine SQLite-Testkopie aus.
Das belegt einen Originalhandler, keine vollständige Kassen-/VBA-Kompatibilität.
Unterstützung, Tests und offene Grenzen stehen unter [DAO-Laufzeit](docs/DAO_RUNTIME.md).
`Recordset.GetRows(Anzahl)` liefert bei ausdrücklich positiver Zeilenzahl echte
zweidimensionale Variant-Arrays mit typisierten Feldwerten; VBA-Indizierung,
Dimensionsgrenzen und mehrdimensionales `ReDim Preserve` sind nativ angebunden.
`ReDim ... As Typ` prüft außerdem primitive Elementtypen in vorhandenen
Variant-Variablen; `Set` erhält native Objektverweise in Array-Zellen, und
`As Object` initialisiert sie mit `Nothing`. Feste/typisierte Array-Deklarationen
sind noch nicht vollständig.
Optionale Defaults und vollständige DAO-/Array-Parität bleiben offen;
siehe [Datenübergabe](docs/DAO_GETROWS_RUNTIME.md).


## Experimentelle VBA-Ausführung

Dieser Code wird vom neuen C++-Teilinterpreter in einem LinAccess-Formular
unterstützt:

```vb
Private Sub Form_Load()
    MsgBox "Formular geöffnet"
End Sub

Private Sub btnSpeichern_Click()
    If IsNull(Me!Nachname) Then
        MsgBox "Bitte Nachnamen eingeben"
        Exit Sub
    End If
    DoCmd.RunCommand acCmdSaveRecord
End Sub
```

Die native Laufzeit führt die dokumentierte Teilmenge mit echten Formular- und
DAO-Effekten aus. Normale SQL-Aktionen benötigen keine zusätzliche Zustimmung
pro Aktion. Nicht unterstützte Syntax wird vor dem Prozedurstart gemeldet;
Laufzeitfehler können nach bereits ausgeführten Anweisungen auftreten.
**Generische COM-/ActiveX-/Windows-API-Ausführung fehlt; DAO ist eine Teilmenge.**
Der [Kalenderadapter](docs/CALENDAR_RUNTIME.md) und
[With, Tastatureingaben und SELECT INTO](docs/WITH_KEYBOARD_RUNTIME.md) sind mit
Originalartefakten getestet. Positionsparameter, begrenztes ByRef/ByVal und
Label-Fehlerhandler sind implementiert. Arrays, Select Case, While/Wend,
allgemeine Klassen, vollständige Typkonvertierung und ein vollständiger
VBA-Parser fehlen weiterhin. Die gesamte Kassen-Verkaufskette läuft noch nicht.

Ein zusätzlich unterstütztes Beispiel:

```vb
Public Function Summe() As Long
    Dim i As Long
    Dim wert As Long
    wert = 0
    For i = 1 To 4
        wert = wert + i
    Next i
    Summe = wert
End Function
```

Diese Funktion liefert im eingeschränkten Interpreter `10` zurück.

## Build-Status

Letzter vollständiger lokaler Lauf (Debian 13 / GCC 14 / Qt 6.8,418719d):
**48/48 Tests**, einschließlich vollständigem 29-Tabellen-Zellvergleich und
originalem Kassen-DAO-Buttonereignis. Separate unveränderliche e2d64c9-Kopien
bestehen lokal Debian12/Qt6.4 und Ubuntu24.04/Qt6.4 ebenfalls mit jeweils48/48,
Paketinstallation, Build-ID, SQLite und Fixture-Hashes. Der folgende kleine
OpenRecordset-Modus-Fix ist gezielt geprüft; exakte Grenzen und neueste Belege
stehen in [PROGRESS.md](docs/PROGRESS.md). GitHub startet die aktuellen Jobs wegen
Kontozahlungs-/Ausgabenlimits nicht. Lokale Containerprüfungen sind keine neuen
GitHub-CI-Erfolge und kein Windows-Access-/Gesamtparitätsnachweis.

Früherer Referenzstand:

**Verifiziert:** Ubuntu 24.04 / GCC 13 / Qt 6.4 / CMake 3.28:

- Native GUI und C++-Kern wurden erfolgreich kompiliert.
- `ctest` besteht mit 8 Tests (bei Vorliegen beider privaten MDB-/Bridge-Testdateien): VBA-/CSV-Kern, echte JET4-MDB mit typisiertem
  Import, echter MDB-Katalogimport über eine Qt-Offscreen-Oberfläche sowie
  Bridge-Import und Manipulationsschutz und realer Bridge-Export mit
  Ausführung von `Form_Open` und komplexer Bridge-Import mit 192 Quellen
  **ohne** Ausführung von Original-Kassen-VBA (Qt-Offscreen),
  47 reale Layouts/Qt-Eventausführung und **29/29** kopierte Kassen-Tabellen.
- Die App startet mit `QT_QPA_PLATFORM=offscreen` ohne Absturz.
- Ein installierbares Paket `linaccess-0.7.0-Linux.deb` wurde gebaut und sein
  Inhalt mit `dpkg-deb` geprüft. Die benötigten Qt-Bibliotheken werden über
  Paketabhängigkeiten installiert.

**Noch nicht verifiziert:** interaktive Desktop-Bedienung auf einem echten
Linux-Desktop, reale `.accdb`-Dateien, reproduzierbarer Windows-
PowerShell-Bridge-Build an allen Access-Versionen, volle VBA-Kompatibilität.

## Debian- und Ubuntu-Pakete

Der Quellcode baut mit C++20, CMake und Qt 6.4+ auf beiden Systemfamilien.
GitHub Actions enthält getrennte Jobs für **Ubuntu 24.04**, **Debian 12**
und **Debian 13**: jeder Job erstellt ein auf seiner Distribution gebautes
`.deb`-Paket. Die aktuellen bestätigten Läufe sind im Fortschrittsprotokoll verlinkt.

Ein unter Ubuntu 24.04 erzeugtes `.deb` ist wegen der Qt-6-`t64`-Paketnamen
ausdrücklich **nicht** als Debian-12-Paket freigegeben. Für Debian ist ein
separat auf Debian gebautes Paket bzw. Selbstkompilierung notwendig.

## Architektur

```text
Qt Widgets GUI
  ├─ SQLite-Projekt / Tabellen / Abfragen
  ├─ Native Formulare und Berichte (neue LinAccess-Objekte)
  ├─ Access Inspector → MDBTools (Typkopie / Objektkatalog / Jet-SQL)
  ├─ Access Bridge → Quellobjektkatalog und SHA-256-Prüfung
  └─ VBA-Modulkatalog → Analyse + experimentelle Teilinterpretation
                          └─ später vollständige VBA-Laufzeit + Access-Objektmodell
```

Neue `.linaccess`-Dateien sind SQLite-Datenbanken mit reservierten
Metadatentabellen `_la_objects`, `_la_import_log` und `_la_access_artifacts`.

## Nächste, in dieser Reihenfolge geplante Schritte

1. Weitere echte Testdateien (insbesondere ACCDB) und interaktive UI-Tests
2. Den typisierten Import um Schlüssel, Beziehungen, Indizes, Standardwerte
   und Jet/ACE-Abfragekonversion erweitern
3. Windows-Export-Brücke mit weiteren Access-Versionen testen und die bereits
   begonnene `SaveAsText`-Teilkonvertierung auf mehr Formelemente erweitern
4. Drag-and-drop-Formular-/Berichtdesigner mit Ereignisbindung
5. VBA-Laufzeit systematisch erweitern (Parameter, Prozeduraufrufe, weitere
   Schleifen/Fehlerbehandlung und vor allem DAO/Access-Objektmodell)
6. Systematische Kompatibilitätstests anhand realer `.accdb`-/`.mdb`-Projekte
7. Linux-Binaries und Distributionspakete (AppImage, `.deb`)

Die Windows-Export-Brücke wäre ein optionales **Importwerkzeug**, nicht die
Linux-Laufzeit. Vollständige VBA-/Access-Kompatibilität ist auch damit nicht
automatisch erreicht; besonders COM, ActiveX und Windows-API-Aufrufe bleiben
problematisch.

Ein komplexer originaler `Form_Load` läuft jetzt mit echten eingebetteten
Unterformularen; 15 dynamische Tags sind gegen die tatsächlichen Datensätze
geprüft. Berechnete Grundausdrücke, Error-Variants und ComboBox-Spalten sind
teilweise unterstützt. Ein anfänglicher Null-Konvertierungsfehler und vollständige
Kassenabläufe bleiben offen: [Details und Regressionen](docs/NATIVE_FORM_RUNTIME.md).


[OpenForm mit WhereCondition](docs/OPEN_FORM_RUNTIME.md) lädt im begrenzten
nativen Laufzeitpfad echte gefilterte Datensätze. Ein unveränderter Kassen-Handler
öffnet sein originales Zielformular. Dessen zuvor roter OptionGroup-Test wird
mit nativen Auswahlknöpfen und numerischem Gruppenwert erfüllt. Weitere
Access-Fenstermodi und vollständige Ereignissemantik bleiben offen.


[Unabhängige Dokumente und Quit](docs/DOCUMENT_LIFETIME_RUNTIME.md) erhalten
OpenForm-Ziele nach dem Schließen und Zerstören des Aufrufers. Der unveränderte
originale Kassenöffner und seine Zielereignisse werden ausgeführt; Projektwechsel
und Quit prüfen native Unload-Abbrüche. [Dynamische ComboBox-Eigenschaften](docs/DYNAMIC_COMBO_RUNTIME.md)
verwenden echte mehrspaltige Qt-Modelle mit getrennten Anzeige-/Bindewerten und
Spaltenbreiten. Das ist keine vollständig geprüfte ursprüngliche Verkaufskette.

[Native Tastaturbefehle und Tastaturereignisse](docs/KEYBOARD_RUNTIME.md) umfassen
einen begrenzten, tatsächlich ausgeführten Pfad für Escape, Umschalt+Enter,
F9/Umschalt+F9, F2, Enter sowie Strg+W/Strg+F4. KeyPreview und gebundene VBA-
KeyDown/KeyPress/KeyUp-Ereignisse können Eingaben abbrechen oder ASCII-Zeichen
ersetzen. Vollständige Access-2003-Oberflächen- und Tastaturparität ist weiterhin offen.

[Currency in einfachen UPDATE-Anweisungen](docs/CURRENCY_UPDATE_RUNTIME.md)
prüft Zielspalten vor dem Schreiben und erhält vier Dezimalstellen. Native
DAO-/VBA-/Qt-Tests prüfen auch den vollständigen Rollback bei Konvertierungsfehlern;
komplexe Aktionsabfragen und vollständige Jet-Semantik bleiben offen.

[Currency in einfachen INSERT-VALUES-Anweisungen](docs/CURRENCY_INSERT_RUNTIME.md)
konvertiert Geldbeträge vor dem Einfügen. DAO-/VBA-/Qt-Regressionen prüfen
Spaltenreihenfolge, Parameter, Rundung und Fehler ohne teilweise eingefügte Zeile.
INSERT-SELECT und vollständige Jet-Semantik bleiben offen.

[Form.CurrentRecord](docs/CURRENT_RECORD_RUNTIME.md) liefert für gebundene native
Formulare die aktuelle Position als schreibgeschützten VBA-Long, auch nach
Navigation und Filterung. Entwurfsansicht und vollständige Windows-Semantik
bleiben offen.

[Native Datumsfunktionen](docs/CALENDAR_FUNCTIONS_RUNTIME.md) umfassen DateDiff,
Year/Month/Day, Hour/Minute/Second und Weekday für eindeutige ISO-Werte in VBA
und SQL. Date-Variants, regionale Datumskonvertierung und vollständige
Kalendersemantik bleiben offen.

[Native Formularsortierung](docs/FORM_ORDER_RUNTIME.md) setzt OrderBy/OrderByOn
gegen echte Tabellen- und Abfragequellen um. Tabellenzeilen behalten beim
Umsortieren ihre Identität; vollständige Jet-Sortier- und Kollationssemantik
bleibt offen.

[Native Tabellen- und SQL-Editoren](docs/APP_EDITORS_RUNTIME.md) verwenden jetzt
geprüfte Access-Feldtypen, vierstellige Currency-Konvertierung und vollständige
lange Memo-Texte. Der SQL-Editor führt Access-Abfragen mit Parametern und
DAO-Aktionsabfragen aus. Vollständige Datenblatt-, Designer- und Jet-Parität
bleibt offen.


## Weitere native Laufzeit- und Datenblattfunktionen (2026-11-04)

- [Native Dictionary-Objekte und typisierte Bindungen](docs/SCRIPTING_DICTIONARY_RUNTIME.md),
  [qualifizierte DAO-Bindungen](docs/DAO_OBJECT_BINDINGS.md), explizite
  [Static-Variablen](docs/STATIC_LOCAL_RUNTIME.md) und
  [Standardwert-Zuweisungen mit Let](docs/OBJECT_DEFAULT_ASSIGNMENT_RUNTIME.md)
  nutzen echte native Werte und Referenzen.
- [Strg+H: ganze Textfeldwerte ersetzen](docs/TABLE_REPLACE_RUNTIME.md) nutzt
  den vorhandenen Entwurf-/Save-/SQL-Pfad.
- [Esc: aktuelle Datensatzänderungen verwerfen](docs/TABLE_RECORD_UNDO_RUNTIME.md)
  erhält andere Entwürfe und macht gespeicherte SQL-Daten nicht rückgängig.
- [For Each über Byte-Arrays](docs/VARIANT_ARRAY_RUNTIME.md) liefert echte
  Byte-Werte, auch bei mehrdimensionalen Arrays.

Die verlinkten Dokumente nennen die jeweiligen Grenzen. Diese Ergänzungen
belegen weiterhin keine vollständige Access-2003-Kompatibilität.

[Geprüfte Datenblatt-Navigation](docs/TABLE_KEYBOARD_RUNTIME.md) umfasst außerdem
begrenzte tatsächliche Speicher-/Navigationspfade für Pfeil hoch/runter,
Tab/Umschalt+Tab zwischen Nachbarzeilen und Strg+Pos1/Strg+Ende. Die jeweiligen
Kontext-, Zeilen-, Editor- und Tastaturgrenzen bleiben ausdrücklich dokumentiert.

[Native Finanzberechnungen](docs/NATIVE_FINANCIAL_FUNCTIONS.md) umfassen begrenztes
NPer und DDB in VBA/SQL sowie NPV/MIRR mit tatsächlichen eindimensionalen
Double-Arrays in VBA. Typ-, Perioden- und Fehlergrenzen bleiben ausdrücklich
dokumentiert; ein Access-Referenzlauf steht weiterhin aus.

[Berechnete Abfragespalten](docs/QUERY_COLUMN_LABELS.md) zeigen für unterstützte
Ausdrücke lesbare Quellbezeichnungen statt interner SQL-Hilfsnamen. Eigene
Spaltenaliase bleiben erhalten; Access-spezifische Expr1000-Regeln sind weiterhin
nicht durch einen Referenzlauf bestätigt.

[Native FileSystemObject-Adapter](docs/NATIVE_FILE_SYSTEM_OBJECT.md) arbeiten mit
tatsächlichen Dateien und Ordnern: Lesen von Metadaten, Auflisten, Anlegen,
Kopieren, Umbenennen, Verschieben und Löschen innerhalb dokumentierter Grenzen.
[Native TextStreams](docs/NATIVE_TEXT_STREAM.md) lesen und schreiben ASCII oder
ausdrücklich gewähltes UTF-16LE mit BOM; Anhängen und Surrogatgrenzen sind
abgegrenzt. Vollständige COM-, Windows-Pfad-, Kodierungs- und Freigabeparität
bleibt offen.

[Strg+P für native Dokumente](docs/FORM_PRINT_SHORTCUT_RUNTIME.md) nutzt den
echten Druckdialog: Formulare geben das aktuelle Layout aus, Berichte ihren
unterstützten Seitenumfang. Formular-KeyDown kann den Aufruf abbrechen;
vollständige Access-Druck- und Tastaturparität bleibt offen.

[Tabellen-/Abfragedatenblatt drucken und als PDF ausgeben](docs/DATASHEET_PRINT_RUNTIME.md)
verwendet echte sichtbare, gefilterte/sortierte Daten und atomare PDF-Speicherung.
Ungespeicherte Entwürfe und nicht unterstützte Inhalte werden ausdrücklich gemeldet;
vollständige Access-Druck-/Formatparität bleibt offen.
