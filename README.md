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
Boost.Multiprecision ab1.74 wird nur als Header beim Bauen benötigt; eine
zusätzliche Boost-Laufzeitbibliothek ist nicht erforderlich.

```bash
sudo apt update
sudo apt install build-essential libboost-dev cmake qt6-base-dev libqt6sql6-sqlite mdbtools mdbtools-dev pkg-config dpkg-dev file

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

Neue Access-Importe erhalten AutoNumber-Markierungen und gespeicherte
Quellzählerstände auch nach gelöschten Zeilen. Neue Tabellen im Designer verwenden
dieselbe verwaltete 32-Bit-Vergabe. DAO reserviert die Nummer bei `AddNew`;
Abbruch und Workspace-Rollback geben sie nicht wieder frei. Der nachgewiesene
signed-Long-Übergang ins Negative ist implementiert. Genaue Grenzen und verbleibende
Jet-Abweichungen:
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
12/13 und installiert die erzeugten Pakete. Nachweise erfolgreicher Läufe müssen
zum jeweils geprüften Commit passen; commitbezogene Ergebnisse und Grenzen
stehen im [Fortschrittsprotokoll](docs/PROGRESS.md).

`mdbtools` ist optional, aber zum Lesen von `.mdb` und versuchsweise `.accdb`
im Access-Inspektor erforderlich. Die Unterstützung ist je nach MDBTools-Version
und Dateiformat unterschiedlich. Access-Dateien werden dabei nie überschrieben.

Nur die C++-Parser-Tests (ohne Qt) mit GCC:

```bash
g++ -std=c++20 -Wall -Wextra -pedantic -Isrc \
    tests/core_tests.cpp src/vba_analyzer.cpp src/vba_runtime.cpp \
    src/vba_expression.cpp src/vba_currency.cpp src/vba_decimal.cpp \
    src/vba_round.cpp src/vba_calendar.cpp src/vba_array.cpp src/csv.cpp -o core_tests
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
- Auf geschlossenen Tabellen über **Ansicht → Tabellenfelder bearbeiten** normale Felder mit Typ, Textgröße und Standardwert ergänzen oder nach Bestätigung löschen; [Umfang und Grenzen](docs/DAO_TABLEDEF_FIELDS.md)
- Auf ausgewählten Tabellen über **Ansicht → Tabellenindizes bearbeiten** normale/Unique-Indizes mit geordneten ASC-/DESC-Feldern anlegen und löschen; [Umfang und Grenzen](docs/DAO_INDEXES_RUNTIME.md)
- Über **Ansicht → Beziehungen bearbeiten** numerische Beziehungen mit echter Integritätsprüfung/Kaskaden oder ausdrücklich ohne Integritätsprüfung verwalten; [Umfang und Grenzen](docs/DAO_RELATIONS_RUNTIME.md)
- Gespeicherte SQL-SELECT-Abfragen mit [visuellem Datenquellenentwurf](docs/VISUAL_QUERY_DESIGNER.md): echte Tabellen/gespeicherte Auswahlabfragen, Felder und Aliase, Kriterien, Parameter, berechnete Felder, TOP mit Gleichständen, Gruppierung/Aggregate und begrenzte INNER-/LEFT-/RIGHT-Verknüpfungen mit Quellenaliasen, Selbstverknüpfungen und zusammengesetzten Schlüsseln. Tabellenerstellungs-/Anfügeabfragen sowie begrenzte Aktualisierungs-/Löschabfragen für einzelne Tabellen werden ebenfalls visuell erstellt und erst über Ausführen gestartet. Derselbe Entwurf erstellt Formular- und Berichtsquellen; Abfrageformulare bleiben im bestehenden schreibgeschützten Umfang.
- Native Aktionsmakros erstellen, speichern und über `DoCmd.RunMacro` ausführen;
  Umfang und Grenzen: [Makrolaufzeit](docs/NATIVE_MACROS_RUNTIME.md).
  [Formularereignisse zuordnen](docs/NATIVE_FORM_EVENT_DESIGNER.md) geht im nativen Designer.
- [Eigene Standardmodule](docs/NATIVE_STANDARD_MODULES_RUNTIME.md) erstellen und bearbeiten;
  unterstützte öffentliche Funktionen aus VBA und SQL verwenden;
  [Application.Run](docs/NATIVE_APPLICATION_RUN.md) ruft öffentliche Funktionen und Subs dynamisch auf.
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
- **Freie Formularerstellung:** Ungebundene und leere Formulare sowie eigene
  Textfelder, Labels, Schaltflächen, Checkboxen, Kombinations- und Listenfelder
  mit tatsächlicher VBA-/SQL-Laufzeit und Designer-Undo.
  [Umfang und Grenzen](docs/FREE_FORM_CONTROLS.md).

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
Der lesende [Beziehungskatalog](docs/DAO_RELATIONS_RUNTIME.md) erschließt originale
Beziehungsnamen und geordnete Primär-/Fremdfelder aus realen Importen; vollständige
DAO-Katalogmutation und Jet-Parität bleiben offen.
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

Die Datenblatt-Seiteneinrichtung verwendet echte Papier-/Randgeometrie für Vorschau, Druck und PDF. Einstellungen gelten je geöffnetem Dokument und lassen sich ausdrücklich im nativen Projekt speichern; Details und Grenzen stehen in der verlinkten Druckdokumentation.

[Begrenzter Jet-Rohbinärimport](docs/JET_BINARY_IMPORT.md) kopiert ausdrücklich hexexportierte Binärwerte als echte SQLite-BLOBs. OLE-Anzeige, ACE-Anhänge und Originalformat-Roundtrip bleiben offen; Java/Jackcess dienen ausschließlich der optionalen technischen Referenzerzeugung.

Originale Indexnamen und Feldrichtungen werden bei neuen MDB-Importen über einen
lesenden nativen Helfer geprüft; Umfang und Grenzen: [Indexmetadaten](docs/SOURCE_INDEX_METADATA.md).

Der lesende [DAO-Indexkatalog](docs/DAO_INDEXES_RUNTIME.md) verwendet diese Originalnamen
und prüft Felder, Richtungen und Schlüssel gegen den vollständigen nativen Indexbestand.

Bei tatsächlichen MDB-Importen werden bestätigte endliche Single-Werte direkt aus
den ursprünglichen Bits übernommen: [Präzisionsumfang und Grenzen](docs/SOURCE_SINGLE_PRECISION.md).

Auch ursprüngliche Double-Bitmuster werden bei tatsächlichen MDB-Kopien erhalten;
Umfang, unabhängiger Originaldatenvergleich und Grenzen: [Double-Präzision](docs/SOURCE_DOUBLE_PRECISION.md).

Der native Tabellenfeld-Editor kann bestehende Felder anlegen/löschen, Textgröße,
Pflichtfeld und Literalstandard gemeinsam ändern sowie normale Felder umbenennen.
Datenwerte bleiben bei Eigenschaftsänderungen erhalten; unzulässige Änderungen
werden zurückgerollt. Umbenennungen erfordern das Anpassen gespeicherter
Access-Abfragen und Formularbindungen. [Umfang und Grenzen](docs/DAO_TABLEDEF_FIELDS.md).

Eigene native Tabellen lassen sich unter Ansicht → Tabelle umbenennen auch mit
AutoNumber, Indizes und Beziehungen tatsächlich umbenennen; Counterstände und
gespeicherte Seiteneinstellungen bleiben erhalten. Gespeicherte Access-Abfragen
und Formularbindungen müssen anschließend auf den neuen Namen angepasst werden.
[Umfang und Grenzen](docs/DAO_TABLEDEF_FIELDS.md).

Neue native Tabellen und Felder unterstützen auch Single (32-Bit-Gleitkomma),
einschließlich quantisierter Datenblatt-/DAO-Werte und Literalstandardwerte.
[Unterstützte Schreibpfade und Grenzen](docs/DAO_TABLEDEF_FIELDS.md).

Gespeicherte native Makros lassen sich im Datenbankfenster auswählen und über
**Ansicht → Ausgewähltes Makro ausführen** starten; [Umfang](docs/NATIVE_MACROS_RUNTIME.md).

Native VBA-Debugger im Modul-Editor: F9-Haltepunkte, F5/F8 und echte lokale Werte/Aufrufrahmen sowie Fehlerfortsetzung; [Umfang und Grenzen](docs/NATIVE_VBA_DEBUGGER_RUNTIME.md).

Unter Extras ist der [Projektbericht zu Fähigkeiten und Ausführungsbeobachtungen](docs/PROJECT_CAPABILITY_REPORT.md) mit geschütztem JSON-Export verfügbar. Syntaxannahme, fehlender Kontext und tatsächliche Aufrufergebnisse bleiben getrennt.

Native [Boolean-Bedingungen und strukturierte If-Fortsetzung](docs/VBA_BOOLEAN_CONDITIONS.md) verwenden gemeinsame geprüfte Konvertierung; vollständige VBA-/Access-Parität bleibt offen.


Typisierte Function-/Property-Get-Arrayrückgaben sind implementiert: geschlossene
native Elementtypen, vorhandene Array-Typ-/Rang-/Grenz-/Wertkopiesemantik und echte
Leseindizierung mit einmaliger Erzeugerausführung. Auch Variant-Arrayergebnisse,
mehrdimensionale Getter und separate Indexqueue nach parametrisiertem Aufruf.
Bekannte typisierte Array-SQL-Funktionen vor Bodyeffekten3251; keine erfundenen
SQLite-Skalarwerte. Reale Formularfassade und gemeinsamer SQL/Modul-Zustandszähler
beweisen Lesezugriff bzw. fehlende SQL-Seiteneffekte. Details NATIVE_VBA_PROPERTIES.md.
Anwendung gebaut; erstes bestehendes Kern-/GUI-Bündel2/2 (16,91s) bestanden, keine
neuen CTests/Original-/Windows-/Gesamtsuitennachweise. Vorstufe1313196 privat
veröffentlicht und frischeDEB1313196 samt Build-ID/Payload/Permissions und gepinnten
SQLite/MDBTools geprüft. Aktuelle Array-Erweiterung liegt nach diesem Paket.
Katalog456/386/70 unverändert; Array-Ergebniszuweisungen/Klassen-/UDT-/COM-Typen
bleiben begrenzt. Direkt weiter: For Each über tatsächliche Array-Funktions-/Getter-
ergebnisse mit einmaliger Ausdrucksauswertung und erhaltenen Named-Root-Guards.


For Each akzeptiert echte Array-Ausdrücke einschließlich typisierter/Variant-
Funktions- und Getterergebnisse. Quelle nach MS-VBAL5.4.2.4 genau einmal auswerten,
Ergebniswert während Schleife halten; bisherige intrinsische Namensliste entfällt.
Benannte Wurzeln behalten vorhandene Live-Elementänderungen und Austausch-/Resize-
Guards. Erstes bestehendes Kern-/GUI-Bündel2/2 (17,83s) und Anwendung erfolgreich:
Erzeugerzähler, Matrix-Reihenfolge, Objektinhalte, ByRef-Input/Wertkopieunabhängigkeit,
verschachtelte Schleifen/Exit For. Details VARIANT_ARRAY_RUNTIME.md. Keine neuen
CTests/Original-/Windows-/Gesamtsuiten-/Paketnachweise; letzte geprüfteDEB1313196.
Arrayrückgabe-Vorstufe2ce6323 privat veröffentlicht/Remote-Baum verifiziert.
Katalog456/386/70 unverändert. Direkt weiter: native Def-Typdirektiven statt der
bisherigen pauschalen Verweigerung, mit konsistenter Typauflösung und Grenzen.


Native DefBool/Byte/Cur/Date/Dbl/Int/Lng/Obj/Sng/Str/Var-Typdirektiven sind integriert:
aktive Moduldeklarationen, ASCII-Einzelbuchstaben/auf- und absteigende Bereiche,
A-Z als universelle Vorgabe einschließlich Unicode-Identifier. Dieselbe Auflösung
für Modul-/Lokalskalare und Arrays ohne As, Parameter, Function/Get-Rückgaben,
Setterwerte und gewöhnliche/statische implizite Locals; explizites As gewinnt,
Const bleibt aus Ausdruck typisiert. Option Explicit bleibt wirksam. Fehlerhafte/
überlappende/64-Bit-/prozedurinterne Direktiven vor Bodyeffekten abweisen. Implizite
Currency-For-Zähler exakt und neue ReDim-Ziele bei Def-Vorgabe typisiert; keine
Behauptung vollständigen Deklarations-Hoistings oder Suffix-/Klassenparität.
Details VBA_DEF_TYPES.md. Erste GUI-Prüfung erfolgreich, erster Kernblock fand
leere DefInt-Spezifikation; genau diesen Parserfehler behoben und fehlende Schleifen-/
ReDim-Bindung ergänzt. Final bestehendes Kern-/GUI-Bündel2/2 (17,68s; Kern4,33s),
Anwendung gebaut; alle7 Upload-/Fixturehashes unverändert. Keine neuen CTests,
Original-/Windows-/Gesamtsuitennachweise. Vorstufe a6b2429 privat veröffentlicht und
Remote-Baum verifiziert. Letzte geprüfteDEB1313196; Paket jetzt aktualisieren.
Katalog456/386/70 unverändert; ohne Zwischenabschluss weitere offene Umfänge umsetzen.


Echtes Tabellen-Datenblatt: Dialogbereich Alle sichtbaren Textfelder und atomare
Mehrfeld-Alle-ersetzen-Aktion implementiert. Vorbereitete Zellen über mehrere
Spalten/einen SQL-Transaktionsblock; späterer Feld-/Datensatzfehler rollt frühere
Schreibzugriffe zurück und entfernt nur eigene Ersetzungsentwürfe. Primärschlüssel,
generierte/ausgeblendete/numerische/Binärfelder ausschließen, Nullwerte erhalten,
generierte Werte normal ableiten. Aktuell-Ersetzen bleibt Entwurf am aktuellen
Textfeld und behält Allfeld-Suchbereich. Validierungszeitliche Quell-/Feld-/Sicht-
Änderungen vor Schreiben2501; sortierten aktuellen Schlüssel/Spalte wiederherstellen.
Details LITERAL_REPLACE_RUNTIME.md. Anwendung gebaut, erstes bestehendes GUI-/Modell-
Bündel2/2 (16,60s): reale Buttons/SQL, vier Feldwerte, später CHECK-Rollback, fremder
Entwurf, Scopewechsel und doppelte Spalten geprüft. Keine neuen CTests/Original-/
Windows-/Gesamtsuitennachweise. Def-Typ-Vorstufe cd0cd59 privat veröffentlicht/Remote-
Baum verifiziert; frischeDEBcd0cd59 samt Payload/Rechten/Build-ID und gepinnten
SQLite/MDBTools erfolgreich geprüft. Dieses Paket liegt vor Mehrfeld-Ersetzen.
Katalog456/386/70 unverändert; ohne Zwischenabschluss weitere offene Umfänge umsetzen.


Bearbeitbare vollständige Einzeltabellen-Abfragen erweitert: SELECT */alias.*,
explizite/implizite Aliasnamen, qualifizierte WHERE-Felder und bis10 unterschiedliche
physische ORDER-BY-Felder mit ASC/DESC. Nur geprüfte Quelltabelle/-felder; Alias-
Filterqualifikation kanonisch umsetzen, Literale erhalten. Nativer Tabellenmodell-
Sortierpfad verwendet geprüfte Index-/Richtungsmetadaten, escaped SQL und behält
sekundäre Reihenfolge bei Select/Save/Requery. Kopf zeigt erstes Feld, andere
Benutzersortierung ersetzt Quellsortierung; vorhandene Draft-/Editor-Guards bleiben.
Echte gespeicherte Abfrage schreibt Basisdatensätze, erhält Schlüssel bei Sortier-
bewegung und entfernt nicht mehr passende Zeilen aus Filteransicht. Keine pauschale
Join-/Projektions-/Parameter-/Nullschlüssel-/Shadow-Updatebarkeit. Details
OPEN_QUERY_RUNTIME.md. Anfangs bestehendes GUI/Modell2/2 (16,82s); anschließend
Mehrfeldsortierung und echte neue GUI-/Modellfälle komplettiert, fehlendes Modell-
Test-Include korrigiert und nur dessen Ziel neu gebaut. FinalGUI/Modell2/2 (16,59s,
Modell0,09s), Anwendung gebaut. Keine neuen CTests/Original-/Windows-/Gesamtsuiten-/
Paketnachweise. Mehrfeld-Ersetzungs-Vorstufe1c66d81 privat veröffentlicht/Baum
verifiziert; letzte geprüfteDEBcd0cd59 liegt vor beiden Datenblatt-Erweiterungen.
Katalog456/386/70 unverändert; F079/F373-Beschreibungen bestehenden Umfängen
zugeordnet. Ohne Zwischenabschluss weiter, danach Paket aktualisieren.


Native eigene Klassenmodule und echte Instanzen implementiert: Klasseneditor/
Explorer/kind=class, gespeicherte CLASS/VB_Name-Quelle, gemeinsamer kollisionsgeprüfter
Modulkatalog/AllModules mit echter Editor-IsLoaded-Erkennung, keine Registrierung
von Klassenmethoden als freie SQL-Funktionen. New erzeugt getrennte Feld-/Array-/
Static-Zustände und führt privaten parameterlosen Class_Initialize einmal aus.
Öffentliche Felder/Konstanten/Sub/Function/Get/Let/Set, Me-/Objektidentität,
Arrayrückgaben, native Default-Weiterleitung und echte private Aufruferscopes.
Reentrant fremde Prozeduren/Ausdruckshosts erben keine privaten Klassenrechte.
Kontextdienste verwenden echte Projekt-/DAO-Hosts ohne erfundenes Formular;
Instanzen überleben aufrufende Editor-/Formfenster im lebenden Projekt. Beobachtete
Quellrevision-/Lösch-/Projektänderungen invalidieren dauerhaft. Grenzen ausdrücklich:
Object/Variant-Referenzen statt benannter Custom-As-Typen, externe Argumente ByVal,
kein importierter generischer Klassenbau/As New/Implements/WithEvents/Default-
Instanzen/COM; Class_Terminate vor Konstruktion verweigert. Details NATIVE_VBA_CLASSES.md.
Erster Kern/DAO bestanden; Original fand Def-Vorprüfung gewöhnlicher Kontrollnamen,
GUI erreichte20s-Watchdog im parallelen Block. Def-Bindung auf echte typisierte Ziele
und anwendbare Def-Namen begrenzt, Fehler kontrolliert transportieren; New-Prüfung
verwechselt rohe Kontrollnamen nicht mit Konstruktoren. Private Reentry-Scope-Lücke
vor Veröffentlichung geschlossen. Final Kern/GUI/Original3/3 (143,28s:4,08/17,58/
121,62s), DAO vorher3,49s. Anwendung gebaut, KernQt-frei, alle7 Original-/Uploadhashes
unverändert. Keine neuen CTests/Windows-/Gesamtsuitennachweise; Original bleibt am
bekannten fehlenden Makroquelltext statt positiver Verkaufsausführung begrenzt.
Vorstufe3fe7245 privat veröffentlicht/Remote-Baum verifiziert; frischeDEB3fe7245 mit
beiden Datenblatt-Erweiterungen vollständig geprüft, liegt vor Klassen-Erweiterung.
Katalog456/386/70 unverändert; ohne Zwischenabschluss weitere offene Umfänge umsetzen.


Eindeutig gewöhnliche importierte Klassen können jetzt mit New aus der neuesten
wirklichen VBA-Archivrevision konstruiert werden. CLASS/VB_Name und kanonischer
Archivname müssen passen; VB_Base und tatsächlich zugeordnete Form_/Report_-Module
bleiben Dokument-CodeBehind. Unbelegter Präfix allein wird nicht zum Dokumenttyp
geraten. AllModules enthält nachgewiesene gewöhnliche Importklassen. UTF8/BOM,
UTF16LE und vorhandener CP1252-Ausführungsview ohne Änderung der Rohbytes;
verlustbehaftete undefinierte Fallbackbytes3251. Herkunft, Archiv-ID und rohe
Revision dauerhaft gebunden, auch semantisch gleiche anders kodierte Neuversion
invalidiert alte Instanz. Namespacekollisionen, TEMP-/Lösch-/Kontextguards bleiben.
Details NATIVE_VBA_CLASSES.md. Erstes bestehendes GUI/DAO-Bündel2/2 (20,44s:
17,74/2,70s), App gebaut: tatsächliche Konstruktion/Methoden/SQL, Kodierung/
Versionen, Metadaten, Dokumentausschlüsse und Originalarchivhash auf Testkopie.
Keine neuen CTests/Original-/Windows-/Gesamtsuitennachweise. Original-Bridge nur
aggregiert geprüft:28 Klassendateien/28 deklarierte Dokumentmodulnamen/0 gewöhnliche
Kandidaten; technische Quellen sind kein Windows-Exportoracle. Klassenkern bleibt
auf voriger geprüfter Umsetzung. Vorstufe7f40893 privat veröffentlicht/Baum geprüft,
frischeDEB7f40893 samt Payload/Rechten/Build-ID/gepinnten SQLite/MDBTools bestanden.
Katalog456/386/70 unverändert. Ohne Zwischenabschluss mit benannten Klassen-
Typbindungen und weiteren offenen Anforderungen fortfahren.


Benannte Projektklassen-Typbindungen: eigene/gewöhnliche archivierte echte Klassen
in skalaren Dim/Feldern/Parametern/Rückgaben/GetSet/TypeOf; optionalNothing und
owning-ByRef, Typfehler13, unbekannte Bindung3251 vor Body. Selbsttypen ohne
rekursive Instanziierung; direkte Formularfelder und SQL-Klassenresultatgrenze
geprüft. Details NATIVE_VBA_CLASSES.md. Klassenarrays/AsNew/Interfaces/externes
ByRef bleiben offen. Katalog456/386/70 unverändert; keine Paritätsbehauptung.

Nachweise dieses Blocks: native Kernprüfung nach If-Typbindungsfix1/1 (4,15s),
mit OptionExplicit-Fixture erneut1/1 (4,04s); Original-Ereignispfad1/1 (123,33s)
für die native VM, anschließend nur noch Qt-Formularfeld-Set ergänzt. Finales
vorhandenes GUI-Ziel1/1 (17,80s), App gebaut. Erster GUI-Basisblock17,88s bestand;
Erweiterungen korrigierten technische SQL-Registrierung/Deklarationsreihenfolge
und fanden den fehlenden echten Formularfeld-Set-Pfad. Betroffene Ziele gezielt
wiederholt; keine neue CTest-ID oder Vollsuite. Sieben Upload-/Fixturehashes
unverändert. Kein neuer Windows-/Distributions-/DEB-Nachweis; letzter geprüfter
DEB-Stand7f40893 liegt vor archivierten und typisierten Klassenerweiterungen.


## Benannte Projektklassen-Arrays (2026-11-09)

Feste/dynamische lokale und Modul-/Klassenfelder, ReDim/Preserve, geprüfte Set-
Elementzuweisungen, gleiche ByRef-Arrayparameter und Function-/Get-Rückgaben
verwenden jetzt die reale gebundene Projektklasse als Elementtyp. Initialisierte
und durch Erase zurückgesetzte Plätze enthalten Nothing; falsche Objekt- oder
Arraytypen scheitern13. Kopien behalten getrennte Arraywerte und gemeinsame reale
Objektreferenzen. Klassenname/Enum-Grenzen werden einschließlich expliziter ReDim-
Typangaben vor Effekten gebunden. Ein Array wird nicht als skalares Klassenobjekt
behandelt. VarType ist8201; VBA-TypeName löst den kanonischen tatsächlichen Namen
im Projekt auf, auch für undimensionierte Rückgaben. Debugger-/C++-Typanzeigen
behalten die normalisierte Deklarationsschreibweise, ohne interne Typpräfixe.

For Each behält die normativ erforderliche Variant-Schleifenvariable für Arrays;
die Klassentypisierung der Elemente hebt diese Regel nicht auf (MS-VBAL5.4.2.4).
Reale Rückgabearrays werden einmal enumeriert. Bestehende Rang-/Bounds-/Preserve-/
Speicher-/ByRef-Element- und Mutationsgrenzen bleiben erhalten. Keine allgemeinen
Collection-/DAO-Typarrays, AsNew, Interfaces oder vollständige Referenzparität.
Dies ersetzt die vorherige Klassenarray-Grenze für den beschriebenen Umfang.

App gebaut; vorhandenes Kern-/GUI-Bündel nach Fixture-/Binderkorrektur2/2 bestanden
(22,53s:4,01/18,52s). Erste Prüfung fand die falsche ForEach-Fixture-Deklaration;
danach war der skalare Objektbindungszweig für ganze Klassenarrays zu korrigieren.
Nur betroffene bestehende Ziele wiederholt, keine neue CTest-ID/Original-/Windows-
oder Vollsuite. Der GUI-Archivtest umfasst gewöhnliche importierte Klassen mit
solchem Arrayfeld und erhält den ursprünglichen Archivhash. Kein neuer DEB-Nachweis.
F042/F043 bleiben offen; Katalog456/386/70 unverändert. Ohne Zwischenabschluss an
nachgewiesenen benutzerdefinierten Klassen-Standardmitgliedern und weiteren offenen
Anforderungen weiterarbeiten.


## Deklarierte Klassen-Standardmitglieder (2026-11-09)

Ein tatsächlich vorhandenes Memberattribut VB_UserMemId=0 bindet jetzt einen
öffentlichen Property Get oder eine öffentliche Function als Klassenstandard.
Die ursprünglichen Attributzeilen/Zeilennummern bleiben erhalten; der Parser
überspringt nur den passenden geschlossenen0-Attributsatz der gewählten Prozedur.
Unbekannte/private/nicht lesbare Standardmitglieder, widersprüchliche Namen und
nicht implementierte andere Member-IDs/Enumeratoren werden vor Konstruktion
abgewiesen. Ein unannotiertes Item wird nicht zum Klassenstandard erfunden.
Ohne deklarierten Standard schlägt ein direkter Standardzugriff438 fehl.

Implizite skalare Auswertung verwendet nur einen Standardleser, der ohne
Pflichtargument aufrufbar ist. Getterfehler und die vorhandene16-stufige Grenze
rekursiver Standardwerte bleiben erhalten. Explizite Positions-/Namensargumente,
Missing/optionale Indizes und vorhandene Property-Schreibbinder werden verwendet;
Standardargumente bekommen die echte Signatur, unabhängig von einem anderen
öffentlichen Item-Member. Namensfehler448 werden vor Getteraufruf erkannt.
Standard-Schreiben verwendet tatsächliche Let/Set-Accessoren beziehungsweise
vorhandene geprüfte Rückgabeweiterleitung. Benannte Setterindizes, Klassen-Bang-
Syntax, vollständiges externes ByRef, allgemeine COM-Attribute/-IDs und
Enumeratoren bleiben offen; objektwertige Standard-Set-Kombinationen sind kein
separater Windows-Nachweis. Diese Umsetzung ersetzt die frühere pauschale
Standardmitglied-Attributgrenze nur für den beschriebenen0-Pfad.

App gebaut, erstes Laufzeitbündel der vorhandenen Kern-/GUI-Ziele2/2 bestanden
(22,36s:4,07/18,29s). Gemeinsame technische Quellen prüfen Default-Get/Let,
impliziten Skalarwert, optionale/namentliche Indizes, unverwechselbares Item,
private Zustände und widersprüchliche/andereIDs. Der echte Projektpfad sowie eine
gewöhnliche archivierte Defaultklasse und die unveränderte ursprüngliche
Archivsignatur sind geprüft. Keine neue CTest-ID/Original-/Windows-/Vollsuite.
Katalog456/386/70 unverändert, F042/F043 bleiben offen. Nach privater Veröffentlichung
aktuelles Paket prüfen und ohne Zwischenabschluss weitere Anforderungen bearbeiten.


## Skalare automatische Klassenvariablen As New (2026-11-09)

AsNew ist nun für skalare lokale/statische Variablen, Modul-/Klassenfelder und
öffentliche Formularfelder gebunden: gewöhnliche tatsächliche Projektklassen
sowie die schon vorhandenen nativen Collection/Dictionary/FileSystemObject-
Konstruktoren. Prüfung bindet die reale konstruierbare Klasse vor Effekten,
führt aber keine Initialisierung aus. Erst tatsächliches Lesen erzeugt eine
Instanz; Set auf Nothing erzeugt keine Zwischeninstanz, späteres Lesen eine neue.
Normbasis MS-VBAL5.2.3.1.4/5.3.1.6. Eigene Parameter bleiben normale ByRef-Parameter;
eine Rücksetzung durch sie erhält die automatische Eigenschaft des Aufrufers.

Statische/Modulzustände und gewöhnliche Felder behalten tatsächliche Referenzen.
Direkter Me-/öffentlicher Feldzugriff und lokale Aliasansichten werden bei AsNew
synchronisiert, ohne ausstehenden lokalen Set-Wert zu überschreiben. Metadaten-
und Debuggeranzeigen lesen den Zustand ohne Erzeugung. Skalare Standardzuweisung
verwendet den realen Default-Setter. Direkte Formularzugriffe halten den Modulframe
bei Initialisierungs-Reentry fest und prüfen Schließen; Zielobjekte werden während
Default-Aufrufen gehalten. Arrays mit AsNew/automatischen Elementen, WithEvents,
Defaultinstanzen und komplette Terminierung/GC bleiben offen, keine allgemeine
AsNew- oder Vollklassen-Parität behauptet. Unbekannte/nicht konstruierbare Klassen
werden nicht als leere Scheinobjekte angelegt.

App gebaut; erstes bestehendes Kern-/GUI-Bündel2/2 (22,90s:4,40/18,49s). Nach
abschließendem Aliasabgleich ergänzt und geprüft:2/2 (22,46s:4,25/18,20s).
Gemeinsame technische Quellen zählen tatsächliche Initialisierungen, prüfen
verzögerten ersten Zugriff/SetNothing/Identität, ByRef-Reset, Static-/Modul-Erhalt,
Native-Collection/Dictionary-Neuanlage, echte Default-Zuweisung, Me-/unqualifizierte
Feldsicht und echte öffentliche Formularfelder. Kein neues CTest/Original-/Windows-
oder Gesamtsuite-Orakel. Katalog456/386/70 unverändert, vollständige F042/F043 offen.

Vorstufe dced5e6 privat remote/Baum geprüft; aktuelles DEB dieses Stands unter
/workspace/work/class-default-members/package/linaccess-0.7.0-Linux.deb besteht
Payload/Build-ID/Rechte/keineFixtures, MDBTools1.0.1 und SQLite3.54.1 samt
NestedJOIN/NULL-/Counter-Prüfung; bestehendes DAO-Ziel1/1 (2,65s). Dieses Paket
enthält die zuvor geprüften Klassen-/Array-/Default-Stufen, noch nicht AsNew.
Ohne Zwischenabschluss mit belegten weiteren Deklarations-/Laufzeitbindungen
fortfahren; keine pauschale Referenzanforderung oder Etappen-Abschlussanruf.


Enum-Typbindungen (2026-11-09): tatsächliche quelllokale Enum-Namen werden gemäß
MS-VBAL5.2.3.4 in As-Deklarationen als Long gebunden, einschließlich Parametern/
Rückgaben, Feldern/Static, Arrays, GetLet und typisierten Konstanten. Einzeilige
If-Pläne erben den echten Deklarations-/Options-/Konstantenkontext; Cache vermischt
keine fremden Typbindungen. Details VBA_ENUM_TYPE_BINDINGS.md. Bestehendes Kern-/
GUI-Bündel nach Fixturekorrektur2/2 (23,08s:4,60/18,48), App; Original-Ereignispfad
1/1 (121,37s), bekannte Makrogrenze beibehalten. Sieben Hashes unverändert. Kein
neues CTest/Windows-/Vollsuite-/DEB-Orakel. Crossmodule-/qualifizierte Typreferenzen,
UDTs und vollständige Unicode-/Compilersemantik offen;456/386/70 unverändert.
Vorstufe455abfa privat remote/Baum geprüft. Ohne Zwischenabschluss mit echten
editierbaren einzelnen Tabellenfeldprojektionen in Abfragedatenblättern fortfahren.


Editierbare einzelne Abfragefelder: EDITABLE_QUERY_PROJECTIONS.md beschreibt die
geprüfte physische Feldzuordnung mit Alias/Subset/Reihenfolge, verborgenem echtem
Schlüssel, SQL-Speichern unveränderter ausgelassener Felder, Sortierung nach
Ausgabealias sowie Layout-/Header-Erhalt nach Requery. Nur angeforderte Projektion
wird erweitert; andere Stern-Caller bleiben unverändert. App/GUI1/1 (18,42s),
nach Aliasnormalisierung/Unicode-Namensschutz nur betroffenes Modellziel1/1 (0,11s).
Keine neuen CTests/Original-/Windows-/Vollsuite-/Paketbehauptungen. Mehrspaltige
umgeordnete Zwischenablage zunächst geschützt3251; als nächste zusammenhängende
Etappe echte sichtbare Spaltenabbildung für Kopieren/Ausschneiden/Einfügen/Undo
implementieren. F079/F340/F373/456/386/70 bleiben offen/unverändert. Vorstufe03b3de2
privat veröffentlicht, Remote/Baum geprüft; ohne Zwischenabschluss weiterarbeiten.


Sichtbare Zwischenablage-Spaltenabbildung: PROJECTED_GRID_CLIPBOARD.md ersetzt die
vorherige mehrspaltige Projektionsverweigerung. Kopieren/Ausschneiden/Einfügen und
Undo/Redo verwenden dieselben geprüften realen Spalten, verborgene Schlüssel/
ausgelassene Werte bleiben geschützt; mapped Append bewahrt tatsächliche SQL-
Defaults. Vorhandenes GUI/Currency-Feld/Modell-Bündel im ersten Lauf3/3 (18,69s:
18,11/0,48/0,11), App gebaut. Reale SQL-/Undo-/Cut-/MIME-/Negativfälle, keine neuen
CTest-/Original-/Windows-/Vollsuiten-/Paketbehauptungen.456/386/70 unverändert;
vollständige F079/F340 bleiben offen. Vorstufe2d95062 privat remote/Baum geprüft.
Ohne Zwischenabschluss den tatsächlichen Makroreferenz-Blocker über die bestehende
Codex-Call-Schnittstelle prüfen und unabhängig weitere Anforderungen implementieren.


Native Objektarrays (2026-11-09): NATIVE_OBJECT_ARRAYS.md beschreibt die geschlossene
DAO-/Access-/Scripting-/VBA-Typbindung für Deklarationen, ByRef, Rückgaben, Preserve
und Set-Elemente mit echter Schnittstellenprüfung statt Klassenname. Reale lokale
Forms-Arraybindung hat Vorrang vor globaler Sammlung. Nach gezielter Korrektur
bestehendes Kern-/GUI-Bündel2/2 (24,46s:4,85/19,61), DAO zuvor1/1 (2,83s), App.
Keine neuen CTests/Original-/Windows-/Vollsuiten-/Paketbehauptungen.456/386/70
unverändert; vollständige F042/F043 offen. Vorstufe9bc07b2 privat remote/Baum geprüft.
Gespeicherten Makro-Codex-Call einmal ohne neuen Anruf fortgesetzt: failed, keine
Antwort/Entscheidung; Ende vor Agent-Abschlussverarbeitung. Makroreferenz bleibt
offen; unabhängig ohne Zwischenabschluss weiterprogrammieren.


ReDim-Deklarationswirkung: REDIM_DECLARATIONS.md beschreibt echte prozedurweite
Arraydeklaration bei ungebundenem Namen (auch OptionExplicit/übersprungene Zweige),
Def-/Elementtyp, ByRef/Rückgaben/Static und Schutz bestehender Modulbindungen.
App; bestehendes Kern-/GUI-Bündel2/2 (22,96s:4,74/18,22). Gemeinsame Integration
mit nativen Objektarrays: DAO2,71s + Original-Ereignispfad119,91s,2/2 in122,62s.
Bekannte Makrogrenze beibehalten; sieben Hashes unverändert. Keine neuen CTests/
Windows-/Vollsuiten-/Paketbehauptungen;456/386/70 unverändert. Vorstufe58b1723
privat remote/Baum geprüft. Ohne Zwischenabschluss mit belegten Err-Reset-
Grenzen nach MS-VBAL6.1.3.2.1.1 weiterprogrammieren.


Automatisches Err.Clear: VBA_ERROR_CLEAR_BOUNDARIES.md dokumentiert echte aktive
OnError-/ExitSub-/ExitFunction-/ExitProperty-Grenzen nach MS-VBAL6.1.3.2.1.1,
alle fünf Eigenschaften, übersprungene Anweisungen und zurückbehaltene ErrObject-
Referenzen. App; bestehendes Kern-/GUI-Bündel im ersten Lauf2/2 (23,07s:4,76/18,30).
Keine neue CTest-ID/Windows-/Vollsuitenbehauptung. Vollständige umgebungsweite
Err-Identität und Fehlerhistorie bleiben offen;456/386/70 unverändert. Vorstufe
113b829 privat remote/Baum geprüft; Original zuletzt119,91s und sieben Hashes
unverändert vor dieser begrenzten Reset-Korrektur. Ohne Zwischenabschluss ein
aktuelles Debian-Paket erstellen und weitere echte Laufzeitlücken bearbeiten.


Projektweite native Err-Instanz: VBA_ERROR_ENVIRONMENT.md beschreibt echte
Host-/Projektbindung statt pro-Prozedur-Neuanlage, gemeinsame Referenzen über
Modul/Formular/Klasse, lokale Handler, tatsächliche Fehlerweitergabe und getrennte
Datenbanken. Reset invalidiert alte Referenzen; geschlossene Verbindung verweigert
Eigenschaftszugriff3420. App; initial GUI18,42s/DAO2,75s/Original118,10s bestanden.
Eine alte Kernfixture benötigt für ihren Null-Start ausdrücklich Err.Clear; nur
Kern wiederholt, bestanden5,03s. Sieben Hashes unverändert; bekannte Makrogrenze.
Keine neue CTest-ID/Windows-/Vollsuitenbehauptung;456/386/70 unverändert. Vorstufe
55096a7 privat remote/Baum geprüft. Dessen aktuelles DEB unter
/workspace/work/error-clear-boundaries/package/linaccess-0.7.0-Linux.deb geprüft
(Build-ID/Payload/Rechte/keineFixtures, MDBTools1.0.1, SQLite3.54.1 JOIN/NULL/Counter),
noch ohne diese Projektbindung. Ohne Zwischenabschluss mit echter CVErr-
Fehlerwertübernahme/Let-Konvertierung nach MS-VBAL6.1.2.3.1.14 weiterprogrammieren.


CVErr-Fehlerwertübernahme: CVERR_CONVERSION.md beschreibt unveränderte echte
Error-Argumente und sonstige Long-Let-Konvertierung, typisierte Rundung und
Grenzfehler. App; bestehendes Kern-/GUI-Bündel im ersten Lauf2/2 (22,83s:4,65/18,17).
Echte gespeicherte skalare SQL-VBA-Funktionen lesen denselben Projekt-Err.Number907
und liefern CVErr(Error2015) als Text, ohne einen neuen Fehler auszulösen. SQL-
Objekt-/Array-Schutz bleibt. Keine neue CTest-ID/Original-/Windows-/Vollsuiten-/
Paketbehauptung;456/386/70 unverändert. Vorstufe725b4bd privat remote/Baum geprüft.
Aktuell geprüftes DEB55096a7 liegt vor Projekt-Err/CVErr. Ohne Zwischenabschluss
mit nativen Error/Error$-Funktionen und getrennter zuletzt-ausgelöster Fehlernummer
nach MS-VBAL6.1.2.3.1.15 weiterprogrammieren.


Native Error/Error$: VBA_ERROR_TEXT.md beschreibt optionale/benannte/qualifizierte
und nackte Aufrufe, Long-Konvertierung, begrenzte invariante englische Meldungen
und getrennte zuletzt-ausgelöste Fehlernummer. Err.Number-Zuweisung und Clear/
OnError/Exit ersetzen diese Historie nicht. Echte SQL-VBA-Funktionen unterscheiden
Err.Number907 von zuvor ausgelöstem91. Projektfunktion/lokales Array Error haben
unqualifiziert Vorrang; VBA.Error bleibt Intrinsik, Access-Feld Error bleibt Feld.
App; initial bestehendes Kern-/GUI-Bündel2/2 (23,15s:4,81/18,33). Nach Feldnamens-
Schutz nur betroffenes Kernziel wiederholt1/1 (4,87s). Keine neue CTest-ID/Windows-/
Vollsuitenbehauptung;456/386/70 unverändert. Vorstufe2d3e8db privat remote/Baum
geprüft. Letztes Original118,10s/Hashes vor CVErr/Errortext; DEB55096a7 vor Projekt-
Err/CVErr/Errortext. Ohne Zwischenabschluss aktuelles Paket prüfen, danach echte
ByRef-Arraysperren nach MS-VBAL5.4.3.3 weiterbearbeiten; Erase-/ForEach-Grenzen
nicht ungeprüft als vollständige native Semantik deklarieren.


Echte ParamArray-Aliase/Arraysperren: PARAMARRAY_ALIASES.md dokumentiert skalare/
Objektvariablen und unterstützte Arrayzellen, eigene Literal-/Klammerwert-Slots,
doppelte Aliase, echte ByRef-Weitergabe, typisierte Rückschreibfehler und Schutz
unaufgelöster/Fassaden-/Arraywertfälle. Aktive ByRef-Arraywurzel liefert10 statt
3251 und wird nach Aufrufende freigegeben. App; initial GUI18,94s/DAO2,78s/
Original118,62s bestanden. Alte Literalpack-Erwartung angepasst und gewöhnliche
Aliaswerte geliehen statt Byte/Memo-Kopie; nur betroffenes Kern-/DAO-Bündel
wiederholt2/2 (7,97s:5,06/2,90), GUI/Original nicht erneut. Sieben Hashes unverändert;
bekannte Makrogrenze. Keine neuen CTests/Windows-/Vollsuitenbehauptungen;456/386/70
unverändert. Vorstufe9affc5a privat remote/Baum geprüft, dessen frisches DEB unter
/workspace/work/error-text-functions/package/linaccess-0.7.0-Linux.deb besteht
Build-ID/Payload/Rechte/keineFixtures, MDBTools1.0.1 und SQLite3.54.1 JOIN/NULL/Counter;
Paketintegration DAO2,76s/Original120,25s,2/2 in123,02s. Noch ohne diesen Aliasumfang.
Ohne Zwischenabschluss mit echten benannten Property-Setter-Indexargumenten
weiterprogrammieren; keine Getter zur Metadatenentdeckung ausführen.


Benannte Property-Setter-Indizes: NAMED_PROPERTY_SETTERS.md beschreibt echte
Let-/Set-Metadaten (auch ohne Get), optionale/Objektindizes und gemeinsame
Namens-/ByRef-Adresszuordnung einschließlich ParamArrayzellen. Projektklasse,
Default-Property und echte Formularfassade geprüft; externe ByRef-Grenze bleibt.
Getter werden nicht zur Entdeckung ausgeführt. App/Kern5,02s; GUI-Watchdog zeigte
fehlende Bereinigung der eingeführten NamedWriter-Testklasse vor späterer TEMP-
Katalogprüfung. Klasse scoped bereinigt, Katalogschutz unverändert; nur GUI
wiederholt, zuletzt18,54s bestanden. Sprach-Doppelmatrizen bleiben im Kern; GUI
behält echte Qt-/SQL-/Klassen-/Formularprüfungen und präzise20s-Schranke. Keine
Nachrichteninhalte in dauerhaftem Watchdoglog, keine neuen CTests/Windows-/
Vollsuitenbehauptungen;456/386/70 unverändert. Vorstufe859d2ef privat remote/Baum
geprüft. Letztes Original118,62s/DAO2,90s/Hashes liegen vor diesem Umfang; geprüftes
DEB9affc5a unter /workspace/work/error-text-functions/package/linaccess-0.7.0-Linux.deb
liegt vor ParamArray/benannten Settern. Ohne Zwischenabschluss Set-RHS-Auswertung
nach MS-VBAL5.4.3.9 vor dynamischer Zielauflösung weiterimplementieren.


Set-RHS-Reihenfolge: SET_EVALUATION_ORDER.md beschreibt einmalige Objektauswertung
vor dynamischem Empfänger/Index und Alias-Aktualisierung nach MS-VBAL5.4.3.9.
Echter Empfängerwechsel, Default, ReDim durch RHS, eigener/echter Formularsetter,
RHS-Fehler ohne Zielaufruf und danach Nothing→91 geprüft. Fixtureobjekte/Klassen
scoped bereinigt. App; erstes bestehendes Kern-/GUI-/DAO-/Original-Bündel4/4 in
145,89s (5,53/19,19/2,75/118,41), sieben Hashes unverändert. Bekannte Makrogrenze
bleibt; keine neuen CTests/Windows-/Vollsuiten-/Paketbehauptungen;456/386/70
unverändert. Vorstufe d618eaf privat remote/Baum geprüft. Letztes geprüftes
DEB9affc5a unter /workspace/work/error-text-functions/package/linaccess-0.7.0-Linux.deb
liegt vor ParamArray/benannten Settern/Set-Reihenfolge. Ohne Zwischenabschluss
mit echtem Undo/Redo vorgemerkter nativer Mehrfachlöschungen weiterprogrammieren.
