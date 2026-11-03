# Codex Cloud — LinAccess vollständige Weiterentwicklung

## Auftrag

Entwickle dieses private Repository `BlueDragon4251/LinAccess` selbstständig
bis zu einem möglichst vollständigen, tatsächlich funktionsfähigen,
Linux-nativen Microsoft-Access-Ersatz. Die Programmierung erfolgt **in C++20
mit Qt 6** (nicht Python, nicht Electron); primäre Plattformen sind **Debian
12/13 und Ubuntu 24.04+**. Nutze die hier vorhandene Codebasis als Ausgangspunkt.
Microsoft Access 2003/JET4 ist der erste verbindliche Kompatibilitätsmaßstab,
später Access 2007–365/ACE (`.accdb`).

**Wichtig: Quellimport und statische VBA-Prozedurenerkennung sind keine
VBA-Ausführung.** Baue echten Laufzeit-Support, nachvollziehbar am Verhalten
der beigefügten Access-2003-Anwendungen verifiziert. Nicht nur Mockups,
Syntax-Highlighting, Status-Badges oder Platzhalter. Der Nutzer erwartet eine
möglichst identische Benutzererfahrung und dass Originalprojekte inklusive
VBA-Code möglichst ohne manuelle Änderungen laufen.

## Bestehender Stand (0.7.0, keine Abnahme)

- Qt6-Desktop-App mit SQLite-basiertem `.linaccess`-Projektspeicher,
  Tabellen-/SQL-/Formular-/Bericht-Grundfunktionen.
- MDBTools-Import: Keksrezepte 2/2 Tabellen; Kassa 29/29 Tabellen in
  **verlustbehaftete** SQLite-Kopie. Datentypen, PK/FK, Indizes, BLOBs,
  Standardwerte und Jet-SQL sind noch nicht allgemein semantiktreu.
- Die Windows/Access-Bridge exportiert `SaveAsText`-Formulare, Berichte,
  Abfragen und VBIDE-Module. Kassa-Bridge aus Microsoft Access 2003 (11.0):
  192 Artefakte (47 Formulare, 9 Berichte, 100 Abfragen, 4 Access-Module,
  32 VBA-Quelldateien), ~187 erkannte VBA-Prozeduren.
- Layoutparser erkennt Kassa-Formulare einschließlich 685 Controls und
  43 Buttons; die Vorschau ist ausdrücklich nicht 1:1.
- Experimenteller C++-VBA-Interpreter führt nur eine kleine Syntax-/API-Teilmenge
  aus, z. B. `Form_Load`, Klickereignisse, einfache Variablen, `If`, `For`,
  einige `DoCmd`-Aufrufe, `Me`, `CurrentDb.Execute` und einfache SQL-Aktionen.
  **DAO/ADO-Recordsets, allgemein gültige VBA-Semantik, COM/ActiveX und der
  vollständige Access-Objektgraph fehlen.** Die Kassenanwendung läuft nicht.
- `docs/COMPATIBILITY.md` ist der ehrliche Kompatibilitätsstatus.

**Neue reproduzierte Auffälligkeit beim Upload (30.10.2026):** Ein frischer
Linux-Build mit MDBTools 1.0.0 importierte beim Kassen-MDB-Regressionslauf
schwankend nur 0 bis 20 von 29 Tabellen; Fehler umfassen `MDBTools: offset
... is beyond EOF` und Schema-/CSV-Spaltenfehler. Eine zeitweise gekürzte
Arbeitskopie wurde durch SHA-256-geprüfte Archivwiederherstellung repariert,
aber auch mit intakter Originaldatei blieb der MDBTools-Import instabil.
Der aktuelle `ctest`-Lauf ist für `complex_real_mdb_table_copy` **rot**.
Das ist **kein Erfolg für vollständigen Tabellenimport**.
Vor Architektur- oder Vollständigkeitsversprechen die Ursache anhand originaler
SHA-256-Dateien und MDBTools-Versionen exakt analysieren, im Issue festhalten
und reparieren. Die vormals dokumentierten 29/29 waren ein früherer Teststand,
nicht der aktuelle reproduzierte Befund.

## Reproduzierbare Testdaten — zuerst auspacken

**Das Git-Repository ist privat und MUSS privat bleiben.** Es enthält
originale Kassen-/Rezepte-Daten und ausführbaren VBA-Code, potentiell mit
Zugangsdaten. Keine Datenbankinhalte oder VBA-Literale in öffentliche Logs,
Issues, Releases oder externe Services kopieren.

```bash
./scripts/restore-fixtures.sh
```

Das setzt vier unveränderte Dateien mit SHA-256-Verifizierung wieder zusammen:

- `fixtures/raw/Keksrezepte.mdb`
- `fixtures/raw/Kassa_CSE_V2.0.mdb`
- `fixtures/raw/Keksrezepte-Bridge.zip`
- `fixtures/raw/Kassa_CSE_V2.0-Bridge-2003.zip`

Die Bridge-Inhalte werden zusätzlich nach `fixtures/unpacked/` entpackt;
Originalordner siehe `fixtures/README.md`. Vor jedem Feature Daten und
Quellcode daraus untersuchen, Referenzen auflösen, Tests formulieren. Nur mit
**Kopien** arbeiten, um echte Testdaten nicht versehentlich zu zerstören.
Es geht um die normale Access-Semantik: VBA darf Tabellen verändern, Datensätze
löschen und SQL `DELETE`, `UPDATE`, `INSERT`, `DROP` ausführen. Keine
LinAccess-spezifischen Zustimmungsdialoge für jede Datenänderung einbauen.
Access-eigene `DoCmd.SetWarnings`/Dialog-/Fehlermodalitäten korrekt abbilden.

## End-to-End-Zielkatalog — nicht bei Teilimplementierungen stehenbleiben

1. **Dateien / Speichermodell**: `.mdb`/`.accdb` importieren; Tabellen,
   Primärschlüssel, Indizes, Beziehungen, Constraints, Standardwerte,
   Null-Semantik, Lookups, Unicode, OLE-/Attachment-Daten und Systemmetadaten
   erhalten. Eigene Dateien projektübergreifend konsistent speichern;
   Import-/Export-/Roundtrip-Tests. Wo natives Schreiben in Access-Binärformate
   objektiv nicht möglich ist, präzise Kompatibilitätsgrenzen dokumentieren.
2. **Access-SQL**: Jet4/ACE-Abfragen und Parameterabfragen vollständig
   analysieren und ausführen oder semantisch korrekt übersetzen (JOINs,
   DISTINCT, UNION, TOP, Transform/PIVOT, Ausdrucksfunktionen, Datumsformate,
   Wildcards, Typkonversion, DDL/DML, Null). Importierte 100 Kassenabfragen
   einzeln testen; abhängige Formular-RecordSources müssen funktionieren.
3. **VBA-Engine**: tatsächlich kompilieren/interpretieren, nicht per Regex
   erraten. Module, Klassen, Formular-/Berichts-CodeBehind, Typen und
   Variants, Arrays, Strings, ByRef/ByVal, Default-Properties, Objektvariablen,
   Scope/Lifetime, Error Handling (`On Error`, `Err`, `Resume`), `With`, `Select
   Case`, `Do`, `For Each`, `Optional`, `ParamArray`, Calls/Returns, Properties,
   Collections, `Set`/`Nothing`, Datums-/Währungssemantik, Operatorpriorität,
   Referenzbindung, Ereignisaufrufe und Debugger mit Breakpoints/Locals.
4. **Access/DAO/ADO-Objektmodell**: `Application`, `Forms`, `Reports`, `Me`,
   `DoCmd`, `CurrentDb`, `DBEngine`, `DAO.Database`, `Recordset`/`Recordsets`,
   `Fields`, `QueryDef`, `TableDef`, `Workspace`, `Parameters`, Transaktionen,
   `DLookup`/`DCount`/`DSum`, `OpenRecordset`, `.Edit/.AddNew/.Update/.Delete`,
   `MoveFirst/Last/Next`, `.EOF/.BOF`, `RecordSource`, `ControlSource`,
   `Requery`, `Filter`, `Sort`, Bookmarks. ADO-Verweise gesondert behandeln.
   Keine nur scheinbaren Methoden ohne reale Wirkung.
5. **Ereignisse**: Access-Reihenfolge und Cancel-Semantik für `Open`, `Load`,
   `Current`, `BeforeUpdate`, `AfterUpdate`, `Dirty`, `Click`, `DblClick`,
   `KeyDown/Press/Up`, `GotFocus/LostFocus`, `Close`, `Unload`, Timer,
   Kombinationslisten, Subform-Events. Originale Kassen-Befehle per UI-Klick
   ausführen und Änderungen an Datensätzen verifizieren.
6. **Oberfläche**: Access-ähnlicher Objekt-Navigator mit Tabellen, Abfragen,
   Formularen, Berichten, Makros, Modulen; Design-/Datenblatt-/Formularansicht,
   Properties, Toolbox, grafische Designer, Reiter/Fenster, Ausdruckseditor,
   SQL-Designer, Formular-Navigation. Access-2003-TWIP-Layout, Schriftarten,
   Größen, Farben, Ausrichtung, Ereigniszuordnung, Combo-/List-Boxes,
   Subforms, Registerkarten und Steuerelemente möglichst genau nachbauen.
7. **Berichte**: Layout/Wiederholbereiche, Header/Footer, Gruppierung,
   Summen, Formatierung, RecordSources, Report-VBA, Druckvorschau,
   PDF/Drucken und Seitenumbrüche auf reale Kassenberichte prüfen.
8. **Makros/Referenzen**: Access-Aktionsmakros lesen/ausführen, Verweise auf
   Bibliotheken klassifizieren und wo möglich native Äquivalente bauen.
   Referenzen auf Windows-only COM/ActiveX/WinAPI präzise identifizieren:
   Linux-native 1:1-Ausführung ist dort technisch nicht generell möglich.
   Explizite Diagnose, ggf. optionales getrenntes Windows/Wine-Kompatibilitäts-
   backend erforschen; niemals fehlende Features als funktionierend ausgeben.
9. **Bridge**: Access-2003/365-Exporter robust machen, korrekte 32-/64-bit
   COM-Registrierung/Versionsprüfung, unveränderte Modulbytes, Encoding,
   Layout-Properties, VBA-Referenzen, Objektabhängigkeiten und Fehlermanifest.
   Export für `.mdb` und `.accdb` prüfen; GUI-Import mit Einzelobjekt-Status.
10. **Produktreife**: Absturzsicherheit, Undo/Redo, konsistente Transaktionen,
    verständliche Fehlermeldungen, Übersetzungen DE/EN, Dokumentation,
    reproduzierbare Debian-/Ubuntu-.deb-Pakete, automatisierte Builds,
    UI- und Integrationstests; end-to-end Betrieb ohne Windows.

## Verbindliche Validierung

- `cmake -S . -B build -DCMAKE_BUILD_TYPE=Release`, `cmake --build build`,
  `ctest --test-dir build --output-on-failure`, `cpack -G DEB ...`.
- Tests auf Debian 12, Debian 13 und Ubuntu 24.04 (CI existiert).
- Original-MDBs bleiben SHA-256-identisch; Aktionen gegen isolierte Kopien.
- Jedes der 192 Kassen-Bridge-Objekte erfassen, jede der 32 VBA-Dateien und
  alle erkannten Prozeduren nach Eigenschaften klassifizieren; laufende,
  fehlerhafte und nicht unterstützte Fälle explizit ausweisen.
- Regressionen vergleichen: Formularöffnen, Navigieren, gebundenes Editieren,
  Buttons, `CurrentDb.Execute`, Recordsets, Abfragen und Berichtsausgabe.
  Ausführbare Test-Assertionen statt bloßer Importzähler.
- Keine Angabe "vollständig/100% kompatibel", solange reale Tests oder
  Plattformbeschränkungen das nicht belegen. Das ursprüngliche Ziel ist
  maximale Funktions-/UI-Parität, nicht ein künstlich geschöntes Ergebnis.

## Arbeitsweise

Erstelle eine nachvollziehbare Checkliste (GitHub-Issue und/oder Datei),
zerlege in überprüfbare Arbeitspakete, implementiere Code selbstständig,
nutze Commits, PRs und CI. Untersuche die vorhandene Implementierung, bevor
du Architektur ersetzt. Priorität: **echtes Kassa-VBA mit DAO und
Ereignisbindung > korrekter Jet-SQL-/Datenzugriff > UI-/Bericht-Parität >
ergänzende Designerfeatures**. Arbeite iterativ und setze nach Tests mit
der nächsten offenen Aufgabe fort, bis reale Kompatibilitätsgrenzen oder
externe Blocker begründet dokumentiert sind. Liefere nach jeder größeren
Etappe messbare Belege und verbleibende Fehler, keine unbelegten
"alles fertig"-Meldungen.


## Fortschrittsnachtrag: native Date-Werte (2026-11-03)

Der Laufzeitkern besitzt jetzt `VbaValue::Kind::Date` mit Double-OLE-Serienwert;
maßgebliche Dateien sind vba_expression.*, vba_calendar.*, vba_runtime.cpp und
vba_qt_value.cpp. VarType 7 nicht mit DAO-Feldtyp 8 verwechseln. Now/Date/Time,
CDate, DateSerial/TimeSerial und DateAdd sind begrenzt nativ angebunden. Qt-/DAO-
Datumsfelder behalten Date; SQLite verwendet bisher ISO mit ganzen Sekunden.
Keine Untersekunden still abschneiden und keine Gebietsschemaparität behaupten.
SQL-Projektionen binden Date-Operatoren typisiert; keine vollständige SQL-AST-/
Filter-/Aktions-/Aggregatparität daraus ableiten. QCalendarWidget ist noch nicht
vollständig auf denselben Date-Vertrag umgestellt. Details und Grenzen:
`docs/DATE_RUNTIME.md`, aktuelle Belege in `docs/PROGRESS.md`.

Der vollständige gemeinsame Projektionsstand a8d3061 bestand 48/48 Tests,
GUI-freien Build, Debian-Paketprüfung und vier Fixture-Hashes. Spätere Inkremente
haben zunächst gezielte Prüfungen; ihren tatsächlichen Git-/Build-Stand prüfen.
Das ist weiterhin keine vollständige Access-2003-Parität.


## Technische Fortschreibung: native Collections und Bearbeitung (2026-11-03)

VbaObject hat jetzt optionale className-/enumerationGuard-Hooks. Standardwerte
sind leer; unbekannte Objektklassen bekommen keinen erfundenen TypeName. Die
native Collection erhält Variant-/Objektreferenzen und bindet typisierte
Modulfelder/Parameter/Rückgaben einschließlich optionalem Nothing. ByVal teilt
das Objekt, nicht die Ersetzung der Referenz; ByRef kann die Referenz ersetzen.
Eine Enumeration-Lease verhindert Bestandsänderungen während For Each vor der
Mutation. Regionale Schlüssel, zyklische Lebensdauer und vollständige Klassen-/
COM-Semantik bleiben offen; Details in docs/COLLECTION_RUNTIME.md.

DoEvents pumpt echte Qt-Ereignisse im nativen Dokumenthost. Akzeptiertes Quit
muss auch die suspendierten VBA-Frames beenden, selbst mit Resume Next.
Native Ausführungstiefe und verzögerte Dokumententsorgung beibehalten.
Timer liefert echte lokale Sekunden seit Mitternacht als Single.

native_clock_keys.hpp enthält die gemeinsame Datums-/Uhrzeit-Tastenzuordnung
für Formulare und editierbare Tabellen-Textzellen. ISO-/Sekundenformat bleibt
beschränkt. Nach Tabellenspeicherung wird ein bestehender Datensatz anhand
seines vollständigen Primärschlüssels wieder aktiviert; NULL/neue generierte
Schlüssel, schlüssellose Tabellen und große Datenmengen bleiben begrenzt.
Das Bearbeiten-Menü routet echte Qt-Tasten an den fokussierten Texteditor und
registriert keine konkurrierenden globalen QAction-Tastenkürzel. Separate
Form-/Reportfenster benötigen weiterhin vollständige MDI-/Menüintegration.


## Fortsetzung nach abgelaufener CLI-Anmeldung (2026-11-03)

Der GitHub-CLI-Token wurde während der Entwicklung ungültig. Der authentifizierte
GitHub-Connector konnte den privaten Branch per Git-API dennoch aktualisieren:
Runtime-Checkpoint d94ec0382b637dac6d306d5af87f02e8a8756c19 auf
development/access-compatibility hat exakt denselben Git-Quellbaum
9ab113b0ba92d0021944b5ef4b6a5b2e8a13fe33 wie der lokale Checkpoint
b96e7a6817f7f1e084070065d62101d976daefa4. Alle 25 geänderten Blob-Hashes, der
vollständige Tree und der tatsächliche Remote-Branch wurden abgeglichen. Der
Remote wurde ohne Force vom bisherigen 825047c weitergeführt; privat bestätigt.
Die einzelnen lokalen Entwicklungscommits bleiben erhalten, zusätzlich gesichert
unter refs/linaccess/checkpoints/b96e7a6. Lokale und API-Commit-IDs unterscheiden
sich wegen der Git-Metadaten; die geprüften Quellen sind identisch.

Solange die CLI-Anmeldung ungültig ist, den tatsächlichen Remote-Stand mit dem
GitHub-Connector lesen. Weitere Git-API-Checkpoints dürfen nur mit identischem
geprüftem Source-Tree, privaten Repo-Metadaten und expected_sha-Schutz auf dem
Arbeitsbranch angelegt werden. Lokale Historie nicht mit Force/Hard-Reset
verwerfen. Nach erneuerter CLI-Anmeldung erst fetch und Status/Diff prüfen, dann
die Quellbaum-identische API-Historie ohne Arbeitsverlust zusammenführen.
Die externe private Checkpoint-Metadatei liegt unter
/workspace/linaccess-onboarding/private-api-checkpoint.json; sie enthält keine
Zugangsdaten. Der aktuelle Runtime-Checkpoint bestand 48/48 CTests (87.26 s),
GUI-freien Core (1.55 s), Debian-Paket und alle vier Originalfixture-Hashes.
Vollständige Access-Parität ist weiterhin nicht erreicht.


## Bereichseinfügen und wiederholbares Speichern (2026-11-03)

Striktes TSV-Bereichseinfügen ist für bestehende Zellen eines sauberen
Manual-Submit-Modells implementiert, inklusive vollständiger Typ-Vorprüfung.
Native Tabellen speichern mit AccessTableModel::submitTransactionally: alle
SQL-Arbeiten vor dem Commit, keine vorzeitige Qt-Cache-Submission. Bei späterem
Constraint-Fehler bleiben Updates/Inserts/Deletes für Korrektur und Wiederholen
erhalten. Geerbtes direktes Qt submitAll bietet diese Garantie nicht.
Details und Grenzen in docs/DATASHEET_CLIPBOARD_RUNTIME.md.
Vollsuite 48/48 (86.21 s), GUI-freier Core 1/1 (2.63 s), Debian-Prüfungen
und vier Originalfixture-Hashes bestanden. Keine vollständige Access-Parität.


## Privater Stand und Datenblatt-Tastenkürzel (2026-11-03)

Geprüftes Bereichseinfügen/Speichern: lokal 24854bf, privat remote 33981e2,
identischer Tree caca8e5ef7c1de87a10390a3b14d7272a700c895. Vollsuite 48/48.
Danach Strg+Apostroph im nativen Tabellen-View ergänzt: voriger Modell-Datensatz,
gleiches Feld, geprüfte Currency/Date/Boolean/Null-Übernahme nur in den Entwurf.
Keine SQL-Autospeicherung, keine Binär-/AutoNumber-Übernahme. Bestehende
MainWindow-/Currency-Targets gezielt bestanden; Details/Grenzen in
docs/KEYBOARD_RUNTIME.md. Gesamtsuite/Paket für diese Ergänzung nicht neu behaupten.


## Datenblatt-F2 und nativer Zoom (2026-11-03)

F2 wechselt Tabellen zwischen Einfügemarke und Feldnavigation, validiert nur
in den Entwurf und erhält ungültige Eingaben. Umschalt+F2 öffnet echten nativen
mehrzeiligen Zoom aus Tabellen-View oder Zelleditor. Abbrechen/unverändertes OK
erhalten Werte, Null und originale Zeilenenden; geänderte Texte bleiben im
Manual-Submit-Entwurf. Besitzer-/Indexlebensdauer im Dialog wird bewacht.
Formular-TextBox-Zoom ist im folgenden Inkrement ergänzt; vollständige
Windows-Darstellung bleibt offen.
Details docs/KEYBOARD_RUNTIME.md. Vollsuite 48/48 (86.44 s), Debian-Prüfungen
und vier Originalfixture-Hashes bestanden; GUI-freier Core zuletzt 1/1 (2.63 s),
danach ausschließlich Qt-Änderungen. Keine vollständige Access-Parität.


## Formular-Zoom und nativer VBA-Befehl (2026-11-03)

Umschalt+F2 verwendet nun denselben nativen Zoom für Formular-TextBoxes.
Eigener Feldpuffer, Change, Schreibschutz und abbrechbares BeforeUpdate bleiben
erhalten. Während des Dialogs wird nur das editingFinished dieses Feldes
unterdrückt, nicht der globale Loading-Zustand über die Ereignisschleife gesetzt.
DoCmd.RunCommand acCmdZoomBox=179 löst das aktive Dokument/Steuerelement auf;
unpassender Fokus ergibt 2046. Application.RunCommand unterstützt denselben
vorhandenen Befehlsumfang samt benanntem Command-Argument, unbekannte Optionen
weiterhin 3251. Bare RunCommand, alle AcCommand-Optionen und volle Zoom-Parität offen.
Bestehender Layout-Target 1/1 (14.90 s), GUI-freier Core 1/1 (1.53 s).
Letzte Vollsuite lokal 2fcaf98 / privater API-Stand ab5e6cd: 48/48 (86.44 s).
Für dieses Inkrement keine erneute Vollsuite behaupten. Details KEYBOARD_RUNTIME.md.


## SQL-Commit und verlustfreier Zahleneditor (2026-11-03)

Zwei echte Datenverlustpfade auf synthetischen Daten reproduziert und korrigiert:
Nach erfolgreichem Commit werden Entwürfe vor dem Neuladen entfernt; dessen
Fehlschlag kann beim erneuten Speichern keinen INSERT duplizieren. Fehlermeldung
nennt explizit den erfolgten Commit und den Ansichtsfehler. Vor-Commit-Fehler
erhalten weiterhin alle Entwürfe nach Rollback. Native REAL/FLOAT/DOUBLE-Zellen
verwenden max_digits10-Texte statt Qt-Spinbox mit zwei Dezimalstellen.
Unveränderte NULL-Felder erzeugen keine Spin-/Datums-Standardwerte; REAL,
Currency und Date durch tatsächliches SQL-Speichern geprüft. Grenzen in
docs/DATASHEET_CLIPBOARD_RUNTIME.md, exaktes Decimal/Windows-Format weiterhin offen.
Vollsuite 48/48 (88.41 s), Debian-Prüfungen und vier Originalfixture-Hashes
bestanden. GUI-freier Core 1/1 (1.53 s), danach nur Qt-Änderungen.


## Datenblattnavigation und Eingabeschutz (2026-11-03)

Strg+Ende lädt den unterstützten Tabellenumfang vollständig nach, statt am
Ende der ersten Qt-Ladeportion zu bleiben. Strg+Pos1 kehrt zum Anfang zurück.
Tab/Backtab/Umschalt+Tab/Return/Enter validieren native Editoren vor dem Schließen;
ungültige Werte bleiben sichtbar/fokussiert, gültige gehen nur in den Entwurf.
Gemeinsamer Commit-Helfer auch für Finish/F2. Maus-/FocusOut-Verhalten und volle
Access-Navigation weiterhin offen. Bestehende MainWindow-/Currency-Targets
2/2 (2.01 s), einschließlich 300 Datensätzen und unverändertem SQL bis Save.
Vollsuite/Paket zuletzt e15ff33 / privat c381897, 48/48 (88.41 s); diese kleine
Ergänzung ist gezielt geprüft. Details docs/KEYBOARD_RUNTIME.md.


## Mausvalidierung und Datenblatt-Kontextmenü (2026-11-03)

Normaler Editorabschluss validiert vor dem Verwerfen des Widgets. Ein ungültiger
Wert bleibt bei Mauswechsel erhalten, aktuelle Zelle/Fokus werden zurückgesetzt;
Escape/RevertModelCache bleibt bewusstes Verwerfen. Verlust vor Fix reproduziert.
Natives Tabellen-Kontextmenü: echtes Copy/Paste/Zoom/Datensatz-/Spalten-/Gesamtwahl
über die bestehenden geprüften Operationen. Dirty-Paste deaktiviert, Auswahl bei
Rechtsklick darin erhalten, Tastaturmenü an aktueller Zelle. Keine Filter-/
Cut-/vollständige Kontextmenü- oder Pixelparität behaupten. Vollständige
Fenster-/Fokuspolitik und externe Modell-Resets weiterhin offen.
MainWindow-/Currency-Targets 2/2 (2.25 s), keine neuen Targets. Vollsuite/Paket
zuletzt e15ff33 / privat c381897, 48/48 (88.41 s); diese View-Ergänzungen gezielt
geprüft. Details docs/KEYBOARD_RUNTIME.md.


## Echte Tabellen im Fenstermenü und Wechsel (2026-11-03)

Fenstermenü und Strg+F6/Strg+Umschalt+F6 nutzen im Produktions-MainWindow eine
Runde aus Datenbankfenster, offenen Tabellenseiten in Tab-Reihenfolge und
sichtbaren nativen Formularen/Reports. Andere Hosts behalten den nativen Fallback.
Tabellenfokus beendet den vorherigen aktiven Formular-Kontext; gültige Feldwerte
gehen nur in den Manual-Submit-Cache, ungültige blockieren Menü-/Tasten-/Tabklick-
Wechsel. Nur echte Tab-Navigationstasten werden vorab geprüft, keine reinen
Modifier-Tasten (vorübergehender Editor-Absturz damit korrigiert).
Native Form-/Modal-/BeforeUpdate-Pfade bleiben erhalten. Details und offene
MDI-/Query-/Editor-/Windows-Politik in docs/WINDOW_CYCLE_RUNTIME.md.
Gezielt 3/3 (16.31 s), Vollsuite 48/48 (92.38 s), Debian-Prüfungen und vier
Originalfixture-Hashes bestanden. GUI-freier Core 1/1 (1.53 s), danach nur Qt.
Vorherige Maus-/Kontextmenü-/Navigationsänderungen sind damit ebenfalls in
der Vollsuite geprüft. Privates Stacktrace-Werkzeug liegt nur außerhalb des
Repositories unter /workspace/linaccess-onboarding/private-backtrace.*.


## Bestehende Tabelle wiederverwenden / SelectObject (2026-11-03)

Erneutes Tabellenöffnen aktiviert dieselbe Seite/dasselbe Modell, erhält den
Entwurf und speichert SQL nicht automatisch. DoCmd.SelectObject acTable,Name,False
aktiviert tatsächlich geöffnete native Tabellen über den Produktions-MainWindow.
Quellvalidierung und Modalität bleiben wirksam; Me bezieht sich weiterhin auf
den lebenden Aufrufer, dessen Schließen die Tabelle nicht schließt. Lookup über
existierende Tabellenmetadaten, Groß-/Kleinschreibung gezielt geprüft.
Geschlossene Tabellen/fehlender UI-Host weiterhin explizit 3251; exakte Access-
Closed-Object-Fehler, OpenTable, Queryfenster und allgemeine Tabellen-DoCmd-
Befehle sind damit NICHT implementiert. Keine stillen Aktionen auf dem alten
Formular nach Tabellenaktivierung behaupten.
MainWindow/Layout 2/2 (15.56 s), keine neuen Targets. Vollsuite/Paket zuletzt
671d997 / privat fea77b0, 48/48 (92.38 s). Diese UI/Runtime-Ergänzung gezielt
geprüft. Details docs/DATABASE_WINDOW_RUNTIME.md und WINDOW_CYCLE_RUNTIME.md.


## Natives OpenTable / schreibgeschützte Tabellen (2026-11-03)

DoCmd.OpenTable öffnet echte Projekttabellen als normales Datenblatt mit
Standard/acEdit oder acReadOnly. Modellseitige Sperren verhindern Zell-, Checkbox-,
Insert-/Delete-/Paste-/Submit-Schreibzugriffe im Readonly-View. Vorhandene dirty
Entwürfe können nicht durch einen Policywechsel verworfen werden. Native
Objektargumente unterstützen benannte/ausgelassene Positionen; Missing ist ein
Default, explizites Empty wird für Enums zu 0 und nicht zum Edit-Default.
Design/Preview/acAdd und widersprechender Modus eines bereits offenen Views
bleiben 3251. Readonly ist View-Politik, keine globale DAO-/Dateisperre. Allgemeine
Tabellen-DoCmd-Befehle und volle Access-Fehler-/Fensterparität weiterhin offen.
Kanonische Spezifikation: docs/TABLE_DOCUMENT_RUNTIME.md (inkl. Access-2003-Referenz).
Vollsuite 48/48 (92.13 s), GUI-freier Core 1/1 (1.68 s), Debian-Prüfungen und
vier Originalfixture-Hashes bestanden. Keine neuen Testtargets.


## SaveRecord für aktive Tabellen (2026-11-03)

DoCmd.RunCommand acCmdSaveRecord sowie Application.RunCommand nutzen jetzt den
gemeinsamen nativen Tabellen-Speicherpfad (inkl. UI-Button, Transaktion und
PK-Auswahlwiederherstellung). Der direkte Interpreter-Save-Pfad ist ebenfalls
angebunden. Readonly-Save ergibt 2046, unbekannte neue Tabellen-Dispatcherbefehle
3251; kein Rückfall auf das alte Aufruferformular. Programmpfad ohne zusätzliche
SQL-Dialoge, native Fehlermeldungen werden zu VBA zurückgegeben.
Nie aktivierte Formulare behalten ihren Initialisierungskontext, sofern sie
nicht explizit eine Tabelle geöffnet/selektiert haben. Explizite Tabellenkontexte
werden getrennt von echten Formular-Activate-Ereignissen markiert. Dirty-Hintergrund-
tabelle blieb beim Konstruktor-Current-Save unverändert, Me bleibt der Aufrufer.
Details docs/TABLE_DOCUMENT_RUNTIME.md. Trio 3/3 (16.13 s), letzter Initfall 1/1
(1.93 s). Vollsuite/Paket zuletzt e054910 / privat 0328119, 48/48 (92.13 s).
Dieses Qt-Controller-Inkrement ist gezielt geprüft, keine neuen Targets.


## Native Tabellen schließen (2026-11-03)

Explizites/default DoCmd.Close schließt echte Tabellen nach Feldvalidierung und
transaktionalem Datensatzspeichern. Auch acSaveNo verwirft keine Datensatzdaten;
Designspeichern und vollständige Access-Fehlerpriorität bleiben offen. Fehler
lassen Seite, Editor und Cache bestehen. Verdeckte aktive Zellpuffer bleiben
prüfbar; reentrantes Schließen beendet das zugehörige native Kontextmenü sicher.
Details docs/TABLE_DOCUMENT_RUNTIME.md. Vollsuite 48/48 (91.15 s), GUI-freier
Core 1/1 (2.01 s), Debian-Paketprüfung und alle vier Fixture-Hashes bestanden.
Damit ist auch der vorherige SaveRecord-Controller vollständig mitgeprüft.
Keine neuen Testtargets. Privater API-Push mit Baumidentität und Ref-Lease;
lokale/API-Metadatenhistorien getrennt erhalten, kein Force-Push.


## Tabellen-Auswahl und Kopieren (2026-11-03)

acCmdSelectRecord und acCmdCopy adressieren die echte aktive Tabellenansicht
über den nativen Dispatcher (DoCmd/Application). Schreibgeschützte Zeilen sind
kopierbar; schmutzige Cachewerte werden kopiert, ohne SQL zu speichern. Ungültige
Zelleneingaben blockieren Auswahl/Kopie und erhalten die Zwischenablage.
Details docs/TABLE_DOCUMENT_RUNTIME.md; MainWindow-Regression 1/1 (1.97 s).
Letzte Vollsuite/Paket: lokal 2dd59b2 / privat 32cb556, 48/48 (91.15 s),
GUI-freier Core 1/1 (2.01 s). Dieses Inkrement ist gezielt geprüft.


## Nativer Tabellen-Zoombefehl (2026-11-03)

acCmdZoomBox arbeitet auf der tatsächlichen aktiven Tabellenzelle, mit nativer
Qt-Ansicht, typgeprüftem Cache und Readonly-Anzeige. Abbrechen/Readonly-Schließen
sind erfolgreiche Befehlsrückkehr ohne Änderung. SelectRecord erhält die aktive
Spalte; Qt verschob sie zuvor auf Spalte 0 (über Zoom reproduziert, korrigiert).
Details docs/TABLE_DOCUMENT_RUNTIME.md. MainWindow/Currency 2/2 (2.47 s).
Vollsuite/Paket weiterhin 2dd59b2 / privat 32cb556, 48/48 (91.15 s); dieses
Qt-Inkrement gezielt geprüft, keine neuen Targets.


## Native Tabellen-Datensatznavigation (2026-11-03)

GoToRecord für benannte/aktive Tabellen und RunCommand First/Previous/Next/Last
navigieren reale SQL-Zeilen in View-Reihenfolge; Last/GoTo lädt über Qt-Batches
hinaus (1-Million-Grenze). Ein allein geänderter bestehender aktueller Datensatz
mit Primärschlüssel wird vor dem Wechsel transaktional gespeichert. SQL-Fehler
lassen Zeile/Cache zum Korrigieren/Retry bestehen. Fremde Entwürfe, Inserts/Deletes,
fehlende sichere Identität und neue Platzhalterzeilen bleiben explizit begrenzt.
Details docs/TABLE_DOCUMENT_RUNTIME.md. Trio 3/3 (16.79 s), Vollsuite 48/48
(91.33 s), GUI-freier Core 1/1 (2.49 s), Debian-Paket und vier Fixture-Hashes
bestanden. Keine neuen Testtargets; auch vorherige Tabellenbefehle voll mitgeprüft.


## Native Tabellen-Sortierbefehle (2026-11-03)

RunCommand SortAscending/SortDescending sortiert das aktuelle Tabellenfeld über
Qt/SQL. Primärschlüssel-Datensatz und Spalte bleiben nach begrenztem Nachladen
erhalten. Readonly-Sortierung und geprüftes Speichern eines allein geänderten
aktuellen Datensatzes funktionieren; fremde Entwürfe bleiben geschützt.
Details docs/TABLE_DOCUMENT_RUNTIME.md, MainWindow 1/1 (2.08 s). Vollsuite/Paket
zuletzt e4e0900 / privat 714a96e, 48/48 (91.33 s), GUI-frei 1/1 (2.49 s).
Dieses Controller-Inkrement ist gezielt geprüft; keine neuen Testtargets.


## Native Tabellen-Textsuche (2026-11-03)

Echter modeless Qt-Suchdialog, Strg+F (auch Zell-Editor), Umschalt+F4, Bearbeiten-
Menü und DoCmd/Application Find/FindNext sind an reale Tabellen angeschlossen.
Wörtliche sichtbare Textfelder mit Teil/Ganz/Anfang und Groß-/Kleinschreibung;
Nachladen bis 1 Mio Textzellen, Umlauf in tatsächlicher View-Reihenfolge.
Geprüftes Speichern eines allein geänderten aktuellen bestehenden Datensatzes;
weitere Entwürfe bleiben geschützt. Dialog aktiviert seine Tabelle bei Bedarf
und wird mit ihr entsorgt. Reentranz/Lebensdauer mit schwachen Besitzern gesichert.
Ansichtsfehler setzen jetzt den tatsächlichen nativen Modellfehler für VBA;
Binär-Copy bleibt 3251, Zwischenablage und Payload bleiben unverändert.
Kanonisch docs/TABLE_FIND_RUNTIME.md; Zahlen/Datum/Format/Wildcards/Locale/Replace
und DoCmd.FindRecord nicht von diesem Inkrement abgedeckt.
Vollsuite 48/48 (92.69 s), GUI-freier Core 1/1 (2.45 s), Debian-Paket und vier
Fixture-Hashes bestanden, keine neuen Testtargets. Enthält auch Sortierbefehle.


## Nativer DoCmd.FindRecord-Teilumfang (2026-11-03)

FindRecord durchsucht das aktuelle gebundene Textfeld tatsächlicher aktiver
Tabellen/Formulare mit SearchAll, Literaltext, Teil/Ganz/Anfang, Groß-/Kleinschreibung
und erstem/folgendem Suchbeginn. Benannte/ausgelassene Argumente sind Missing;
explizites Empty wird normal koerziert (Match=0, FindFirst=False), nicht ersetzt.
Readonly-Felder sind suchbar, kein Treffer ist normale Rückkehr. Formvalidierungs-
abbruch wird 2501 vor nachfolgendem VBA, Entwurf und aufrufendes Me bleiben erhalten.
Expliziter Tabellenkontext verhindert auch beim nie aktivierten Aufrufer einen
Rückfall auf das alte Formular aus dem Datenbankfenster. Initialisierung ohne
expliziten Dokumentwechsel bleibt unverändert im eigenen Kontext.
Allfelder/Zahlen/Datum/Format/Richtung/Wildcards/volle Access-Fehlerpriorität offen.
Kanonisch docs/FIND_RECORD_RUNTIME.md mit geprüften Microsoft-Parameterreferenzen.
Trio 3/3 (15.83 s), Vollsuite 48/48 (88.05 s), GUI-freier Core 1/1 (1.99 s),
Debian-Paket und vier Originalfixture-Hashes bestanden. Keine neuen Testtargets.


## Suchrichtungen auf echten Tabellen/Formularen (2026-11-03)

Dialog und FindRecord unterstützen jetzt Up/Down ohne Umlauf sowie SearchAll mit
Umlauf. FindNext erhält die Richtung; FindFirst beginnt beim ersten Datensatz,
sonst nach dem aktuellen Datensatz in Suchrichtung. Gegenüberliegende Grenzen,
UI-Wechsel und benannte VBA-Richtungen geprüft. Trio 3/3 (16.70 s); Vollsuite/Paket
zuletzt a71acc9 / privat f9367b0, 48/48 (88.05 s), GUI-freier Core 1/1 (1.99 s).
Dieses Qt-Inkrement ist gezielt geprüft, keine neuen Targets. Kanonisch
FIND_RECORD_RUNTIME.md und TABLE_FIND_RUNTIME.md; übrige Suchgrenzen bleiben offen.
