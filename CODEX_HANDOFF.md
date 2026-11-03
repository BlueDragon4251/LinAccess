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


## Typisierte native Suche (2026-11-03)

Tabellen-/Formulardialoge und FindRecord durchsuchen echte Felder als Number,
Currency, Date und Boolean. Vergleich erhält benachbarte Double-Werte, exakte
Currency-Einheiten, kanonische native Datumswerte und NULL-versus-False. Geprüfte
QVariant-Weitergabe ohne Currency-Float-/JSON-Zwischenwert. Datums-Negativalias
läuft durch die tatsächliche Date-Speicherpipeline; Subsekunden explizit begrenzt.
Einzelne typisierte Felder erzwingen Ganzes Feld und zeigen Eingabehinweise;
Text-Gesamtsuche bleibt Textscope. Typkonversionen zwischen inkompatiblen Feldern,
Teiltreffer nichttextueller Werte, Allfelder-API/Format/Locale/Wildcards bleiben offen.
Readonly-Suche, ungültige Datumseingabe, Caller-Me, FindNext und unveränderte
SQL-Werte/NULLs/Binärdaten geprüft. Kanonisch docs/TYPED_FIND_RUNTIME.md.
Trio 3/3 (2.79 s), Vollsuite 48/48 (97.33 s), GUI-freier Core 1/1 (1.86 s),
Debian-Paket und vier Originalfixture-Hashes bestanden; keine neuen Testtargets.

## Native Tabellen-Tasten und DoCmd.SendKeys (2026-11-03)

Echte Zell-/Find-/Zoom-Fokussierung pro Taste, unveränderliche Projektidentität,
benannte/ausgelassene DoCmd.SendKeys-Argumente und Wait=False-Standard. FIFO-Queue
setzt innerhalb modaler Keyhandler fort; Wait=True wartet, DoEvents leert mit
interaktiver Qt-Verarbeitung. Text bleibt bis Validierung vom gebundenen Value
getrennt; Save erreicht echtes SQL. Keine Rückleitung auf alte Aufrufer bei
Datenbankfenster-Fokus. Kanonisch docs/TABLE_KEYBOARD_RUNTIME.md. Originaler
Wait=False-Probe wartet jetzt auf die Queue statt auf genau einen Qt-Durchlauf;
Empfänger-, Inhalts- und Ereignisprüfungen bleiben bestehen. Paar 2/2 (17.52 s),
originaler Kassen-Ereignislauf 1/1 (38.68 s), keine neuen Testtargets.

Vollsuite 48/48 (93.79 s), GUI-freier Core 1/1 (2.27 s), Debian-Paketprüfung
inklusive Build-ID/SQLite 3.54.1 sowie alle vier Originalfixture-Hashes bestanden.

## Native SelectAllRecords (2026-11-03)

DoCmd/Application acCmdSelectAllRecords (109) und Strg+A verwenden denselben
geprüften Tabellenpfad mit vollständigem begrenztem SQL-Nachladen, Feldvalidierung
und Erhalt der aktuellen Zelle ohne SQL-Speichern. Bestehender MainWindow-Test
1/1 (3.02 s), GUI-freier Core 1/1 (1.68 s); keine neuen Targets. Kanonisch
docs/TABLE_DOCUMENT_RUNTIME.md. Letzte Vollsuite/Paket-Quelle: lokal 934c26d /
privat 914cc78, 48/48 (93.79 s), Paket und vier Fixture-Hashes bestanden.

## Nativer Tabellen-Paste-Befehl (2026-11-03)

DoCmd/Application acCmdPaste (191) läuft durch denselben echten TSV-Pfad wie
Strg+V: vollständige Typvorprüfung vor Cachemutation, Schreibschutz 2448,
vorhandene Entwürfe 2501, kein implizites SQL-Speichern. Fehlendes Text-MIME
setzt jetzt frischen 2046 statt einen alten Modellfehler weiterzureichen.
Ungültige Folgezelle atomar abgelehnt, echter Zweizeilenentwurf und unveränderte
SQL-Werte geprüft. MainWindow 1/1 (3.12 s), GUI-freier Core 1/1 (1.63 s),
keine neuen Targets. Kanonisch docs/TABLE_DOCUMENT_RUNTIME.md; Texteditor-Routing,
Entwurfszusammenführung, Record-Append/Binär/Locale bleiben begrenzt/offen.

Vollsuite 48/48 (97.38 s), GUI-freier Core 1/1 (1.63 s), Debian-Paketprüfung
inklusive Build-ID/SQLite 3.54.1 sowie vier Originalfixture-Hashes bestanden.
Dies schließt das vorherige SelectAllRecords-Inkrement ein.

## Tabellen-Neuabfrage und Datensatz-Speichertasten (2026-11-03)

DoCmd.Requery ohne ControlName und Umschalt+F9 laden echte SQL-Quelldaten aktiver
Tabellen neu. Geprüfter aktueller Einzelentwurf wird vorher transaktional
gespeichert, weitere/Insert-/Delete-Entwürfe bleiben mit 3251 geschützt. Danach
erster Ergebnisdatensatz, bestehende Spalte/Sortierung bleiben; allgemeine
Access-Positions-/Event-/Fehlerpriorität und benannte Controls offen. Umschalt+
Return sowie Umschalt+Nummernblock-Enter nutzen echten Tabellen-Save auch aus
dem Zell-Editor, inklusive Validierung und PK-/Spaltenerhalt. MainWindow 1/1
(3.12 s), keine neuen Targets; GUI-freier Core unverändert 1/1 (1.63 s), nur Qt-
Pfade geändert. Kanonisch docs/TABLE_DOCUMENT_RUNTIME.md / KEYBOARD_RUNTIME.md.
Letzte Vollsuite/Paket-Quelle lokal 375bd82 / privat 89bbcb6: 48/48 (97.38 s),
Debian-Paket und vier Originalfixture-Hashes bestanden.

Bestehendes MainWindow/Layout-Paar 2/2 (17.72 s) bestanden, einschließlich
gemeinsamem Formularhost-Requery und Tabellen-Speichertasten; keine neuen Targets.

## Tabellen-Datensatzbefehle und echte Insert-Schlüssel (2026-11-03)

RunCommand 28/223 verwendet echte Insert-/Delete-Entwürfe und transaktionales SQL
beim Save. Mehrfach-Löschwahl vor Mutation abgelehnt, weitere Insert-Entwürfe
geschützt, readonly 2448; SelectAllRecords benötigt nach Reset keine aktuelle
Zelle mehr. SQLite-INSERT ermittelt innerhalb der Transaktion den tatsächlichen
PK über unverdeckten Rowid-Alias/last_insert_rowid; erst nach Commit veröffentlicht,
Save stellt echte Zeile/Spalte wieder her. Kein MAX/geschätzter Zähler.
WITHOUT ROWID/alle Rowid-Aliase verdeckt: kein erfundenes Insert-Bookmark,
SQL-Erfolg bleibt erhalten. Staged Tabellen-Delete ist keine Behauptung über
sofortigen Access-Commit/Bestätigung/Events. Kanonisch TABLE_DOCUMENT_RUNTIME.md.
Bestehendes MainWindow/Currency-Insert-Paar 2/2 (3.35 s), keine neuen Targets.

Finale Vollsuite 48/48 (91.07 s), GUI-freier Core 1/1 (2.38 s), Debian-Prüfung
inklusive Build-ID/SQLite 3.54.1 sowie vier Fixture-Hashes bestanden. Unterdrückter
INSERT veröffentlicht keinen alten Schlüssel; erster Datensatz einer leeren
Tabelle samt erzeugtem PK/Spaltenerhalt geprüft. Voriges Requery/Speichertasten-
Inkrement eingeschlossen, keine neuen Testtargets.

## Explizites Tabellen-acNewRec und vollständiges Nachladen (2026-11-03)

Benanntes/aktives DoCmd.GoToRecord acNewRec nutzt den echten Einfügepfad mit
Seitenaktivierung, leerer Tabelle, readonly 2448 und Insert-Entwurfsschutz 3251.
Tabellen-Insert lädt vorhandene SQL-Batches begrenzt vollständig nach, dann
kommt der neue Entwurf tatsächlich hinter alle vorhandenen Zeilen; Button/
Strg+Plus teilen den Pfad. 300 teilweise geladene Zeilen -> Entwurfszeile 301,
SQL weiterhin 300. Kein implizites Next-hinter-Ende oder volle Addition-Row-Parität.
MainWindow 1/1 (3.53 s), keine neuen Targets. Kanonisch TABLE_DOCUMENT_RUNTIME.md.
Letzte Vollsuite/Paket-Quelle lokal a93e2ed / privat cb6a3ea: 48/48 (91.07 s),
GUI-freier Core 1/1 (2.38 s), Debian-Prüfung und vier Fixture-Hashes bestanden.

## Fokussiertes Tabellen-Paste (2026-11-03)

acCmdPaste prüft tatsächlichen QLineEdit-Fokus innerhalb der aktiven Tabelle
oder ihres Find-Dialogs vor dem TSV-Gridpfad. Echte Textauswahl/Cursorposition,
Textpuffer vor Modellvalidierung/SQL-Save; Find-Eingabe trotz readonly Quelle
bearbeitbar. Andere Formulare/Projekte sind keine Ersatzempfänger. Vorhandene
Zwischenablagegrenzen/Readonly-/kein-Text-Fehler bleiben; allgemeines Multiline-/
Richtext-Routing offen. Kanonisch docs/TABLE_DOCUMENT_RUNTIME.md. MainWindow
1/1 (3.17 s), keine neuen Targets; vorheriges NewRec/Fetch-Inkrement auch im
MainWindow/Currency-Feld/Vorwert-Trio 3/3 (4.65 s) bestanden.

Vollsuite 48/48 (91.66 s), GUI-freier Core 1/1 (1.96 s), Debian-Prüfung
inklusive Build-ID/SQLite 3.54.1 und alle vier Originalfixture-Hashes bestanden.
Benanntes/aktives acNewRec, vollständiges Nachladen und tatsächliches Zell-/
Readonly-Find-Text-Paste eingeschlossen; keine neuen Testtargets.

## Fokussierte Editorbefehle und Einfg-Tasten (2026-11-03)

Tabellen-Copy/Cut 190/189 nutzen echte aktive QLineEdit-Auswahl, auch im Find-
Dialog einer readonly Quelle. Ungültiger Zahlentext bleibt kopierbar, aber
SelectRecord/Save lehnt seine Validierung weiter ab; fehlende Auswahl 2046 lässt
Zwischenablage unverändert. Undo/Redo 292/199 verwenden den tatsächlichen
Editorverlauf ohne vorzeitige Feld-/SQL-Übernahme. Maskierte/disabled/readonly
Felder und fehlender Verlauf begrenzt; Grid-Cut/-Undo/-Redo sowie allgemeines
Richtext/Multiline/Form-Routing bleiben offen. Strg+Einfg/Umschalt+Einfg teilen
geprüftes Grid-Copy/Paste. Kanonisch TABLE_DOCUMENT_RUNTIME.md / KEYBOARD_RUNTIME.md.
Bestehendes MainWindow/Currency-Feld/Vorwert-Trio 3/3 (4.25 s), GUI-freier Core
1/1 (1.71 s), keine neuen Targets.

Vollsuite 48/48 (93.40 s), GUI-freier Core 1/1 (1.71 s), Debian-Prüfung
inklusive Build-ID/SQLite 3.54.1 und vier Originalfixture-Hashes bestanden.
Fokussiertes Copy/Cut/Undo/Redo und Einfg-Aliase eingeschlossen, keine neuen Targets.

## Native CurrentProject-Dateimetadaten (2026-11-03)

CurrentProject und Application.CurrentProject liefern tatsächlichen nativen
Dateinamen, absolutes Verzeichnis/Dateipfad als readonly Strings. Schwacher
Identitätscache pro Verbindung/Datei und MainWindow-Generation; kein Halt des
aufrufenden Formulars. Nach dessen Schließen bleibt Referenz gültig, geschlossene
Verbindung/Owner oder Projektwechsel ergeben 3420, auch erneutes Öffnen derselben
Datei. Abgebrochener Unload hält Generation/Referenz. Fremde Preview-Verbindungen
unter gleichem Fenster werden getrennt. Memory/URI/andere Treiber keine erfundenen
Dateipfade (3251), Schreiben/andere Member 438, Argumente 450. Benutzer-Lokalnamen
und Funktionen namens CurrentProject bleiben auflösbar. Kanonisch
docs/CURRENT_PROJECT_RUNTIME.md; komplette COM-/Katalog-APIs bleiben offen.
Bestehendes Core/MainWindow-Paar 2/2 (4.57 s), keine neuen Testtargets.

Finale Vollsuite 48/48 (92.20 s), GUI-freier Core 1/1 (1.56 s), Debian-Prüfung
inklusive Build-ID/SQLite 3.54.1 und vier Originalfixture-Hashes bestanden.
MainWindow hält absoluten Öffnungspfad fest; standalone relative Dateinamen
ergeben keine geratenen Metadaten. Keine neuen Testtargets.

## Projektkataloge und Report.Visible (2026-11-03)

CurrentProject.AllForms/AllReports lesen gespeicherte native/importierte Objekte
aus echtem SQL: Long Count, Name/Index/named Item, For Each und unabhängige
readonly AccessObject-Deskriptoren. IsLoaded erkennt versteckte/modale/embedded
Previews, geschlossene Eltern und tatsächliche native Runner/Designer-Dialoge.
Mehrdeutige Namen 5 statt willkürlicher Wahl; Version/Sortierung/Snapshots und
allgemeine COM-/Eventzeitpunkt-Grenzen dokumentiert. Report.Visible nutzt nun
geprüfte echte Qt-Sichtbarkeit; Me.Visible aus HideSelf/ShowSelf setzt Code fort,
IsLoaded bleibt beim Verbergen wahr, Report_Activate/Deactivate erzeugen echte
SQL-Auditzeilen. Kanonisch PROJECT_CATALOG_RUNTIME.md / FORM_VISIBILITY_RUNTIME.md.
Bestehendes MainWindow/Form-Visibility/Layout-Trio 3/3 (18.72 s), keine neuen Targets.

Finale Vollsuite 48/48 (97.60 s), GUI-freier Core 1/1 (1.71 s), Debian-Prüfung
inklusive Build-ID/SQLite 3.54.1 und vier Fixture-Hashes bestanden.
Gespeichertes Hauptschema/Temp-Isolation, echtes leer/kein-Katalog, Katalogleben
und Report.Visible mit VBA/SQL eingeschlossen. Keine neuen Testtargets.

## Native CurrentData-Tabellen-/Abfragekataloge (2026-11-03)

CurrentData/Application.CurrentData liefern tatsächliche Hauptschema-Nutzertabellen
und gespeicherte native/importierte Abfragedefinitionen über Long Count,
Name/Index/named Item, For Each und readonly Name/IsLoaded. Interne Tabellen
ausgeschlossen; physische Views nicht als Tabellen und erzeugte Query-Views
nicht als zusätzliche gespeicherte Abfragen ausgegeben. Echte Tab-Mitgliedschaft
hält Hintergrund-/readonly-Datenblätter geladen, entfernte Seiten sind geschlossen.
Benannte Query-Editor-/Quell-/Ergebnisdialoge werden tatsächlich verfolgt.
Referenzen überleben Aufrufer, Projektgeneration invalidiert sie. Benutzerfunktion
CurrentData bleibt auflösbar. Kanonisch docs/CURRENT_DATA_RUNTIME.md; keine
pauschale Query-/COM-Parität. Core/MainWindow-Paar 2/2 (6.16 s), GUI-freier
Core 1/1 (1.51 s), keine neuen Targets.

Vollsuite 48/48 (101.66 s), GUI-freier Core 1/1 (1.51 s), Debian-Prüfung
inklusive Build-ID/SQLite 3.54.1 und vier Originalfixture-Hashes bestanden.
CurrentData-Metadaten/echter UI-Zustand eingeschlossen; keine neuen Testtargets.

## Echte readonly Abfragedokumente (2026-11-03)

DoCmd.OpenQuery öffnet native/importierte gespeicherte SELECTs als echte Tabs
mit ausdrücklich acReadOnly/Normalansicht, tatsächlichen SQL-Ergebnissen und
typisierten Parameterdialogen. Missing/default acEdit bleibt unverändert,
bisher nicht unterstütztes Editieren wird 3251 statt still readonly; Aktions-/
MakeTable-/andere Ansichten vor Ausführung abgelehnt. Parameterabbruch 2501
erzeugt kein Dokument, Currency-Konversion am Ergebnis geprüft. Root-eigener
Tab überlebt aufrufendes Formular, dessen Me bleibt erhalten. Wiederverwendung,
benanntes/default Query-Close, Strg+F4 und CurrentData.IsLoaded folgen echter
Tab-Lebensdauer. Fensterwechsel fokussiert die tatsächliche Ergebnistabelle.
Kanonisch docs/OPEN_QUERY_RUNTIME.md; allgemeine Query-Befehle/Updateability,
Requery/SelectObject und vollständige Access-UI-/Fehlerpriorität bleiben offen.
MainWindow gezielt 1/1 (3.35 s), keine neuen Testtargets.

Vollsuite 48/48 (90.08 s), GUI-freier Core 1/1 (1.58 s), Debian-Prüfung
inklusive Build-ID/SQLite 3.54.1 und vier Originalfixture-Hashes bestanden.
Parameterloses Close schließt aktive readonly Abfrage statt Aufrufer. Keine neuen Targets.

## Schreibgeschützte Abfragen neu laden (2026-11-03)

Aktive Abfrage-Requery und Shift+F9 laden Definition/Parameter neu und ersetzen
das echte SQL-Ergebnis erst nach Erfolg. Fehler/Abbruch behalten das alte Modell.
Ein echter SQLite-Callback prüft rekursive Ausführung (3251), Projektwechsel-
Sperre und Zerstörung des alten Tabs während SQL, ohne Wiedereröffnung.
Gezieltes bestehendes Mainwindow-Ziel 1/1 (3,44 s); kein neues Testziel.
Details/Grenzen: docs/OPEN_QUERY_RUNTIME.md. 332 geprüfte Checklist-Punkte,
72 offene Punkte; keine Gesamtparitätsquote. Vollprüfung: Desktop 48/48
(89,91 s), GUI-frei 1/1 (1,57 s), Debian-Paket und alle vier Original-SHA-256
bestanden.

## Native Abfrage-Navigation und Zwischenablage (2026-11-03)

GoToRecord auf aktivem Abfrageergebnis und ausdrücklich acDataQuery (auch benannt),
RunCommand erster/vorheriger/nächster/letzter Datensatz, SelectObject acQuery,
SelectRecord/SelectAllRecords, Copy/Strg+C/Strg+Einfg und projektgebundene SendKeys
arbeiten auf tatsächlichen Resultat-Tabs. Grenzen 2105, readonly-Neuanlage 2448;
explizite Tabellenziele werden nicht auf Abfragen umgedeutet. TSV-Aufbereitung
mit Tabellen geteilt, echte Typkonvertierung, kein Binär-/Hidden-/Disjoint-Export.
Umfang/Grenzen: docs/OPEN_QUERY_RUNTIME.md. Vollprüfung 48/48 (91,62 s), GUI-frei
1/1 (1,54 s), Debian-Paket und vier Original-SHA-256 bestanden; keine neuen
Testziele. 336 erfasste erledigte / 72 offene Checklist-Punkte, keine Gesamtquote.

## Native Abfragesuche (2026-11-03)

Abfragefenster: echter modelloser Suchdialog, Strg+F/Umschalt+F4, Find/FindNext
und aktuelles Feld über DoCmd.FindRecord. Typisierte Vergleiche verwenden
nachgewiesene Projektions-/physische Spaltentypen; separate positionsbezogene
UI-Metadaten lassen den bisherigen DAO-Typvertrag bestehen. Unbekannte Ausdrücke
werden nicht anhand angezeigter Werte typisiert. Suchtext-Bearbeitung nutzt
geteilte echte QLineEdit-Befehle. Abfrageöffnung wählt erste Ergebniszelle sofort,
damit direkt folgender Currency-FindRecord ohne Qt-Zwischenschritt funktioniert.
Ergebnistausch prüft Lebensdauer nach Qt-Callbacks. Umfang/Grenzen stehen in
OPEN_QUERY_RUNTIME.md. Desktop 48/48 (90,46 s), GUI-frei 1/1 (1,50 s), Debian-
Paket und vier Original-SHA-256 bestanden. 337 erledigte / 72 offene erfasste
Checklist-Punkte, kein Gesamtparitätsprozentsatz und keine neuen Testziele.

## Abfrage-Bearbeiten-Menü und Zoom (2026-11-03)

Bearbeiten erkennt echte Abfrageansichten: Kopieren/Alles markieren leiten
Tastenereignisse an die View; schreibende Grid-Aktionen bleiben deaktiviert.
Copy wird direkt verarbeitet, gleiche geprüfte Formatierung; Zellgrenze wird
vor der Index-Allokation geprüft. acCmdZoomBox/Umschalt+F2 öffnen den echten
schreibgeschützten Qt-Zoom mit unveränderten SQL-/Modellwerten nach Schließen.
Gezielt Mainwindow/Layout 2/2 (17,70 s), Debian-Paket und vier Original-SHA-256
bestanden. Keine neuen Testziele; letzte vollständige 48/48-Suite gehört zum
Abfragesuche-Checkpoint (90,46 s). 339 erledigte / 72 offene erfasste Punkte,
kein Gesamtprozentsatz. Grenzen: OPEN_QUERY_RUNTIME.md.

## Dreiteilige Abfrageherkunft und Originalkorpus (2026-11-03)

query.relation.field wird über nachgewiesene direkte Projektionen der echten
Temp-/Main-View-Definitionen an eindeutige Ausgabespalten gebunden, auch durch
TOP-Rang-Wrapper und bei zwei Eingangsviews derselben physischen Tabelle.
Keine Wahl bei Mehrdeutigkeit; UNION-Physikmetadaten gelten nicht als vollständiger
Herkunftsnachweis. Maximal 16 Ebenen, keine Auswertung von Werteausdrücken.
QUERY_ORIGINS.md ist verbindlicher Umfang. Ein weiteres unverändertes Original
läuft: jetzt 78/100 (73 ohne zusätzliche Parameter, fünf mit datengeprüften
Testparametern), weiterhin kein Access-Ergebnisoracle. Verbleibend 9 fehlende
Abhängigkeiten, 8 Bindungs-/Kontextfälle, 3 gegen Originalquellen/-schema geprüfte
fehlende Spalten, 2 geschützte Joined-Aktionen. Keine Ersatzspalten erfinden.
Vollsuite 48/48 (90,47 s), GUI-frei 1/1 (1,67 s), Debian-Paket und alle vier
Original-SHA-256 bestanden; keine neuen Testziele. 340 erledigte / 72 offene
erfasste Checklist-Punkte, keine Gesamtparitätsquote.

## Mehrdimensionale Variants und DAO GetRows (2026-11-03)

Native Variant-Arrays speichern Rang/Dimensionsgrenzen und unterstützen echte
mehrdimensionale Lese-/Schreibindizes, Long-LBound/UBound mit Dimension,
Wertkopien und For Each (erste Dimension läuft am schnellsten). ReDim bis 60
Variant-Dimensionen; Preserve darf nur letzte obere Grenze ändern. Join/Filter
bleiben ausdrücklich eindimensional; typisierte/feste/dynamische Array-
Deklarationen und vollständige COM-/SAFEARRAY-Parität sind weiterhin offen.
DAO GetRows mit explizit positiver NumRows liefert echte nullbasierte Feld-/
Datensatzmatrizen aus SQL-Daten, typisierte Currency/Date/Boolean/Null-Werte und
Byte-Array-Zellen für native SQLite-BLOBs. Konvertierung ist vor Cursorfortschritt
abgeschlossen. Omitted/Missing/nichtpositive NumRows bleiben 3251: den optionalen
DAO-Default nicht aus ADO ableiten. Keine MDB/OLE-Importbehauptung. Details:
DAO_GETROWS_RUNTIME.md und VARIANT_ARRAY_RUNTIME.md. Gesamtsuite 48/48 (92,36 s),
frisch GUI-frei 1/1 (1,50 s), Debian-Paket und vier Original-SHA-256 bestanden.
343 erledigte / 72 offene erfasste Checklist-Punkte, keine Gesamtparitätsquote.

## Typisiertes ReDim in Variant-Containern (2026-11-03)

Vorhandene Variant-Variablen akzeptieren ReDim ... As Byte/Integer/Long/Single/
Double/Currency/Date/Boolean/String/Variant mit echten typisierten Anfangswerten,
geprüften Elementzuweisungen und passendem VarType/TypeName. Ohne As bleibt der
Elementtyp erhalten; Non-Preserve initialisiert neu, Preserve erhält Werte/Typ
und die bekannten Dimensionsgrenzen. Neuer typisierter Variablen-Declaration-
Pfad bleibt 3251: feste/dynamische/typisierte Deklarationen nicht als erledigt
behaupten. Object/UDT/feste Strings und vollständige Byte-Enumeration/Erase sind
offen. Numerische typisierte Join/Filter-Pfade bleiben ausdrücklich 3251.
VARIANT_ARRAY_RUNTIME.md ist der Umfang. Vollsuite 48/48 (89,81 s), frisch
GUI-frei 1/1 (1,45 s), Debian-Paket und vier Original-SHA-256 bestanden;
keine neuen Testziele. 344 erledigte / 72 offene erfasste Punkte, keine Gesamtquote.
