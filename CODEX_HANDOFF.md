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

## Objektverweise in Arrays und Byte-Erase (2026-11-03)

Set array(indices)=object/Nothing speichert echte native Referenzen. Arraykopien
haben unabhängige Einträge bei gemeinsam referenzierten Objektinstanzen; bestehende
Variant-Variablen unterstützen jetzt ReDim As Object mit Nothing, VarType 8201,
Object() und Preserve. Echte DAO-Objekte aus Arrays führen geprüfte SQL-/Cursor-
Methoden aus. Skalare Set-Werte scheitern 13, mehrdeutige/ungültige Indizes bleiben
geprüft; implizite Object-Array-Defaultmember-Schreibzugriffe bleiben 3251.
Explizite Objektmember nutzen die vorhandene Bindung/Lebensdauer. Erase von
Byte-Arrays in Variant-Variablen setzt Empty, inklusive unterstütztem ByRef und
unabhängigen Wertkopien. Deklarations-/Class-Array-/COM-Grenzen bleiben offen.
VARIANT_ARRAY_RUNTIME.md beschreibt den Umfang. Vollsuite 48/48 (92,45 s), frisch
GUI-frei 1/1 (1,79 s), Debian-Paket und vier Original-SHA-256 bestanden.
Keine neuen Testziele. 346 erledigte / 72 offene erfasste Punkte, keine Gesamtquote.


## Native Objektbindung und explizites Static (2026-11-04)

Native Scripting.Dictionary mit geprueften Schluesseln, Snapshots und
Enumerationslebensdauer ist implementiert; Standard-Indexschreiben auf Objekte
und DAO-Felder fuehrt echte gepufferte Feld-/SQL-Aenderungen aus. Scalar
Object-Modulfelder und begrenzte qualifizierte DAO-Felder/Parameter/Rueckgaben
halten echte Referenzen/Nothing und pruefen Typen. Explizite Static-Deklarationen
in Prozeduren behalten Werte/Objekte/Arrays je Host, Prozedur und Quelltext.
Die groesseren Static-Sub/Function-Regeln bleiben ausdruecklich Unsupported.
Details und Grenzen: SCRIPTING_DICTIONARY_RUNTIME.md, DAO_OBJECT_BINDINGS.md,
OPTIONAL_ARGUMENT_RUNTIME.md, STATIC_LOCAL_RUNTIME.md (alle unter docs/).
Dieser Stand: 48/48 Desktop, frisch GUI-frei 1/1, Debian-Paket und vier Original-
Fixture-Hashes bestanden. 372 erledigte / 72 offene erfasste Checklistenpunkte;
keine vollstaendige Access-Paritaet. Originalabfragen-/Login-Grenzen und fehlende
Access-2003-Ergebnisreferenz bleiben bestehen.


## Native Arrays und weitere Datenblatt-Tasten (2026-11-04)

Gepackte Byte-Arrays tragen jetzt in VbaValue.object eine opake interne
Allokationskennung; Kind bleibt ByteArray und IsObject bleibt False.
withVbaArrayElement erhält die Kennung bei Wertkopien/Elementänderungen, ReDim
erzeugt eine neue. For Each nutzt dieselben Quell-/Allokationsprüfungen wie
andere native Arrays und liefert echte Byte-Varianten. Nicht als COM-Objekt
behandeln oder diese Kennung als vermeintlich ungenutztes Feld entfernen.

Erase-Listen und ReDim-/Preserve-Listen bleiben je eine Anweisung/Quellzeile.
ReDim hat quelltextreine Ziel-/Dimensions-/Typdeskriptoren im begrenzten
Übersetzungscache. Erfolgreiche Ziele müssen den Modul-/ByRef-Zustand vor
späteren Dimensionsausdrücken veröffentlichen; der Alias-Abgleich zwischen den
Zielen ist notwendig. Konstantenziele werden vor dem Prozedurrumpf abgewiesen.

Datenblätter unterstützen echte ganze Textfeld-Ersetzung über Strg+H/Menü/
acCmdReplace und aktuelles Datensatz-Verwerfen über Esc/Kontextmenü. Pfeil
hoch/runter, benachbarte Tab-Zeilenwechsel samt Editorabschluss sowie
Strg+Pos1/Strg+Ende nutzen begrenzte native Prüf-/Save-/Navigationspfade.
Vorhandene Entwürfe bleiben bei SQL-Fehlern erhalten. Keine komplette Undo-,
Tastatur-, Ereignis-, Query-Updateability- oder Access-2003-Parität behaupten.
Details: docs/VARIANT_ARRAY_RUNTIME.md, docs/OBJECT_DEFAULT_ASSIGNMENT_RUNTIME.md,
docs/TABLE_REPLACE_RUNTIME.md, docs/TABLE_RECORD_UNDO_RUNTIME.md und
docs/TABLE_KEYBOARD_RUNTIME.md.

Dieser Stand: 48/48 Desktop, GUI-freier Build/Kern 1/1, Debian-Paket und vier
Originalfixture-Hashes bestanden. 382 erledigte / 72 offene erfasste Punkte;
keine globale Prozent- oder Vollständigkeitsbehauptung. Originalabfragen/Login
und Access-Ergebnisreferenz-Grenzen bleiben bestehen.


Fehlerfortsetzung: Nach jedem expliziten Resume/Resume Next/Resume label müssen
die Fehler-Snapshots freigegeben werden. Wiederhergestellte aktive Loop-/With-
Frames halten ihre Ressourcen selbst weiter; veraltete Snapshots dürfen nach
End With/Next keine DAO-Clones oder Collection-Leases festhalten. Auch einzelne
Next-Anweisungen tragen sourceStatementEnd für das Aufräumen beim Überspringen.
Gezielte bestehende Kern-/DAO-Tests bestanden (2/2, 1,67 s). Der DAO-Fall setzt
den Clone explizit mit MoveFirst; die Transaktionssperre bei lebendem Clone
bleibt bestehen. Keine neue CTest-Suite und keine zusätzliche Gesamtpunktzahl.

Abschluss dieser Korrektur: Desktop 48/48 (81,49 s), frisch gebauter GUI-freier
Kern 1/1 (1,62 s), Debian-Paket und vier Originalfixture-Hashes bestanden.
Die wachsende Zahl feiner Checklisten-Unterpunkte nicht mehr als Gesamtstatus
verwenden. Historische Zählungen sind keine Access-2003-Gesamtmetrik.

Verbindliche Statusstruktur ab 2026-11-04: docs/FIXED_WORK_CATALOG.md enthält
alle 456 bisherigen Punkte unter unveränderlichen IDs F001–F456, fachlichen
Kategorien und Unterkategorien. Gesamtzahl nicht mehr erhöhen. Neue Details
und Befunde bestehenden IDs zuordnen; keine neuen Checkbox-Punkte anhängen.
384 bestätigte Teilumfänge / 72 offene Voll-/Teilanforderungen sind keine
84-Prozent-Access-Parität. Breite offene Anforderungen enthalten die weitere
Entwicklung einschließlich neu entdeckter Access-Details.

In Arbeit unter F051: Err.Raise mit fünf positionalen/benannten Argumenten,
Missing-Optionalen, signierten Long-Codes/vbObjectError und beschreibbaren
Err-Feldern. VbaRuntimeError/VbaExecutionResult tragen Source/HelpFile/HelpContext
durch native Unteraufrufe und Ereignisfunktionen. Clear/Resume setzt die
Metadaten zurück. Bestehender Kerntest gezielt bestanden (1/1, 1,25 s);
Gesamtprüfung dieser Erweiterung folgt. Kein Windows-Hilfesystem, kein
vollständiger Access-Fehlertextkatalog/HRESULT-/Debugger-Nachweis.

Abschluss Err-Metadaten/fester Katalog: Desktop 48/48 (77,97 s), frisch gebauter
GUI-freier Kern 1/1 (1,39 s), Debian-Paket und vier Originalfixture-Hashes
bestanden. Gezielt Kern 1/1 (1,25 s). F051 bleibt als breite Vollanforderung
offen; 456 feste IDs, unverändert 384 bestätigte Teilumfänge / 72 offen.

Aktueller Codeabschnitt unter F041/F042/F051 (noch Gesamtprüfung offen):
Error expression ist eine eigene native, quellzeilenbezogene Anweisung und
hängt nicht von einer lokalen Variablen namens Err ab. Unary +/- propagiert
Null nach tatsächlicher Skalarauflösung. Enum-Mitglieder sind echte immutable
Long-Konstanten mit automatischer Nummerierung, expliziten ganzzahligen
Konstanten, Aliasen, negativen Werten und Überlaufprüfung. Dieselben geordneten
Deklarationen dienen Übersetzung und Modulzustand; unqualifizierte und
Enum.Member-Namen samt Public/Private sind unterstützt. Lokale/Instanz-Wurzeln
behalten Vorrang vor Enum-Namensräumen. Kein Laufzeit-Enum-Objekt erfinden.
As-Enum-Variablen/Arrays/Parameter, Typbibliotheken und vollständige
Klassen-/Modulsichtbarkeit bleiben offen. Details: docs/DAO_RUNTIME.md.
Keine neuen Katalog-IDs/Checkbox-Punkte oder CTest-Ziele.

Gezielte Prüfung des kombinierten Error-/Null-/Enum-Abschnitts: Kern/DAO 2/2
(1,64 s), frisch gebauter GUI-freier Kern 1/1 (1,31 s). Qualifizierte Enum-
Konstanten arbeiten auch in Aliaswerten, Modul-Arraygrenzen und Optional-
Defaults; ein echter lokaler Collection-Wurzelname hat weiter Vorrang.
Gesamtsuite/Paket dieser Kombination laufen als nächstes.

Abschluss Error-/Null-/Enum-Konstanten: Desktop 48/48 (82,80 s), frisch gebauter
GUI-freier Kern 1/1 (1,31 s), gezielt Kern/DAO 2/2 (1,64 s), Debian-Paket und
vier Originalfixture-Hashes bestanden. Gesamtzahl unverändert 456; breite
F041/F042/F051 bleiben offen.

Aktueller GoSub-/Return-Abschnitt unter F040/F051: Eigene native
Anweisungen/Labelbindung und ein auf 48 Einträge begrenzter prozedurlokaler
Rücksprungstapel. Aufrufer-If/For/For-Each/With-Scope bleibt lebendig; Return
räumt nur die Scope-Ebenen des Unterprogramms auf. GoTo innerhalb GoSub
behält den jeweiligen Aufrufer-Floor. Return ohne GoSub Fehler 3. Label-Handler
sichern den Rücksprungstapel mit den Fehler-Scopes; eigene Handler-GoSubs und
Resume/Resume 0/Resume Next/Resume label sind angebunden. Veraltete Snapshots
freigeben. Direktes Return vom Handler in einen suspendierten Fehler-GoSub
bleibt mit 3251 gesperrt, bis dessen Referenzsemantik feststeht. Numerische
Labels und Labels innerhalb von Blöcken bleiben offen.
Gezielt Kern/DAO 2/2 (1,80 s) bestanden. Gesamtsuite/GUI-freier Build/Paket
dieses Abschnitts werden als nächstes geprüft. Keine neuen Katalog-IDs
oder CTest-Ziele; 456 bleiben fest. Details: docs/DAO_RUNTIME.md.

Abschluss GoSub/Return: Desktop 48/48 (80,12 s), frisch gebauter GUI-freier
Kern 1/1 (1,29 s), gezielt Kern/DAO 2/2 (1,80 s), Debian-Paket und vier
Originalfixture-Hashes bestanden. Feste 456 Punkte/IDs unverändert.

Aktueller F040-Abschnitt (Gesamtprüfung folgt): GoTo kann Labels im selben
aktiven Block oder einem Vorfahren anspringen. Quelltextreine Unwind-Zähler
räumen nur verlassene If/Select-/For-/With-Ebenen auf; GoSub-Floors bleiben
erhalten. Sprünge in nicht aktive/Geschwisterblöcke werden vor dem Rumpf
abgewiesen. Labelnamen aus Inline-If-Wrappern müssen samt Doppelpunkt erhalten
und nach dem Zusammenführen global gebunden werden. Die temporäre Scope-Tabelle
braucht einen Integer pro Anweisung, keine quadratischen Kopien der Pfade.
Numerische Labels 0–65529 mit/ohne Doppelpunkt und führenden Nullen sind
angebunden. On Error GoTo 0 deaktiviert weiterhin, Resume 0 wiederholt.
Handler-/GoSub-/Resume-Labelziele bleiben top-level; Erl/Debugger bleibt offen.
Gezielt Kern/DAO 2/2 (1,98 s) bestanden; danach vorhandenen DAO-Clone-Fall um
GoTo aus With ergänzt, dessen Ergebnis ist noch zu prüfen. Keine neuen
Katalog-/CTest-Einträge; feste 456 bleiben. Details: docs/DAO_RUNTIME.md.

Abschluss Scope-GoTo/numerische Labels: Desktop 48/48 (76,53 s), frisch
gebauter GUI-freier Kern 1/1 (1,37 s), gezielt Kern/DAO 2/2 (1,98 s) und
DAO-Clone-Abzweig 1/1 (0,37 s), Debian-Paket und vier Originalfixture-Hashes
bestanden. Native With-Clone-Freigabe durch GoTo ist damit geprüft.
Feste 456 IDs/Punkte und breite offene Vollanforderungen unverändert.

Aktueller ByRef-Zellenabschnitt unter F043 (noch uncommittet/Gesamtprüfung
offen): Argument-Metadaten transportieren benannte Arraywurzel, einmalig
geprüften linearen Offset, Elementtyp und opaken Allokationstoken. Token darf
keine Arrayinhalte/DAO-Objekte halten. Sub-/Funktions-/Named-Aufrufe und Scalar-
Weiterleitung nutzen echte gemeinsame getter/setter-Slots. ByRef liest nach
späteren Argumenten aktuellen Zellinhalt, ByVal den früheren Wert. Ganze
Quellersetzung/ReDim/Erase bei aktiver Zelle ist gesperrt; Compound-/ParamArray-
Adressen und Arraywurzeln durch bestehende Zellaliase bleiben 3251.
Argumentkopien echter ByRef-Aliase nach Bindung freigeben, Named-Reihenfolge
verschieben statt alte Kopien zu halten. Nur Alias-Publikation schreibt Quellen;
ByVal-Werte niemals in ursprüngliche Variablen zurückkopieren. Aktive Alias-
Frames anderer Aufrufer nach Publikation aktualisieren, eigenen Frame erst nach
seinen Schreiboperationen: sonst halten alte Modul-/Static-Snapshots DAO-Clones.
Kern/DAO gezielt 2/2 (1,88 s) vor zusätzlicher Compound-Adress-Regression
bestanden; diese sowie Gesamtsuite/GUI-freier Build/Paket noch zu prüfen.
Keine zusätzlichen Katalog-/CTest-Einträge. Details: VARIANT_ARRAY_RUNTIME.md.

Aktualisierung ByRef-Zellen: Auch interne ByVal-Transportkopien werden nach
Bindung freigegeben; die eigentlichen lokalen ByVal-Werte bleiben erhalten.
Frische unbenutzte Zellslots nicht als Objektbesitzer behalten. Öffentliche
C++-Argumentvektoren (ohne interne Referenzbindungen) behalten ihren Vertrag.
Der echte DAO-Fall leitet einen Modul-Array-Clone weiter, löst eine zusätzliche
ByVal-Kopie sowie die ByRef-Quelle auf und startet die Transaktion bereits vor
der Rückkehr. Aktualisierte gezielte Kern-/DAO-Prüfung 2/2 (2,21 s) und frisch
gebauter GUI-freier Kern 1/1 (2,05 s) bestanden. Gesamtsuite/Paket folgen.

Abschluss ByRef-Zellen/Alias-Lebensdauer: Desktop 48/48 (76,34 s), frisch
gebauter GUI-freier Kern 1/1 (2,05 s), gezielt Kern/DAO 2/2 (2,21 s),
Debian-Paket und vier Originalfixture-Hashes bestanden. F043 bleibt breit
offen; feste 456 IDs/Punkte unverändert.


Aktueller Formularabschnitt unter F080: Eigene _la_objects-Formulare laufen
jetzt über AccessFormPreview statt alten modalen QDialog/manuelle SQL-Schritte.
Adapter native_stored_form.hpp erzeugt echte Layout-/Ereignisbindung, ergänzt
Default-Neu/Speichern nur zur Laufzeit, erhält gespeichertes VBA/Originalbytes.
DoCmd.OpenForm löst eigene Formulare auf, native/importierte Namenskollisionen
bleiben Fehler. Feldreihenfolge/Beschriftungen im bestehenden Editor gespeichert;
Entwurf schließt geladenes Zielformular unter Beachtung von Unload-Abbruch.
Bestehender GUI-Target 1/1 (2,24 s): echter Designer, SQL-Save, VBA-Instanz,
Kind öffnen/Aufrufer schließen, IsLoaded und erneuter Entwurf bestanden.
Gesamtsuite/GUI-freier Build/Paket/Hashes noch zu prüfen. Keine neuen IDs/Targets;
456 fest, breite F080-Vollanforderung bleibt offen. Native eigene Unterformulare
und freier visueller Entwurf sind damit nicht implementiert.

Abschluss eigene Formular-Dokumentlaufzeit: Desktop 48/48 (82,61 s),
GUI-freier Build und Kern 1/1 (1,52 s), Debian-Paket und alle vier
Originalfixture-Hashes bestanden. Gezielt GUI 1/1 (2,24 s). Feste 456,
384 bestätigte Teilumfänge und 72 offene Vollanforderungen unverändert.


Folgeabschnitt F080/F058: Eigene Formulare jetzt auch als echte ungebundene
Unterformulare (ohne Master/Child-Verknüpfung). SQL/VBA, Dirty=False-Speichern,
IsLoaded und Eltern-Schließen nutzen vorhandene eingebettete Dokumentlaufzeit.
Native Katalogabfragen qualifizieren main, damit TEMP-Schatten keine eigenen
Definitionen verdecken. Gezielt GUI/Layout 2/2 (11,47 s) bestanden; keine neuen
Targets/IDs. Letzte Vollsuite 48/48 (82,61 s) war vor diesem Folgeabschnitt.
Master/Child-Verknüpfung bleibt ausdrücklich 3251, kein vollständiger Designer.


F080 erweitert: Eigene Formularfelder speichern left/top/width/height in Twips
als field_layout. Adapter validiert ganzzahlige Grenzen, erweitert Canvas und
Buttonpositionen entsprechend echten Feldgrößen. Eigenschaftenanzeige folgt
Auswahl, Häkchenreihenfolge und Drag/Drop; Standardpositionen zählen nur gewählte
Felder. GUI/Layout 2/2 (11,23 s) vor finaler Anzeige-Aktualisierung bestanden.
Vollprüfung dieses Folgeabschnitts steht noch aus; feste 456 unverändert.
Keine Behauptung eines vollständigen visuellen Access-Designers.

Abschluss eigene Unterformulare/Feldgeometrie: Desktop 48/48 (78,56 s),
Debian-Paket und vier Originalfixture-Hashes bestanden. GUI-freier Build/Kern
zuletzt vor diesen GUI-Änderungen 1/1 (1,52 s). Feste 456 bleiben unverändert.


Aktueller Mausabschnitt F058/F148: Native TextBox statt nackter QLineEdit führt
Click nach Editor-Release aus; Cursoränderungen durch VBA bleiben wirksam,
kein implizites Commit durch Klick. OnMouseDown/OnMouseUp im Layout erhalten,
Integer Button/Shift und Single X/Y in nativen 96-dpi-Twips übergeben. Fremde
Zusatzbuttons 3251, Ereignisfunktionen über vorhandene Laufzeit. Fehler stoppen
restlichen Click, Close/Disposal während Handler geschützt. Bestehender GUI-Fall
prüft Reihenfolge, Koordinaten/Modifier, SelStart, unveränderten Dirty-Zustand,
Rechtsklick und DblClick-Cancel. Gezielt GUI/Layout 2/2 (11,43 s) vor finalem
Zusatzbutton-Guard bestanden. Gesamtprüfung folgt. Details KEYBOARD_RUNTIME.md;
MouseMove/andere Steuerelemente/Form-/vollständige Fokus-/Skalen-/Windows-Orakel
bleiben offen. Keine neuen Targets/IDs; 456 fest.

Abschluss TextBox-Mausereignisse: Desktop 48/48 (79,27 s), frisch gebauter
GUI-freier Kern 1/1 (1,56 s), Debian-Paket und vier Originalfixture-Hashes
bestanden. Gezielt GUI/Layout 2/2 (11,43 s). Feste 456 und 48 Targets bleiben.


Folgeabschnitt F058/F148: Native Command-/Toggle-Buttons und eigenständige
CheckBox/OptionButton nutzen gemeinsame typisierte MouseDown/MouseUp-Bindung.
MouseUp läuft vor tatsächlicher Button-Aktivierung; Fehler lösen Druckzustand
und stoppen Click. Doppelclick erzeugt keinen erfundenen zweiten MouseDown.
OptionGroup-Mitglieder bleiben für Mausbindung ausgeschlossen. Eigene native
Aktionsbuttons binden gespeicherte MouseDown/Up-Prozeduren. GUI-Fall beweist:
MouseUp schreibt Feldwert, nachfolgender echter Save persistiert genau diesen.
Vorhandener AllowEdits-Fall: SQL down/up/click-Spur trotz schreibgeschützter
Checkbox, Datenwert bleibt unverändert. Gezielt GUI/Layout 2/2 (11,23 s).
Letzte Vollsuite 48/48 (79,27 s) und frischer GUI-freier Kern 1/1 (1,56 s)
waren vor diesem GUI-Folgeabschnitt. Paketprüfung folgt. Keine neuen IDs/Targets;
456 fest. Gruppen-/Move-/vollständige Fokus-/Skalen-/Windows-Orakel weiter offen.

Button-Folgeabschnitt: Debian-Paket und vier Originalfixture-Hashes bestanden.
Gezielt GUI/Layout 2/2 (11,23 s), Vollsuite zuletzt vor diesem GUI-Abschnitt
48/48 (79,27 s). Keine Ausweitung der 48 Targets oder 456 Katalogpunkte.


Aktueller VBA-Abschnitt F040/F045: On expression GoTo/GoSub mit einmaliger
Selektorauswertung, Let-Integer-Koerzierung vor Bereich 0–255 (5), Index 0/über
Listengröße fällt durch. Alle Ziele vor Ausführung binden, je GoTo-Ziel echte
lexikalische Unwindzahlen; GoSub nutzt vorhandene Return-/Resume-/Scope-Frames.
Inline-Ziele erst nach endgültigem Wrapper-Merge binden. Quellplan-Cachebudget
zählt Zielvektoren/Labelbytes. Bestehende Sprünge-in-inaktive-Blöcke-/GoSub-/
Handler-/Resume-Zielrestriktionen bleiben. Kern/DAO 2/2 (1,94 s) bestanden:
gezählte/fraktionale Selektoren, verschachtelte inline-GoSubs, Resume nach
Selektorfehler, 5/6/94, numerische Labels und Collection-Lease-Freigabe.
Anschließend vorhandenen DAO-Clone-Ausgang zu On 1 GoTo umgestellt; dessen
Ergebnis und Vollsuite/GUI-freier Build/Paket noch zu prüfen. 456/48 unverändert.
Details DAO_RUNTIME.md, öffentliche MS-VBAL-/On-Verzweigungsreferenzen.

Abschluss berechnete VBA-Verzweigungen: Desktop 48/48 (81,46 s), frisch
gebauter GUI-freier Kern 1/1 (1,65 s), Debian-Paket und vier Originalfixture-
Hashes bestanden. Vorhandener DAO-With/Clone-Fall verlässt Scope via On 1 GoTo
und startet danach echte Transaktion erfolgreich. Die Vollsuite deckt auch den
vorherigen Button-Maus-Folgeabschnitt ab. Feste 456 und 48 Targets unverändert.


Aktueller TypeOf-Abschnitt F042/F043/F122: VbaObject::isNativeType als explizite
native Klassen-/Interface-Identität; nicht aus TypeName raten und keine skalaren
Standardwerte lesen. Collection/Dictionary/Err/DAO sowie native Access Form,
Report, Controls/Control-Typen, Application/Forms/Reports markiert. Control-/
Dokument-Metadaten unabhängig von Feldwert/weak Lebensdauer gespeichert.
TypeOf-Operand bleibt Referenz, Nichtobjekt-Variant False, Nothing Fehler 91
(MS-VBAL 5.6.7). Unbekannte Klassen/Referenzen, Value/Literal-Operandformen und
Calculated-Ausdrücke explizit 3251; volle statische Deklarationsprüfung, UDTs,
Klassenhierarchien/Referenzordnung/COM/Windows-Orakel bleiben offen.
Gezielt Kern/DAO/GUI/Layout 4/4 (13,14 s) bestanden. Anschließend echten
DAO-Clone-TypeOf mit sofortiger Transaktion ergänzt (Temp-Clone-Freigabe);
diese Ergänzung/Gesamtsuite/GUI-freier Build/Paket noch prüfen. 456/48 unverändert.
Details VARIANT_ARRAY_RUNTIME.md, öffentliche MS-VBAL-Referenz. Public PDF
nur außerhalb Repositorys unter linaccess-onboarding/ms-vbal-public.* gelesen.

TypeOf-Abschluss: Gesamtlauf 47/48 (79,35 s). Neuer DAO-Probe-Clone war bereits
freigegeben, BeginTrans erfolgreich; nach Rollback wurde allerdings der nach
bestehendem Vertrag invalidierte Cursor weiterverwendet (3420). Probe korrigiert
auf Commit der reinen Lesetransaktion, DAO gezielt 1/1 (0,44 s) bestanden.
Laufzeitcode nach Gesamtlauf unverändert: alle 48 vorhandenen Targets damit über
diese Läufe geprüft, übrige 47 nicht wiederholt. Frischer GUI-freier Build/Kern
1/1 (1,69 s), Debian-Paket und vier Originalfixture-Hashes bestanden. Letzter
sauberer einzelner Vollsuite-Lauf 48/48 war vorheriger Verzweigungscommit
(81,46 s). TypeOf jetzt geprüft inkl. Temp-Clone-Freigabe ohne Guards zu lockern.
Feste 456/48 unverändert; vollständige statische Typbindung/Klassen/COM offen.


Aktueller Access-Typabschnitt F050/F043/F122: native_object_types.hpp als
C++-Vokabular für TypeOf und Typbindung. Bekannte Access-Klassennamen in lokalen/
Modulvariablen, Parametern/Rückgaben kanonisiert; Default Nothing, konkrete
Form/Report/Control-/Controlklassen/Collections/Application per isNativeType
geprüft. Falsche Klassen Fehler 13 vor Slotersetzung/Callee-Körper. Optionale
Access-Objekte =Nothing, ByRef-Aliase Form/Access.Form identisch, VBA.Collection
auf vorhandene Collection normalisiert. Parameter-Signaturfilter erkennt neue
Typen, unbekannte Klassen bleiben vor Körperausführung Unsupported.
Gezielt Kern/DAO/GUI/Layout 4/4 (13,84 s): vollständiges Form_Load explizit
geprüft, typisierte Rückgabe/Modulfelder/Optional/ByRef, dynamische Fehltypen,
realer TextBox-Let-Edit mit SQL unverändert und Undo, echte Report-Referenzen.
Gesamtprüfung/GUI-freier Build/Paket folgen. Direkte kontextsensitive Property-
Actuals, Klassenarrays, vollständige statische Member-/Projekt-/Referenz-/COM-
Bindung und nicht-identische ByRef-Interfacekonvertierungen bleiben offen.
Feste 456/48 unverändert. Details VARIANT_ARRAY_RUNTIME.md.

Abschluss native Access-Typslots: Desktop sauberer Vollsuite-Lauf 48/48
(85,72 s), frischer GUI-freier Build/Kern 1/1 (2,15 s), Debian-Paket und vier
Originalfixture-Hashes bestanden. Deckt TypeOf/Temp-Clone-Korrektur mit ab.
Feste 456, 384 bestätigte Teilumfänge/72 offen und 48 Targets unverändert.


Aktueller Objektargument-Abschnitt F043/F122: VbaFunctionResolver ist weiterhin
mit bisherigen Callbacks konstruierbarer Aufrufer plus optionale Signatur-
Metadaten. VbaArgumentSignature enthält Object-Kontextflags je Parameter.
Gleiches-Modul-Subs/Funktionen und positionale Ereignisfunktionen erhalten echte
Me-/Control-/DAO-Referenzen für deklarierte Objektparameter. Named-Actuals in
Sub/Funktionscode wählen Kontext nach Parametername, skalare/Array-/Intrinsic-
Pfade bleiben erhalten. Native Objektmethoden können Flags ebenso angeben.
Pro-Frame-Signaturcache hält nur Namen/Flags, keine Objekt-/Widget-/DAO-Werte.
Gezielt Kern/DAO/GUI/Layout 4/4 (13,58 s): echte DAO-Bang-Field-Parameter,
direkte Me!first-Control-Actuals, umgekehrt benannte Funktionsparameter sowie
function-bound Click mit echtem Qt-Tag/SQL-Spur. Vollprüfung folgt.
Externe/qualifizierte Modulbindung, generelle Property-ByRef-Setter, komplette
Klammer-/Varianten-/statische Kontextregeln und Named-Ereignisargumentordnung
bleiben offen. Feste 456/48 unverändert. Details VARIANT_ARRAY_RUNTIME.md.

Abschluss Objektargument-Kontext: Desktop 48/48 (85,57 s), frischer GUI-freier
Build/Kern 1/1 (1,99 s), Debian-Paket und vier Originalfixture-Hashes bestanden.
Vorhandene ByRef-Arrayzellen/DAO-Clone-/Transaktions-Lebensdauerprüfungen bleiben
grün. Feste 456 Katalogpunkte und 48 CTest-Targets unverändert.


Aktueller Funktionsbindungsabschnitt F042/F118: Parenthesierte gleiches-Modul-
Prozeduren vor Bibliotheksfunktionen, VBA.qualifiziert weiterhin Bibliothek.
Signaturmetadaten enthalten Function-/Ergebnistypklassifikation; Sub als Wert
vor Körper explizit 3251. Eigene Optional-/Named-/Missing-Regeln vor Builtin-
Argumentpfaden, Objektreferenz-Kontext folgt gewählter eigener Funktion.
Modulzustand hält reine Prozedurnamen, somit keine ganzen Modulsuchen je
Intrinsic-Aufruf; eigene Zustandsprovider ohne Inventar nutzen Frame-Fallback.
Deklarierte ununterstützte Routinen fallen nicht still auf Bibliothek zurück.
MS-VBAL-Left-Ausnahme für parameterlose skalare Funktion bei 2 Actuals vorhanden.
Gezielt Kern/DAO/GUI/Layout 4/4 (14,05 s): eigene Len/Round, VBA.Len, Optional/
Named/Missing, gezählte Körper, Left-Fallback, Sub-als-Wert-Sperre, echte
Control-Argumente an eigene Len, function-bound Click mit Qt-/SQL-Effekten.
Vollprüfung folgt. Bare names, skalare/indexierte Variablen-/Property-/Default-
Member-Bindung und externe/Projekt-/Referenz-/vollständige Klassenscopes offen.
456/48 unverändert. Details VARIANT_ARRAY_RUNTIME.md und öffentliche MS-VBAL.

Abschluss lokale Funktionspriorität: Desktop 48/48 (88,44 s), frischer
GUI-freier Build/Kern 1/1 (2,34 s), Debian-Paket und vier Originalfixture-Hashes
bestanden. Kurzzeitiger Exec-Transportausfall: Workspace/laufende Prüfungen
blieben erhalten, danach Prozesse/Status geprüft, keine Sourceänderung während
Prüfungen. Feste 456/48 unverändert. Nächster Fokus: echte Berichtsdarstellung.

## Aktuelle Fortsetzung: native Domänen, Berichte, Abfragen und Dateiauswahl (2026-11-04)

Der feste Katalog bleibt456 IDs, 14 Unterkategorien, 384 bestätigte Teilumfänge und
72 breite offene Punkte. Kein gewichteter oder nachgewiesener Paritätsprozentsatz.
Neue Details innerhalb vorhandener IDs bearbeiten; keine neuen IDs/Testtargets.

Seit dem früheren Berichtspfad sind folgende begrenzte native Ausführungen
implementiert und privat gesichert beziehungsweise Bestandteil dieses Stands:
- Feldbasierte DLookup/DMin/DMax/DAvg/DFirst/DLast/DVar/DStDev sowie gemeinsame
  globale/Application DSum/DCount mit bewiesenen SQL-Rückgabetypen:
  docs/NATIVE_DLOOKUP.md. Keine vollständige Expr-/Jet-/Oracle-Behauptung.
- Bericht First/Last, typisierte Date/Currency-Min/Max, exakte Currency-Summen und
  Stichproben Var/StDev: docs/REPORT_AGGREGATES.md. Populations-/Gruppen-/vollständige
  Typ-/Rundungsparität bleibt offen.
- Eigener gespeicherter berechneter Berichtsfuß mit Abschnittsseitenwechsel:
  docs/NATIVE_REPORT_FOOTER_DESIGN.md; fester einzelner Fuß, kein Voll-Designer.
- Schreibgeschützte eigene Abfrage-/SELECT-Formulare, Auflösung eigener gespeicherter
  SQL-Definitionen, eigene SELECT-Abhängigkeitsansichten und Revision vor Requery:
  docs/NATIVE_QUERY_FORM_DESIGN.md. Parameter/allgemeine Abfrage-Schreibbarkeit offen.
- Native DAO Byte()-Feldwerte mit Edit/Update/Cancel und begrenztem gespeichertem
  BLOB-FieldSize: docs/DAO_BINARY_FIELDS.md. OLE-Extraktion/-Darstellung, Chunks,
  binäre QueryDef-Parameter und Originalformat-Roundtrip bleiben offen.
- Application.FileDialog FilePicker/FolderPicker mit echten Qt-Dialogen:
  docs/NATIVE_FILE_DIALOG.md. Explizit22 Access- und4 Office-Klassen im geschlossenen
  nativen Typvokabular; kein COM-/TypeName-Raten. Der Office-Adapter ist kontextgebunden,
  keine vollständige Office-Lebensdauer-/Windows-/CWD-Kompatibilitätsbehauptung.

Letzter vollständiger Lauf dieses Abschnitts: Desktop48/48 (91.42 s), frischer
GUI-freier Core1/1 (1.50 s), Debian Nutzlast/Build-ID/SQLite und vier originale
Fixture-SHA-256-Prüfungen bestanden. Keine vollständige Access2003-Parität.
Programmierfortschritt priorisieren, nach Risiko bestehende gezielte Prüfungen
nutzen und erst bei neuen gemeinsamen Laufzeit-/Persistenzrisiken verbreitern.

Eigene benannte DAO QueryDefs: tatsächlicher Name/Ursprung, transaktionstreue
bedingte SQL-Metadatenänderung mit Erhalt fremder JSON-Felder, frische SQL-Sicht
anderer Instanzen und ungültige alte Parameterreferenzen. Importierte Instanz-SQL
und temporäre QueryDefs bleiben getrennt; Originalquellen werden nicht ersetzt.
DAO-SQL-Recordsets aktualisieren Abhängigkeitsansichten vor dem Laden; Workspace-
Lebensdauer gilt auch für QueryDefs/Parameter. Grenzen: docs/OWNED_QUERYDEF_SQL.md.
Vollstand danach48/48 (82.63 s), Paket und vier Originalfixture-Hashes bestanden;
GUI-freier Laufzeitscope unverändert, letzter frischer Core1/1 (1.50 s).

DAO QueryDefs-Sammlung, benanntes/leeres CreateQueryDef und eigenes Delete sind
jetzt im begrenzten nativen Katalog vorhanden: docs/DAO_QUERYDEFS_CATALOG.md.
Count/Item/Refresh/For Each lesen tatsächliche Metadaten. Native Create/Delete-
Generationen verhindern Umlenkung alter Objekte auf gleichnamige Neuerstellungen.
Workspace-Rollback stellt auch die fünf Abfrageansichten-Registry-Eigenschaften
passend zum zurückgerollten SQLite-Temp-Schema wieder her. Importierte Löschung
und Formatänderung bleiben offen, Archivquellen bleiben erhalten. Externes manuelles
Drop/Recreate, Rename/Append/Attributes und Windows-Reihenfolge/Oracle sind nicht
vollständig. DAO.QueryDefs ist ein weiterer expliziter geschlossener nativer Typ.
Vollstand danach48/48 (83.20 s), frischer GUI-freier Core1/1 (1.66 s), Paket und
vier Originalfixture-Hashes bestanden; Katalog unverändert456/384/72.

Eigene Abfrage-/Formular-/Berichtsentwürfe speichern per Strg+S tatsächlich und
bleiben geöffnet; ausdrückliche Speichern-Schaltflächen schließen weiterhin nach
Erfolg. Wiederholtes Speichern erhält andere aktuelle JSON-Metadaten; Berichte
lesen nach der ersten Tastaturspeicherung auch frisch gespeicherte Seiteneinstellungen.
Menüfolge Datei/Bearbeiten/Ansicht/Einfügen/Fenster mit tatsächlichen Objektaktionen;
keine Attrappen für fehlende Access-Menüs. Dateikürzel sind window-scoped, bestehende
modale/KeyPreview-Projektwechselregeln bleiben. docs/DESIGNER_SAVE_KEYS.md.
Aktueller UI-Lauf1/1 (2.63 s), Paket und vier Originalfixture-Hashes bestanden;
letzter breiter Stand48/48 (83.20 s), GUI-freier Runtime-Scope unverändert1/1 (1.66 s).

Partition(Number,Start,Stop,Interval) führt nun dieselbe native VBA-Berechnung auch
über den SQL-Adapter aus, einschließlich benannter VBA-Argumente, Null,
Nearest-even-Long-Konvertierung, aufgefüllter Grenzen und tatsächlicher
GROUP-BY-Buckets. Stop=Long-Max bleibt explizit3251 ohne Windows-Oracle.
docs/NATIVE_PARTITION.md; keine neue CTest-Zielzahl und keine neuen Katalog-IDs.
Vollstand48/48 (84.06 s), frischer GUI-freier Core1/1, Paket/Build-ID/SQLite und
vier Originalfixture-Hashes bestanden. Weiterhin456/384/72, keine Vollparität.

FileDialog.Filters unterstützt nun natives For Each über tatsächliche typisierte
Filterobjekte in der aktuellen Einfügereihenfolge einschließlich Add Position.
Leere/Clear-Sammlung liefert keine Elemente. Bestehende echte VBA-/Qt-Dateiauswahl
prüft Reihenfolge, Typen und leere Enumeration, danach tatsächliche Textdateiauswahl.
Layout-Ziel1/1 (10.56 s), Paket/Build-ID/SQLite und vier Originalfixture-Hashes
bestanden. Letzter Vollstand48/48 (84.06 s), Core1/1; weiterhin456/384/72.

Native MkDir-/RmDir-Anweisungen kompilieren und führen tatsächliche
Linux-Verzeichnisoperationen aus. RmDir nutzt den Betriebssystemaufruf nur für
leere Verzeichnisse und entfernt keine normale Datei. Unicode-/Fehler-/Inhalterhalt-
Regression im bestehenden Core-Ziel. Windows-Pfade/Symlink-Entfernung3251;
vollständige Datei-I/O und Windows-Oracle bleiben offen.
docs/NATIVE_DIRECTORY_STATEMENTS.md; F040 bleibt offen, keine neue ID.
Aktuell48/48 (82.86 s), frischer GUI-freier Core1/1 (1.76 s), Paket/Build-ID/SQLite
und vier Originalfixture-Hashes bestanden; fester Katalog weiterhin456/384/72.

FV/Pmt führen im begrenzten nativen VBA-/SQL-Pfad echte Finanzberechnungen aus;
benannte Argumente nutzen VBA-Type, optionale Werte0, Double-Ergebnis,
log1p/expm1 gegen Nullzins-Auslöschung und echte SQL-Zeilen. Ganze Perioden0..32767,
Rate>-1, Type0/1; andere Randbereiche3251, Pmt mit0Perioden5.
docs/NATIVE_FINANCIAL_FUNCTIONS.md; weitere Finanzfunktionen/Access-Oracle offen.
Gezielt3/3 (2.38 s), frischer GUI-freier Core1/1 (1.54 s), Paket/Build-ID/SQLite und
vier Originalfixture-Hashes bestanden. Letzter Vollstand davor48/48 (82.86 s),
kein neuer Vollsuite-Nachweis behauptet; Katalog unverändert456/384/72, F068 offen.

Extras/Optionen verändert tatsächlichen klassischen Stil und Anwendungsschrift
(Familie/Punktgröße) über Übernehmen/OK/Abbrechen und persistente QSettings.
Startup lädt gespeicherte Schrift auch ohne klassischen Stil; explizite
Quell-Control-Schriften bleiben eigenständig. Menüfolge nun Datei/Bearbeiten/
Ansicht/Einfügen/Extras/Fenster/?; ? öffnet tatsächliche LinAccess-Information.
docs/NATIVE_APPLICATION_OPTIONS.md. Kein vollständiger Access-Options-/Pixel-Claim.
Finale Mainwindow1/1 (2.71 s), Layout1/1 (10.70 s), Paket/Build-ID/SQLite und vier
Originalfixture-Hashes bestanden; vorheriger Vollstand48/48 (82.86 s), GUI-freier
Core unverändert1/1 (1.54 s). Katalog weiterhin456/384/72.

PV nutzt nun denselben nativen Annuitätenpfad; SLN/SYD führen begrenzte
Double-Abschreibungen aus, auch über den echten SQL-Adapter. Benannte Argumente,
abgezinster Zahlungsstrom und vollständige fünfjährige Abschreibungsreihe im
bestehenden Core/DAO-Ziel unabhängig geprüft. Native_Financial-Dokument erweitert:
docs/NATIVE_FINANCIAL_FUNCTIONS.md. Keine vollständige Finanzbibliothek behauptet.
Gezielt3/3 (2.36 s), frischer Core1/1 (1.66 s), Paket/Build-ID/SQLite und vier
Originalfixture-Hashes bestanden. Letzter Vollstand48/48 (82.86 s); 456/384/72.

TableDef.Fields liefert nun echte sichtbare Main-Schema-Spalten ohne Nutzdaten-
Zeilenlesung: Count/Item/Default/Refresh/For Each und typisierte DAO.Fields/Field,
Name/OrdinalPosition sowie bereits bewiesene Type-Zuordnungen. Unbewiesene
Feldeigenschaften/Typen3251, Schemaänderung und nativer Rollback invalidieren
gehaltene Metadaten3420; TableDef.Fields neu abfragen. TEMP-Schatten von Tabelle
oder Schemakatalog3251. TableDef prüft ursprüngliche Workspace-Verbindung.
docs/DAO_TABLEDEF_FIELDS.md; vollständiges DAO-Metadaten-/Refresh-/Identitätsmodell
bleibt F121 offen. Gebündelter aktueller Vollstand48/48 (84.55 s), Paket/Build-ID/
SQLite und vier Originalfixture-Hashes bestanden; Core unverändert1/1 (1.66 s).
Katalog weiterhin456/384/72. Keine neuen CTest-Ziele.

TableDef.Field.Type wertet nun eindeutige gespeicherte Main-Schema-Herkunftstypen
schreibgeschützt aus; ununterscheidbares TEXT/FLOAT bleibt3251. Fehlerhafte,
doppelte oder fremde Feldmetadaten werden abgewiesen. Field.Size liefert bewiesene
feste Größen als Long, einschließlich LongBinary/Memo0 (Metadaten, keine aktuelle
Datenlänge); Textkapazität mangels erhaltener Länge3251. Recordset.Field.Size
nutzt denselben Größenhelfer für seinen weiterhin engeren bewiesenen Typumfang.
docs/DAO_TABLEDEF_FIELDS.md und DAO_BINARY_FIELDS.md aktualisiert.
Gezielt DAO1/1 (0.59 s), Paket/Build-ID/SQLite und vier Originalfixture-Hashes
bestanden; letzter Vollstand davor48/48 (84.55 s), Core unverändert1/1 (1.66 s).
F121 und vollständige Typ-/Speicher-/Access-Gleichheit offen; weiterhin456/384/72.

Neue MDB-Importe erhalten nun Typ-Herkunft und bestätigte VARCHAR-Längen1..255
in main._la_table_schema. Bestätigtes MySQL-TEXT wird dadurch als Memo erkennbar;
TableDef.Field.Size liefert erhaltene Textlängen als Long. Alte Projekte werden
nicht geraten/migriert. Frisch importierte Textspalten erzwingen ihre Grenzen
über eine tatsächliche native SQLite-CHECK-Funktion mit UTF-16-/NUL-Zählung;
Überlänge3163 propagiert in SQL/DAO. Null/Required bleiben getrennt.
57 originale Textlängen und3 Memo-Deklarationen auf Arbeitskopie geprüft;
Unicode-/NUL-Grenzen, fehlgeschlagene SQL-/DAO-Schreibzugriffe ohne neue Zeilen.
docs/SOURCE_TEXT_CAPACITY.md; vollständige Jet-/UI-/Migrations-/Formatgleichheit offen.
Aktueller Vollstand48/48 (85.63 s), Paket/Build-ID/SQLite und vier Originalfixture-
Hashes bestanden; GUI-freier Core unverändert1/1 (1.66 s). Weiterhin456/384/72.

Direkte physische Recordset.Field.Type/Size nutzen jetzt denselben geprüften
Main-Schema-Metadatenleser wie TableDef.Fields. Wertkarten liegen bei der geladenen
Datensatzfamilie, werden mit Clone geteilt und bei Requery erneuert; bestehender
Reload-Rollback-Guard stellt sie bei Fehler atomar mit den Datensätzen wieder her.
Keine UI-/Host-/VBA-Objekte oder SQL-Cursor im Metadatencache. Geschlossen-/Rollback-
Thread-/Verbindungsprüfungen bleiben. Reine TEMP-Quellen nutzen alten Fallback,
Main-TEMP-Schatten werden explizit abgewiesen. Query-/Frozen-Lineage und vollständige
Variant-/Single-Speichergleichheit bleiben offen. docs/DAO_TABLEDEF_FIELDS.md.
57 Textlängen/3 Memo-Typen auf echten Original-Arbeitskopie-Recordsets verifiziert.
Aktueller Vollstand48/48 (84.98 s), Paket/Build-ID/SQLite/vier Originalfixture-Hashes
bestanden; Core unverändert1/1 (1.66 s). Fester Katalog weiterhin456/384/72.

Native literale Tabellen-Standardwerte sind jetzt vom MDB-Export bis SQL, DAO
und gebundenen Formular-Neudatensätzen erhalten. AccessCatalog fordert
--default-values an; sicherer Literalcompiler akzeptiert endliche Zahlen,
Boolean (-1/0), Text und NULL, keine allgemeinen Ausdrücke. AutoNumber behält
seinen Generator. Quelle/normalisierte Defaults werden gespeichert. AddNew liest
Literale vor Update, CancelUpdate schreibt nichts. Form_Current sieht Tabellen-
Literale; ausdrückliche Control-Defaults gewinnen, Anzeige allein bleibt nicht-Dirty.
96 originale Deklarationen/SQL-Defaults/DAO-Pufferwerte auf Arbeitskopie verifiziert.
Teilweise alte Counter-Metadaten ohne source_type bleiben lesbar; Defaults lesen
unabhängig vom Typkatalog tatsächliche Main-DDL, keine Generatorauswertung.
docs/SOURCE_LITERAL_DEFAULTS.md. Dynamische Defaultausdrücke, Field.DefaultValue-
Eigenschaften/-Bearbeitung, vollständige Windows-Timing-/Formatgleichheit offen.
Aktueller Vollstand48/48 (85.12 s); danach nur Importhinweistext geändert,
Inspector1/1 (11.04 s), finales Paket/Build-ID/SQLite/vier Originalfixture-Hashes
bestanden. Core unverändert1/1 (1.66 s). Weiterhin456/384/72, keine neue Testzielzahl.

Neue-Tabelle-Dialog unterstützt nun echte literale Standardwerte mit
Typprüfung und normalisiertem Boolean0/-1. Validierung erfolgt vor CREATE TABLE;
Fehler lassen Dialog/Feldliste geöffnet, Schließen nur nach Erfolg. Auto-id und
das begrenzte Erstellungsschema bleiben; vollständiger Access-Tabellendesigner offen.
Bestehender Mainwindow-Fall prüft abgewiesenen Ausdruck ohne Tabelle/Datenverlust,
anschließende gültige Erstellung und tatsächliches DEFAULT-VALUES-Ergebnis.
UI1/1 (2.82 s), aktuelles Paket/Build-ID/SQLite und vier Originalfixture-Hashes
bestanden. Letzter Vollstand davor48/48 (85.12 s), Core unverändert1/1 (1.66 s).
docs/SOURCE_LITERAL_DEFAULTS.md; weiterhin456/384/72.

Push-Status 2026-11-05: letzter verifizierter Arbeitsbranch-Remote ist
806d1581802b6026ee60688021041c5ee8b40f76 (Repository weiterhin privat).
Geprüfte Literaldefaults liegen lokal in35a68112a70b5460ce594b705439ec8522156858;
ihr identischer GitHub-Git-Tree b12d338dfa3edb269ed84eb7534f3f7a05ac9d6c und
Commit8943d06ee6ca2a92e389fc58754e00a08df3ccfc wurden erstellt, der Branch konnte
wegen GitHub-Serverfehlern bei GraphQL-/REST-Ref-Updates nicht bewegt werden.
Auch Contents-API scheiterte; der Remote blieb unverändert. Nicht als gepusht
behaupten. Private externe Checkpoint-/Patch-Sicherung unter linaccess-onboarding;
keine öffentlichen Artefakte, kein Force-Push, keine Recovery-Branch wurde erzeugt.

DAO.Field.DefaultValue ist jetzt für unterstützte native Literale schreibgeschützt
als String lesbar: direkte Main-Recordsets/Clone sowie geprüfte TableDef-Felder.
SQL-Text wird in kanonischen VBA-Ausdruck umgequotet, explizites NULL ergibt Null,
tatsächlich fehlender Default leeren String. Unbewiesene Projections/Generatoren,
Binary-Defaults, Mehrzeilen-/NUL-Text und Ausdrücke>255UTF-16-Einheiten3251.
Writes383; vollständiges TableDef-Editing/Windows-Originalschreibweise offen.
Alle96 Original-Defaultausdrücke auf Arbeitskopie zusätzlich gelesen; gezielt3/3
(10.48 s), aktuelles Paket/Build-ID/SQLite und vier Fixture-Hashes bestanden.
Letzter Vollstand48/48 (85.12 s), Core unverändert1/1 (1.66 s);456/384/72.
docs/SOURCE_LITERAL_DEFAULTS.md. GitHub-Ref-Push weiterhin nicht bestätigt;
vor nächstem Ref-Update tatsächlichen Remote und private Sichtbarkeit prüfen.

### Native process working directory (2026-11-05)

ChDir plus CurDir/CurDir$ use real Linux process cwd, not a mock or per-drive
Windows adapter. Optional empty Drive accepted, nonempty3251, Null94; bare
locals/procedures/host identifiers retain priority. Failed ChDir preserves cwd;
RmDir refuses the current directory by filesystem equivalence. Existing core
cases restore cwd after unique temporary-directory operations. No SQL CurDir
registration yet. docs/NATIVE_DIRECTORY_STATEMENTS.md; F040 remains open.
Desktop48/48 (83.85 s), fresh headless1/1 (1.79 s), current package/build-ID/
SQLite3.54.1 and original four hashes pass.456/384/72 stays fixed. Private GitHub
branch still806d158; CAS ref update again returned UNKNOWN, no published-head
claim for subsequent private immutable backups.


SQL CurDir extension: compiler rewrites parenthesized CurDir to the same native
VBA implementation via a non-deterministic SQLite callback. Prepared SELECT
reexecution after an actual directory change and DAO String value passed.
Targeted core/DAO/query3/3 (2.33 s), updated package/build-ID/SQLite and all four
hashes passed; no full rerun claimed after bounded SQL adapter change. Full
baseline48/48 (83.85 s), headless runtime1/1 (1.79 s).456/384/72 unchanged.

Native Filters.Delete now supports an explicit one-based numeric filter selector
including named filter. Actual shared-list mutation feeds the real Qt picker;
active-modal mutations refused, omitted/name/object variants3251 rather than
invented default semantics. Native selection-index adjustment is documented
without Windows parity claim. Existing layouts1/1 (10.58 s), current package/
build-ID/SQLite3.54.1 and four hashes pass. Last full48/48 (83.85 s), headless
runtime unchanged1/1 (1.79 s);456/384/72. docs/NATIVE_FILE_DIALOG.md.

Native SysCmd status actions1..5 now drive the actual main-window status bar /
Qt progress widget, with standalone document fallback. Application.SysCmd plus
unqualified statement/function calls and named arguments share native execution;
module procedures retain resolution priority. Successful operations return Null,
SetStatus removes the meter; invalid updates preserve its value. Other actions
3251, no Access version/registry/ActiveX fiction. docs/NATIVE_ACCESS_STATUS.md.
Targeted core/layout2/2 (13.17 s), fresh GUI-free1/1 (1.75 s), current Debian/
build-ID/SQLite3.54.1 and four original hashes passed. Last full baseline48/48
(83.85 s), not rerun after this extension. F068 open;456/384/72 unchanged.

Owned query/form/report designers now track actual modified/new state, reset only
on successful validated save, and replace stale saved messages after editing.
Real input and field-list models are tracked; query result models are excluded.
Existing form geometry display changes block signals. No new close-time prompting,
complete Undo/equality/name-reservation claim. docs/DESIGNER_SAVE_KEYS.md; F080 open.
Existing mainwindow1/1 (2.71 s), current package/build-ID/SQLite3.54.1/four hashes pass;
last full48/48 (83.85 s), unchanged runtime headless1/1 (1.75 s);456/384/72 fixed.

SysCmd GetObjectState10 now consumes the shared checked native project/window
registry for tables/queries/forms/reports. Open1, tracked design-modified2 and
registered new4; data-record dirty does not set design bit. Closed/nonexistent0,
other object types3251, complete anonymous-new-name/COM/VBE/Windows state parity
open. Existing IsLoaded uses the same state reader with original generation and
Qt owner guards. Actual query designer1→3→1→closed0 and pending query-record1
passed. docs/NATIVE_ACCESS_STATUS.md. Targeted core/mainwindow/layout3/3 (15.13 s),
fresh headless1/1 (1.76 s), current Debian/build-ID/SQLite3.54.1/four hashes pass.
Last full baseline48/48 (83.85 s), catalog456/384/72 unchanged.

DoCmd.Hourglass and Screen.MousePointer0/11 now control/observe a real owned Qt
hourglass cursor, with automatic outer-native-execution success/error cleanup and
nested-call retention. Original native raster cursor; no Windows pixel/animation
claim. Foreign overrides are not popped; obscured ownership3251 and full stack
interop remains open. docs/NATIVE_HOURGLASS.md. Existing core/mainwindow/layout3/3
(14.78 s), fresh headless1/1 (1.67 s), current package/build-ID/SQLite3.54.1/four
hashes pass. Last full baseline48/48 (83.85 s); F068 open,456/384/72 unchanged.

Application.Echo and DoCmd.Echo now control actual native Qt document-surface
repainting with documented named-argument spellings. New ordinary forms are
suppressed; modal/PopUp documents/dialogs and status bars excluded; data/VBA/input
continue. Echo True restores prior per-surface update state via weak pointers.
Echo stays off across procedure returns until explicit True; Hourglass cleanup
remains separate. Optional native status text preserves a progress meter. No
unqualified/SQL Echo or complete Windows paint/invalidation parity claim.
docs/NATIVE_ECHO.md. Existing mainwindow/layout2/2 (13.25 s), then full48/48
(85.02 s), current Debian/build-ID/SQLite3.54.1/four hashes pass. GUI-free runtime
unchanged from fresh1/1 (1.67 s). Fixed456/384/72; F068/F080 still open. Private
GitHub branch update remains unconfirmed; immutable private backups are not pushes.

Native anonymous/owned QueryDef.Type now exposes Integer statement-kind metadata
recorded by compiler tokens before TOP/SELECT-INTO/JOIN-UPDATE rewrites. SELECT0,
DELETE32, UPDATE48, append64, make-table80, DDL96, outer UNION128. Original named
imported types explicitly3251 because passthrough/Connect metadata is unavailable;
empty/unclassified3251, assignment383. Native DAO type enum constants defined
without pretending to execute unsupported kinds. docs/DAO_QUERYDEF_TYPE.md.
Existing real actions CREATE/INSERT/UPDATE/SELECT/INTO/DELETE/ALTER/DROP passed;
targeted core/DAO/query3/3 (2.25 s), fresh headless1/1 (2.75 s), current Debian/
build-ID/SQLite3.54.1/four hashes pass. Last full baseline48/48 (85.02 s), no full
rerun after extension. F068/F121 still open;456/384/72 fixed.

Native LSet/RSet plus writing Mid(stringVariable,start[,length]) now mutate actual
String slots through shared typed/ByRef assignment. UTF-16 widths, bounded
padding/truncation, documented Mid examples, constants, complete pairs and refusal
without mutation covered. Actual arrays named Mid retain indexed assignment
priority. Simple existing String/Variant-String targets only; no UDT memory copy,
fixed-length declarations, MidB, general lvalues or arbitrary unpaired-surrogate
claim. Unsupported/range boundaries explicit; docs/NATIVE_STRING_MUTATION.md.
Desktop core1/1 (1.82 s), fresh headless1/1 (1.66 s), current Debian/build-ID/
SQLite3.54.1/four original hashes pass. Last full baseline48/48 (85.02 s), not
rerun after language extension. F040/F068 open;456/384/72 stays fixed.

Native SQL/VBA standard-module scalar bridge implemented in native_vba_sql.hpp,
installed through the verified Jet SQL adapter path. Real SELECT/UPDATE calls use
the existing native interpreter and per-connection module/scalar Static state.
Private/Friend/class/ambiguous definitions and builtin/reserved-name overwrites
are excluded/refused. No permanent QSqlDatabase or Qt form host in callbacks;
weak driver/registry, SQLite-owned tokens, native-handle/thread checks and fresh
connection generations. GUI/DAO host APIs and non-null object/general-array
module/static/public-global state refused; no fabricated context/results.
Initial arguments NULL/finite numeric/UTF8 text, results Null/Boolean/Number/String.
Currency/Date/BLOB/object/array/Error/Empty result transport and complete Jet result
metadata are NOT implemented. No pauschal compatibility or Windows oracle claim.
Statement-generation trace ignores SQL text and logs nothing; metadata scope avoids
per-row source reload. Prepared source changes, source savepoint rollback, TEMP
archive shadows, removed/ambiguous sources, reopened state and real UPDATE passed.
Trace-hook interoperation/concurrent source editing/full VBA transaction-state/
large-project performance remain open. docs/NATIVE_VBA_SQL.md. Full production48/48
(87.35 s), final source guards DAO/original-query2/2 (8.12 s), subsequent actual
UPDATE DAO1/1 (0.74 s), current package/build-ID/SQLite3.54.1/four hashes pass.
GUI-free core runtime unchanged fresh1/1 (1.66 s); DAO SQL bridge uses Qt Core/SQL
without Widgets. Original corpus still78/100; fixed456/384/72, F040/F068/F121 open.
Next useful bridge stage: proven declared Currency/Date return metadata/transport
with stale-signature guards; do not simply convert unknown variants to REAL/TEXT.
GitHub branch still needs verified non-force CAS push; immutable backups are not
published branch updates. Maintain privacy and original test-copy/hash policy.

Declared SQL/VBA Currency/Date transport added using source-only native procedure
signature API (nativeVbaProcedureSignature) and contextual compiler hints. Currency
uses existing checked CURRENCY_REAL limits, Date canonical DATETIME. Live native
result guards reject changed return signatures on retained prepared SQL3251; no
observed-value type guesses. Actual DAO Currency5/Date8/Long4 values and Boolean,
exact Currency sum, typed INSERT, no-effect metadata preparation and signature
restoration passed. Unknown Variant Currency/Date, Single/full DAO scalar metadata,
argument provenance and full compound/nested/view typing remain open. Typed-function
make-table initially refused3251; the subsequent bounded typed make-table stage
below supersedes that restriction for simple SELECT sources.
Unaliased guarded function projections retain original-expression native labels;
no Windows Expr1000 naming claim. docs/NATIVE_VBA_SQL.md. Full48/48 (85.94 s),
fresh headless1/1 (2.36 s), final label fix afterward DAO/query/original3/3 (8.36 s),
updated Debian/build-ID/SQLite3.54.1/four hashes pass. Original78/100 unchanged,
no Access oracle; F014/F068/F121 still open,456/384/72 fixed. Next natural work:
proven make-table output types or read-only/reentrant DAO inside SQL callbacks,
with real object/connection lifetime guards; do not remove restrictions blindly.

Bounded typed SELECT INTO now uses token-derived SELECT source metadata, explicit
CURRENCY_REAL/DATETIME columns and INSERT SELECT inside the existing DAO savepoint.
Output metadata is prepared without source-body execution. Actual three-row money/
date copying and failure rollback pass in the existing DAO target. Unknown output
columns are untyped; typed compound/UNION/TOP/ranked/RIGHT JOIN and duplicate labels
refuse3251. No AutoNumber/keys/defaults/validation or full Jet metadata claim. Source
function state effects are not transactionally rolled back. F014/F130 remain open;
fixed456/384/72 unchanged. See docs/NATIVE_VBA_SQL.md.

Typed make-table milestone (2026-11-05): full current desktop48/48 passed
(87.98 s); Debian payload/build-ID/pinned SQLite3.54.1 and all four original
fixture hashes passed. GUI-free core runtime unchanged from fresh1/1 (2.36 s).
Original-query coverage remains78/100 without Access result oracle. No new
CTest target or catalog ID; fixed456/384/72. Native SQL scope: NATIVE_VBA_SQL.md.

Native unqualified FileCopy now copies real binary files and replaces existing
regular destination files. Missing-source failure preserves the destination;
identical files/final symlinks/Windows drive mapping refuse3251. No full Windows
sharing/error/metadata equivalence claim. Scope: docs/NATIVE_DIRECTORY_STATEMENTS.md.
Existing desktop core1/1 (1.77 s), fresh headless core1/1 (1.69 s), latest Debian/
build-ID/SQLite3.54.1/four fixture hashes pass. Last full typed make-table48/48
(87.98 s); no new target or fixed ID,456/384/72 unchanged.

Value List ComboBox/ListBox AddItem(Item,Index)/numeric RemoveItem(Index) now
mutate real Qt models, including named VBA calls, preserving clean record/bound
value. Stable Form context only; active popup/selection, headers/ordinal binding,
selected combo-row removal, text-key removal and empty/Null-only edge cases remain
bounded refusals. Full scope: docs/COMBO_LIST_PROPERTIES_RUNTIME.md. Final existing
layout1/1 (10.66 s), current Debian/build-ID/SQLite/four hashes pass; headless runtime
unchanged FileCopy1/1 (1.69 s), last full48/48 (87.98 s). F162 stays open;456/384/72.

SQL/VBA registry now reserves source aliases of registered linaccess_jet_ adapters
as well as actual SQLite names. This prevents public VBA FV As Currency from
providing false Currency metadata to builtin SQL FV. Actual builtin result20
and absent false Currency proof pass in existing DAO/query/original3/3 (7.79 s);
Debian/build-ID/SQLite/four hashes pass. Core runtime unchanged FileCopy1/1
(1.69 s), last full48/48 (87.98 s). Complete name-priority oracle remains open.
Scope: docs/NATIVE_VBA_SQL.md; fixed456/384/72 unchanged.

Native NPer(Rate,Pmt,Pv,[Fv],[Type]) now computes real fractional financial
periods in VBA and SQL with qualified/named VBA calls, zero-rate balance and
stable logarithmic nonzero-rate calculation. No finite unique solution5;
bounded rate/timing3251; full Windows numeric/error oracle remains open. Scope:
docs/NATIVE_FINANCIAL_FUNCTIONS.md. Existing core/DAO/query3/3 (2.76 s), fresh
headless1/1 (1.77 s), updated Debian/build-ID/SQLite/four hashes pass. Last
broader full48/48 (87.98 s); no new target/ID,456/384/72 fixed.

Native financial expansion (2026-11-05): NPV(Rate,Values) and
MIRR(Values,Finance_rate,Reinvest_rate) execute real one-dimensional Double arrays
with arbitrary lower bounds, mixed cash-flow signs and bounded finite rates.
NPV first-flow period1, MIRR first-flow period0; no array flattening or SQL array
transport. Other element/rank/empty/single-sign NPV semantics stay bounded3251.
DDB executes real VBA/SQL declining-balance depreciation with default factor2
and actual salvage cap, whole life/period up to32767. No full fractional/asset
range or Windows error/numeric oracle claim. Scope: docs/NATIVE_FINANCIAL_FUNCTIONS.md.
Current full desktop48/48 (85.94 s), fresh GUI-free core1/1 (1.67 s), updated
Debian/build-ID/pinned SQLite3.54.1/four original fixture hashes pass. Original
corpus78/100; no new target or ID. F068 stays open, fixed456/384/72.

Native Rate/IRR now share an actual balance/analytic-derivative root kernel with
default guess0.1, bounded twenty iterations, residual validation and explicit
nonconvergence5. Rate supports VBA/SQL whole annuity periods1..32767; IRR supports
actual one-dimensional Double arrays with documented Values/Guess names and
positive/negative flows. No direct SQL array transport or full Windows root/
error/precision/selection oracle claim. Scope: docs/NATIVE_FINANCIAL_FUNCTIONS.md.
Existing core/DAO/query3/3 (2.52 s), fresh headless1/1 (1.73 s), updated Debian/
build-ID/SQLite/four hashes pass. Last broader full is preceding financial
milestone48/48 (85.94 s). No new target/ID; F068 open,456/384/72 fixed.

Native IPmt/PPmt now calculate actual interest/principal portions in VBA/SQL,
sharing stable annuity factors with existing FV/PV/Pmt and advancing balances
through earlier periods. Beginning/end timing, first beginning-payment interest0,
zero-rate principal and actual SQL credit rows pass. Whole periods1..NPer up to
32767, rate greater than-1 and timing0/1 only; full Windows rounding/error/domain
oracle stays open. Scope: docs/NATIVE_FINANCIAL_FUNCTIONS.md. Final existing
core/DAO/query3/3 (2.68 s), fresh headless1/1 (1.79 s), updated Debian/build-ID/
SQLite/four hashes pass. Last full is fb63120 financial milestone48/48 (85.94 s),
preceding roots/components. No new target/ID; F068 open,456/384/72 unchanged.

Calculated query headers now preserve native source-expression labels before
scalar/concat/backend helper rewriting. Explicit/implicit/plain/quoted aliases,
mixed-star projections and first-leg UNION names stay supported. Internal helper
names are removed from supported generated headers. No Windows Expr1000, duplicate
numbering or complete alias-grammar claim. Scope: docs/QUERY_COLUMN_LABELS.md.
Final full current desktop48/48 (86.70 s), updated Debian/build-ID/pinned SQLite/
four original hashes pass. Qt-only label change leaves GUI-free runtime unchanged
from fresh payment-components1/1 (1.79 s). Original78/100; no Access oracle.
No new target or catalog ID; F068/F080/F148 open,456/384/72 fixed.

Native table Ctrl+H now offers real atomic Alle ersetzen for literal whole-field
text in a visible non-key column of a keyed writable table. Current filtered/
sorted rows are fetched; Qt Unicode case matching, NULL preservation and current
primary-key restoration are implemented. Later SQL failure rolls back the batch
and clears only its own staged replacements; qualified base revertRow bypasses
the saving guard that previously blocked revertAll. Unrelated drafts are refused
and preserved. No full key/query/wildcard/locale/concurrency/Undo/Access oracle
claim. Scope: docs/TABLE_REPLACE_ALL_RUNTIME.md. Final existing mainwindow1/1
(2.92 s), current Debian/build-ID/SQLite/four hashes pass. Last broader full is
preceding labels48/48 (86.70 s), unchanged fresh headless runtime1/1 (1.79 s).
No new target/ID; F373/F160/F168 stay open,456/384/72 fixed.

Native Ctrl+Minus/DeleteRecord now stages multiple completely selected stored
table records, including noncontiguous rows. Partial-cell selection, mixed drafts,
keyless/new rows and reentrant/context changes stay refused. Persistent anchors
and weak view/model guards preserve identity/lifetime. Existing Save commits the
set atomically; a later FK failure retains all SQL rows and pending choices.
Native VBA SelectAllRecords/DeleteRecord now executes the supported selection.
No full Access confirmation/event/Undo/concurrency/selection oracle claim. Scope:
docs/TABLE_MULTI_DELETE_RUNTIME.md. Final targeted mainwindow1/1 (2.80 s), full
current desktop48/48 (84.19 s), updated Debian/build-ID/SQLite/four hashes pass;
core unchanged fresh GUI-free1/1 (1.79 s). No new target/ID; F168/F148 open,
fixed456/384/72.

Native Name sourceExpression As destinationExpression now uses actual Linux
renameat2/RENAME_NOREPLACE for files/directories within one filesystem. No target
overwrite fallback; existing target58. String/bracket/parenthesis-aware As
parsing and local keyword-named assignments are preserved. Source links/special
files, current-directory ancestors, cross-filesystem/Windows mapping and full
sharing/error oracle stay bounded. Scope: docs/NATIVE_DIRECTORY_STATEMENTS.md.
Final desktop core1/1 (1.78 s), fresh GUI-free core1/1 (1.69 s), updated Debian/
build-ID/SQLite/four hashes pass. Last broader full is preceding multi-record
48/48 (84.19 s); no new target/ID, F040 open,456/384/72 fixed.

Native Scripting.FileSystemObject now supports real FileExists, FolderExists,
BuildPath and GetAbsolutePathName through local CreateObject/New and qualified
typed locals/parameters/results. TypeOf uses explicit native identity; wrong
classes13 leave the existing target intact. Scripting type validation/defaults
share a closed canonical family rather than guessing identity from TypeName.
No retained host/database handle, fabricated filesystem state, general COM/
reference resolution, drive/UNC/wildcard mapping or full FSO contract. SQL global
object-state guards remain strict. Scope: docs/NATIVE_FILE_SYSTEM_OBJECT.md.
Final full desktop48/48 (85.42 s), fresh headless1/1 (1.69 s), updated Debian/
build-ID/SQLite/four hashes pass. Original78/100; no Access result oracle.
No new target/ID; F040/F042/F091 open,456/384/72 fixed.

Native FSO GetFileName/GetBaseName/GetExtensionName/GetParentFolderName now
perform actual lexical extraction for bounded Windows/Linux paths, including
UTF-8 and nonexistent paths, without filesystem access or drive mapping.
Unvalidated root/dot/trailing-separator edge cases remain explicit3251. Scope:
docs/NATIVE_FILE_SYSTEM_OBJECT.md. Desktop core1/1 (1.73 s), fresh headless1/1
(1.72 s), current Debian/build-ID/SQLite/four hashes pass. Last broader full
48/48 (85.42 s) precedes this increment. No new target/ID; F091 remains open,
fixed456/384/72.

Native FSO GetFile(FileSpec) returns an independent real File object with
normalized absolute Path, Name and live Size. Actual regular Linux files only;
no host/stream/database retention. Long sizes and larger exact Double sizes
are distinguished; above exact range6. Symlinks, mutations, typed File binding,
full sharing/race/Windows metadata contract remain bounded/unimplemented.
Scope: docs/NATIVE_FILE_SYSTEM_OBJECT.md. Actual binary/UTF-8 metadata and a
retained object reading a resized2-GiB file pass desktop core1/1 (1.82 s) and
fresh GUI-free core1/1 (1.72 s). Current Debian/build-ID/SQLite/four hashes pass.
Last broader48/48 (85.42 s) predates both FSO increments. No new target/ID;
F091 open, fixed456/384/72.

Native Scripting.File binding now uses explicit native identity for qualified
locals/parameters/results. Path is the scalar/default value: CStr(file) reads
the path, VarType(file) is8 and TypeName(file) remains File. DateLastModified
reads actual filesystem modification time through the same shared local-time
conversion as FileDateTime, including UTC/EST changes. Property mutations
remain explicit3251 rather than claiming a writable implementation. Scope:
docs/NATIVE_FILE_SYSTEM_OBJECT.md. Targeted core/DAO2/2 (2.62 s), fresh
headless1/1 (1.66 s), current Debian/build-ID/SQLite/four fixture hashes pass.
Last broader48/48 (85.42 s) predates the FSO increments. No new target/ID;
F091 open, fixed456/384/72.

Native unqualified Kill pathnameExpression now unlinks an actual regular Linux
file only in active VBA branches; a local Kill identifier retains assignment
semantics. No shell/recursive deletion. Missing file53 vs parent76, Null94,
NUL52, read-only/denied70, wrong-kind75; wildcard/final-link/Windows sharing
semantics remain bounded. Scope: docs/NATIVE_DIRECTORY_STATEMENTS.md.
Targeted core1/1 (2.21 s), fresh headless1/1 (1.83 s); final full current
desktop48/48 (84.89 s), current Debian/build-ID/SQLite3.54.1/four original
hashes pass, covering all preceding FSO increments. Original query inventory
remains78/100 without Windows Access result oracle. No new target/ID; F040/F091
remain open, fixed456/384/72.

Native FSO GetFolder(FolderSpec) now returns a real typed Folder with Path/Name
and scalar Path default; trailing separators normalize. Files and SubFolders
produce native typed Files/Folders collections with fresh bounded Count and
For Each snapshots of actual regular files/directories. File.ParentFolder and
non-root Folder.ParentFolder return existing parent Folder objects. No open
handle/host retained; links/special entries, >10000 entries, root parents, Item,
recursive size, mutations and full Windows collection semantics stay bounded.
Scope: docs/NATIVE_FILE_SYSTEM_OBJECT.md. Expanded desktop core1/1 (1.86 s),
fresh GUI-free core1/1 (1.79 s), current Debian/build-ID/SQLite/four hashes pass.
Last broader48/48 (84.89 s) precedes this increment. No new target/ID; F091 open,
fixed456/384/72.

Native FSO CreateFolder(Path) creates one actual directory, returning a typed
Folder. Existing target58 preserves both directory and caller reference; a held
SubFolders collection observes the newly created UTF-8 folder. No recursive
creation, fabricated parent, overwrite or Windows ACL/sharing contract. Named
FSO parameters now follow IFileSystem IDL: GetFile FilePath, GetFolder FolderPath,
CreateFolder/GetFileName/GetAbsolutePathName Path. These supersede earlier
descriptive help placeholders; positional calls are unchanged. Scope:
docs/NATIVE_FILE_SYSTEM_OBJECT.md. Final desktop core1/1 (1.87 s), fresh
GUI-free core1/1 (1.77 s), current Debian/build-ID/SQLite/four fixture hashes
pass. Last broader48/48 (84.89 s) precedes folder/create increments. No new
target/ID; F091 remains open, fixed456/384/72.

Native FSO CopyFile/MoveFile/DeleteFile and File.Copy/Move/Delete/writable Name
now perform actual bounded single-file operations. Shared native_file_operations
keeps VBA FileCopy/Name/Kill and object methods on the same binary-copy, atomic
nonreplacement rename and unlink implementation. Overwrite=False58, live path
updates after File.Move/Name, directory destinations ending /, and read-only
Force policy are real. Trailing separators no longer hide a terminal folder
link from GetFolder or destination checks. Wildcards, links/special files,
cross-filesystem moves, full Windows ACL/sharing/metadata/race/error contracts
stay bounded. Scope: docs/NATIVE_FILE_SYSTEM_OBJECT.md. Final full desktop48/48
(86.15 s), fresh GUI-free1/1 (1.92 s), current Debian/build-ID/SQLite3.54.1 and
all four original hashes after the full run pass. A private native manual probe
confirms Force=False70 preserves a read-only binary test file and Force=True
unlinks it. Original queries78/100; no Windows result oracle. No new test
target/ID; F040/F091 open, fixed456/384/72.

Native read-only FSO OpenTextFile/File.OpenAsTextStream now return an owned
Scripting.TextStream with real Read/ReadLine/ReadAll/AtEndOfStream/Line/Column/
Close. Descriptor validation refuses terminal links/special files; ASCII only,
explicit non-ASCII/Unicode/write/create/count limits3251 rather than invented
conversion. A weak device/inode registry guards shared native copy-destination,
rename-source and unlink operations while a native reader is held; Close releases
it. No host/database retention or full Windows sharing/encoding/EOF oracle.
Scope: docs/NATIVE_TEXT_STREAM.md. Actual native VBA CRLF/LF/position/EOF/NUL/
held-delete/close/reopen workflow passes core/DAO2/2 (2.67 s), fresh headless1/1
(1.68 s), current Debian/build-ID/SQLite/four fixture hashes. Private manual
probe confirms hard-link alias mutation guard, Close release, EOF62 and failed
encoding-read position recovery. Last broader48/48 (86.15 s) precedes this
increment. No new target/ID; F040/F091 remain open, fixed456/384/72.

Native ASCII TextStream writing/appending is now implemented through FSO
CreateTextFile and OpenTextFile/File.OpenAsTextStream modes1/2/8. Real Write/
WriteLine/WriteBlankLines output, CRLF, append, Create and atomic no-overwrite
creation are supported. Payload validation rejects non-ASCII before writing a
prefix. The weak inode registry distinguishes readers/writers; descriptor
validation and sharing registration precede truncation, preserving bytes on
refused overlap. No O_TRUNC occurs before these guards. Unicode/locale codecs,
Folder.CreateTextFile, full Windows sharing/error/mode/position semantics and
filesystem transaction guarantees remain open. Scope: docs/NATIVE_TEXT_STREAM.md.
Final full current desktop48/48 (87.77 s), fresh GUI-free1/1 (2.23 s), both
builds zero compiler warnings; current Debian/build-ID/SQLite3.54.1 and four
original hashes after full tests pass. Current native hard-link/close/EOF/
encoding-recovery probe passes. Original inventory78/100; no Windows result
oracle. No new CTest target/ID; F040/F091 open, fixed456/384/72.

Native TextStream Format=-1/CreateTextFile Unicode=True now implement explicit
UTF-16LE BOM reading, writing and appending. Existing strict UTF converters are
shared in native_utf.hpp with the VBA string runtime, preserving scalar checks.
Actual output bytes include one BOM, correct non-BMP surrogate pairs and CRLF;
append validates rather than duplicates BOM. Read/column counts use UTF-16 units;
a split surrogate read3251 restores position. BOM-less/big-endian/odd-byte/default
code-page and Unicode read/Create=True remain explicit bounds, not guessed
conversion. Scope: docs/NATIVE_TEXT_STREAM.md. Final targeted core/DAO2/2
(2.77 s), full current desktop48/48 (89.99 s), fresh GUI-free1/1 (1.85 s), both
builds zero compiler warnings; current Debian/build-ID/SQLite3.54.1/four original
hashes pass. Private malformed-BOM append preservation, malformed-surrogate
position recovery and refused-create/no-side-effect probe pass. Original78/100;
no Windows result oracle. No new CTest target/ID; F040/F091 open, fixed456/384/72.

Native TextStream Skip/SkipLine now reuse the real checked read/line paths;
AtEndOfLine observes actual next CR/LF without consuming it. Its EOF contract
remains explicit3251. Folder.CreateTextFile now creates an actual plain child
name through the shared atomic-create/guarded-truncate/ASCII-or-UTF16LE pipeline;
path escapes remain3251. Actual VBA line-end/skip/NUL and Folder-created Unicode
readback cases pass core/DAO2/2 (2.69 s), fresh GUI-free1/1 (1.92 s), both builds
zero compiler warnings; current Debian/build-ID/SQLite/four original hashes pass.
Scope: docs/NATIVE_TEXT_STREAM.md. Last broader48/48 (89.99 s) precedes this
increment. No new CTest target/ID; F040/F091 open, fixed456/384/72.

Native Ctrl+P in active top-level forms now follows KeyPreview/KeyDown and can
be cancelled by KeyCode=0 before opening the genuine print dialog. The existing
form toolbar and report print command share a native execution/reentry/active-
window/modal/close guard; pending form controls are confirmed before the dialog.
Closed/cancelled documents are not rendered after acceptance. Report page setup
and the actual layout/report renderers remain in use. Embedded-form contexts,
complete form recordset/page/confirmation and Windows printer/event oracle
remain open. Scope: docs/FORM_PRINT_SHORTCUT_RUNTIME.md. Actual Qt/VBA form
cancellation and genuine form/report dialogs (rejected without physical jobs)
pass targeted layout/mainwindow2/2 (13.00 s); final full current desktop48/48
(86.11 s), current Debian/build-ID/SQLite/four hashes pass. GUI-free core is
unchanged since navigation1/1 (1.92 s). No new CTest target/ID; F148/F135 remain
open, fixed456/384/72.

## Chatwechsel am 2026-11-06

Alle Laufzeitänderungen bis6f77ac842e169ba34c144372dde6b682f82e50c0 sind
vollständig im privaten Repository gesichert. Der inhaltlich identische private
API-Sicherungscommit1a43d9c03c314bb551e16147450f4d9a5379c867 wurde am Chatwechsel
erfolgreich per direktem Git ohne Force auf development/access-compatibility
gepusht und per API zurückgelesen. Beide Historien sind mit
11751d4dc270e3c237514802426db8beb33d4b1f ohne Inhaltsänderung zusammengeführt.
Die früheren API-Ref-Fehler sind damit kein Beleg für einen aktuellen fehlenden
Push. Direkter Git-Zugang funktioniert wieder.

Kopierfertiger Fortsetzungsauftrag: NEXT_CHAT_PROMPT.md. Unverändert456 feste
Punkte/384 bestätigte Teilumfänge/72 offene Vollanforderungen; keine nachgewiesene
Gesamtparität. Letzter vollständiger Lauf48/48 (86,11 s), Debian/Build-ID/SQLite/
vier Originalfixture-Hashes bestanden. Kern seit letztem GUI-frei1/1 (1,92 s)
unverändert. Originalabfragen78/100 ohne Windows-Ergebnisorakel.

Tabellen-/Abfragedatenblatt-Druck/PDF war lediglich die nächste geplante Etappe.
Dafür wurde KEINE Implementierung begonnen. Es verbleibt kein uncommitteter
Laufzeitcode. Alte geschützte WIP-Patches dürfen nicht ungeprüft auf den neuen
Stand angewendet werden. Alle weiteren offenen Anforderungen stehen im festen
Katalog und den technischen Laufzeitdokumenten.


## Native datasheet output and new source comparison (2026-11-06)

Tables and readonly/editable saved-query documents now print actual checked
snapshots and export paginated PDFs atomically. Visible header order, filter/sort,
complete bounded fetching, dirty/binary/Unicode refusal, modal/project/model
lifetime guards and existing-file preservation are implemented. A Qt percentage-
width table initially protruded by its grid padding; explicit point width now
reserves both edges and a live frame-boundary guard prevents actual clipping.
Scope: docs/DATASHEET_PRINT_RUNTIME.md. No physical printer job was accepted.

Final current desktop48/48 (92.62 s), fresh GUI-free core1/1 (1.90 s), current
Debian payload/Build-ID/pinned SQLite3.54.1 and four original fixture hashes pass.
Existing mainwindow target additionally prints both actual recipe tables/39 rows
from the supplied identical MDB, checking text/Null and unchanged original hash.
No new CTest target or fixed ID; F080/F148/F087 remain open,456/384/72 fixed.

The newly supplied Kassa MDB is a separate source variant, not a replacement for
fixtures/raw: SHA-256 d9ff0f9d002b3ab59b3336a34e6c09d9c058d5b9b944b754259a78e277e77102,
31227904 bytes. Same exported schema;28/29 tables have identical CSV bytes.
The sole change is tblAnmeldung: upload0 rows, previous fixture1 row. No login
state or credentials are manufactured. Its isolated copy passes the existing
full table/cell regression and original query inventory:78/100, still without
Access execution/printing oracle. Source comparison logs/copies stay private
under /workspace/work/reference-20261106. Keks MDB and supplied Bridge ZIP are
byte-identical to the existing fixtures; originals remain unchanged.

Remote CI is currently blocked before job execution by GitHub account payment/
spending limits (c0ed5d1, run42221275422); no fresh remote matrix pass is claimed.
Codex Call was read from the actual feat/ask-via-call fd2cea7 checkout. The single
reference-access request received a call ID but ended timeout with no API answer
or confirmed decision; no delivery/conversation claim. User subsequently supplied
source databases in chat. No completion call: the overall parity task is open.


## Datasheet preview and VBA print commands (2026-11-06)

The same checked snapshot now drives a genuine modal QPrintPreviewWidget page
view with actual page count, page navigation and zoom. Its dialog owns the widget
and is destroyed before its printer. It offers no uncontrolled print-to-file
shortcut bypassing the atomic output controller. Failed rendering disables the
page view and preserves an explicit failure; source/project closure cannot be
reported as successful output. The original project/reentry/draft guards remain.

DoCmd.RunCommand acCmdPrint and Application.RunCommand acCmdPrintPreview address
the active table/query through the existing dispatcher, preserving caller Me.
Microsoft's actual AcCommand enumeration was read: Print340, PrintPreview54.
Unsupported form/report RunCommand coverage remains unchanged. The renderer now
uses a 72-dpi measurement device for true 9-point document metrics, independent
of screen DPI; an extracted PDF glyph-height assertion checks the actual result.

Existing core/mainwindow/layout3/3 (18.86 s), fresh GUI-free core1/1 (1.77 s),
current Debian payload/Build-ID/SQLite3.54.1 and four original hashes pass.
Actual VBA Print writes a private PDF, then continues in its caller; named
Application preview navigates real pages and continues with its own Tag intact.
The last single full desktop48/48 (92.62 s) is the preceding40c6b13 output commit,
not a new full-suite claim for this increment. No new target or ID;456/384/72 fixed.


## Native Recordset.Type and local platform validation (2026-11-06)

Recordset.Type exposes the selected native dbOpenDynaset2/dbOpenSnapshot4 kind
as a readonly Integer. SQL-result readonly policy is separate from opening kind;
Database/QueryDef factories now carry the actual mode instead of reinterpreting
readonly as Snapshot. Clones retain that mode. Frozen form cursor adapters lack
proven DAO type and continue to refuse3251. Invalid arguments450, writes383,
closed handles3420; a real field named Type remains available through rs!Type.
Existing DAO execution1/1 (1.02 s), rebuilt current app/all-target relink, current
Debian payload/Build-ID/SQLite3.54.1 and four original hashes pass. GUI-free runtime
is unchanged from1/1 (1.77 s). Full table-type/default/locking/ADO/Access-oracle
semantics remain open under F054/F055. Original32-module inventory has zero dot-
Type spellings; no additional original business execution is claimed.
Reference: Microsoft Recordset.Type (DAO), readonly Integer, and
QueryDef.OpenRecordset (DAO) query/dynaset default; current documentation is not
an Access2003 execution oracle.

Independent immutable e2d64c9 source copies now pass full local Debian12/Qt6.4
48/48 (101.29 s) and Ubuntu24.04/Qt6.4 48/48 (93.65 s), including native preview/
VBA print and actual recipe PDF proof. Both .deb packages pass payload/Build-ID/
SQLite3.54.1 checks, actual installation/backend smoke and four original hashes.
These are local container results, not GitHub CI successes; account payment/
spending-limit rejection remains. The newer Recordset.Type increment has targeted
host evidence, not yet a new full platform-matrix claim.456/384/72 unchanged.


## Checked OpenRecordset mode arguments and latest verification (2026-11-06)

Database/QueryDef/TableDef.OpenRecordset now share checked native Long mode
conversion: round-to-even before selection, native omitted/Missing mode2,
explicit Empty0 refused, Null94, invalid text13 and unsupported modes3251.
Database/QueryDef arity errors are450. This removes Database's truncating cast,
QueryDef's inconsistent exact-Double comparison and TableDef's narrower Integer
conversion. Existing native supported modes remain2/4; table-type1, full Access
workspace defaults, options/locking/ADO and complete error priority remain open.
Existing DAO target1/1 (0.97 s), current app/DAO build, current .deb payload/
Build-ID/pinned SQLite3.54.1 and four original hashes pass. All three supplied
upload hashes also remain unchanged. No new target/checkbox/ID; F055/F122 open.

The preceding418719d runtime passed a fresh Debian13 full48/48 (94.15 s).
Debian12 and Ubuntu24.04 immutable418719d follow-ups rebuilt the actual app/DAO,
passed DAO1/1 (1.55 s /1.06 s), regenerated and installed their matching packages,
verified SQLite3.54.1 and all four original hashes. Their preceding e2d64c9 full
suites remain48/48 each; the final bounded mode-conversion increment above has
host-targeted evidence rather than an invented repeat of those full suites.
GitHub account CI remains blocked; local container results are distinguished.


## QueryDef.Fields without executing query bodies (2026-11-06)

Supported prepared SELECT definitions now expose native DAO.Fields/Field metadata:
Count, zero-based/name Item, enumeration, Name and OrdinalPosition, plus Type/Size
only where the existing native type proof establishes them. Repeated references
preserve identity while held. Preparation does not step rows or invoke native VBA
function bodies; the existing DAO regression checks the actual function-call
counter and unchanged input rows. Action/make-table definitions and ambiguous
field names are refused3251. Query field DefaultValue remains explicitly
unproven3251; unknown types/sizes and other metadata are not guessed.

Collections/fields hold their QueryDef owner but retain no live cursor. SQL
replacement, closure, main-schema changes and workspace rollback invalidate
held metadata3420. Fresh preparation also compares result names/types, rather
than caching UDF declarations forever. Unsupported dependency/context preparation
is still an explicit boundary; full Jet metadata/error precedence and a Windows
Access2003 oracle remain open under F055/F121/F122. No catalog ID or checkbox
changes:456 total,384 scoped confirmations,72 open full requirements.

Nachweis dieser QueryDef.Fields-Etappe: bestehendes DAO-Ziel1/1 (0,96 s), anschließend vollständige Debian13-Desktop-Suite48/48 (98,29 s). Aktuelles Debian-Paket/Build-ID/SQLite3.54.1 und vier Originalfixture-Hashes bestanden. Keine neue Distributionsmatrix/GitHub-CI behauptet.


## Native metadata binding and TableDefs.Refresh (2026-11-06)

QueryDef.Fields and Database.TableDefs now expose the checked optional Index
signature; TableDefs.Item exposes its required Index signature. Actual typed VBA
reads query fields using named arguments, enumerates DAO.Field objects and checks
Currency Type/Size without opening a rowset. TableDefs.Refresh reads the real
current catalog and prunes dead/invalid cached items while preserving held valid
item identity. It neither creates tables nor revives invalid TableDefs. Extra
Refresh arguments are450. The collection now retains/checks the original native
workspace before member access and enumeration; cross-thread access is refused
3251 before querying Qt SQL. Existing close/replacement guards apply. This does
not solve unobserved drop/recreate identity or full Access ordering/refresh.
Existing DAO target only; F055/F121 remain open,456/384/72 unchanged.

Metadatenanbindung: vorhandenes DAO-Ziel1/1 (0,99 s) einschließlich tatsächlichem typisiertem VBA, benannten Zugriffen, Refresh und fremdem Thread bestanden; aktuelles Paket/Build-ID/SQLite3.54.1 und vier Originalhashes bestanden. Die Vollsuite48/48 (98,29 s) gehört zum vorherigen a00eb29; kein neuer Vollsuite-/Matrixlauf für dieses kleine Inkrement behauptet.


## Parameter collection refresh and named lookup (2026-11-06)

QueryDef.Parameters accepts the checked optional Index signature and explicit
Missing returns the actual collection. Parameters.Item supports named Index.
Refresh validates the live QueryDef and its currently compiled declaration
collection; unchanged source preserves cached Parameter identity and assigned
values. A replaced owned SQL source still uses the existing refresh/generation
path, so prior Parameter handles remain invalid rather than being revived.
Closed QueryDef refresh is3420 and extra Refresh arguments450. No rows or
parameter expressions execute during Refresh. This bounded implementation does
not discover undeclared Access context parameters or claim full COM/Jet metadata
refresh behavior. Existing DAO regression executes named VBA lookups, assigns
values, refreshes, checks identity/value preservation, executes actual SQL and
checks replaced/closed definitions. F055/F121/F122 remain open;456/384/72 unchanged.

Parameterfortsetzung: bestehendes DAO-Ziel1/1 (1,09 s), aktuelle App und Debian-Paket/Build-ID/SQLite3.54.1/vier Originalhashes bestanden. Named Index/Refresh und Werterhalt durch echtes VBA geprüft; letzter Vollsuite-Lauf bleibt a00eb29 mit48/48 (98,29 s). Kein neuer Katalogpunkt, keine vollständige Parameter-/Jet-Parität.


## Physical TableDef.RecordCount (2026-11-06)

A live local TableDef now returns checked readonly Long RecordCount from an
actual bounded SELECT COUNT(*) over main.<quoted table>. No rows are loaded into
a Recordset and a same-named TEMP table cannot redirect the count. Empty local
tables return0; current transaction writes/rollback are observed on the actual
connection. Extra arguments raise450 and assignment383; workspace/thread/
closed-table guards apply before SQL. Overflow is6. Linked-table count -1 is not
implemented because this native catalog exposes local physical tables, not a
linked-table adapter. Unobserved drop/recreate identity remains open under F121.

Existing DAO cases cover typed VBA/Long identity, empty/populated counts,
readonly/arity, same-connection insert/rollback and TEMP shadow separation.
The original Kassa DAO case compares TableDef counts for all29 imported physical
tables with independent SQL counts and real snapshot RecordCount. Original
business code and fixture bytes are unchanged. Reference (API contract, not an
Access2003 execution oracle):
https://learn.microsoft.com/en-us/office/client-developer/access/desktop-database-reference/tabledef-recordcount-property-dao
No new test target, catalog ID or checkbox;456/384/72 unchanged.

TableDef.RecordCount: DAO/originales Kassa-Ziel2/2 (53,71 s) bestanden, einschließlich unabhängiger Zählung aller29 Originaltabellen. Aktuelle App/Paket/Build-ID/SQLite3.54.1 und vier Originalfixture-Hashes bestanden. Bestehende unveränderte Syntaxinventur separat187/187 akzeptiert (1,12 s), keine Ausführungsaussage. Kein neuer Vollsuite-/Matrixlauf dieses Inkrements.456/384/72 unverändert.


## Bounded stored binary GetChunk (2026-11-06)

Recordset Field.GetChunk(Offset, Bytes), including named arguments, copies the
requested range of actual non-NULL native BLOB storage into an independent Byte()
Variant. Both arguments use checked Long conversion. Offset0 starts at the first
byte; a positive requested count exceeding the remaining bytes returns the actual
remaining bytes. Embedded0/255 bytes are retained; changing the returned array
does not change the stored field. The returned range is limited to16MiB. Existing
recordset/thread/transaction/current-record guards apply, and closure remains3420.

This deliberately supports stored browse-mode observations with an offset inside
nonempty data and a positive byte count. NULL/empty fields, offsets outside the
value, nonpositive counts, Edit/AddNew observations, Memo/encoding and unsupported
source types refuse3251 pending reference validation; their full Access result/
error precedence is not guessed. Wrong BLOB storage is13, arity450, Null numeric
arguments94 through existing checked conversion. This is copying from the native
loaded rowset, not streaming original MDB/OLE containers. AppendChunk, original
OLE/attachment extraction/roundtrip and full DAO binary semantics remain open.

Existing DAO binary cases verify named middle-byte extraction, first/tail ranges,
over-request truncation, isolated copied bytes, boundary/Null/arity/closed errors
and unchanged independent persisted SQL bytes. No new CTest target or catalog ID;
F016/F055/F121/F122 remain open and456/384/72 unchanged.
Reference: https://learn.microsoft.com/en-us/office/client-developer/access/desktop-database-reference/field-getchunk-method-dao
This current public contract is not an Access2003 execution oracle.

Binär-GetChunk: bestehendes DAO-Ziel1/1 (1,23 s), aktuelle App/Paket/Build-ID/SQLite3.54.1 und vier Originalfixture-Hashes bestanden. Nur gespeicherte positive native Binärbereiche; AppendChunk/OLE/Memo/Null/Randsemantik bleiben offen. Letzte Vollsuite a00eb29:48/48 (98,29 s), nicht diesem Inkrement zuschreiben.456/384/72 unverändert.

Native binäre Field.AppendChunk-Fortsetzung implementiert: Byte()-Chunks im Edit/AddNew-Puffer, erster Chunk ersetzt, folgende hängen an; je Feld getrennt, 16MiB, Update/CancelUpdate/Session-Reset und geschützte Mischverwendung. Genauer begrenzter Umfang in docs/DAO_BINARY_FIELDS.md. F016/F055/F121/F122 bleiben offen,456/384/72 unverändert.

Binär-AppendChunk-Nachweis: finales bestehendes DAO-Ziel1/1 (1,38 s), danach vollständiger aktueller Debian13-Desktop48/48 (98,95 s), aktuelles Paket/Build-ID/SQLite3.54.1 und vier Originalhashes bestanden. Byte-Value-Zuweisung bleibt separat geprüft; Mischverwendungsfehler bewahren den Puffer. Kein neuer Test/Katalogpunkt.
Unveränderlicher629b3c9-Folgecheck: Debian12 DAO1/1 (1,49 s), Ubuntu24.04 DAO1/1 (1,27 s), jeweils aktuelle App/DAO neu gebaut, passende Pakete erzeugt/tatsächlich installiert, SQLite3.54.1 und vier Originalhashes bestanden. Diese629-Nachweise enthalten noch nicht die anschließenden TableDef.RecordCount-/Binärchunk-Inkremente; keine neue Vollmatrix/GitHub-CI behaupten. Letzte früheren vollständigen lokalen Distributionssuiten bleiben e2d64c9 mit48/48 je Plattform.456/384/72 unverändert; kein Gesamtabschluss.

Recordset.Field.Name prüft jetzt wie die übrigen Feldmetadaten die tatsächliche Besitzer-/Verbindungs-/Rollback-Lebensdauer. Der Name bleibt ohne aktuellen Datensatz am EOF lesbar, nach Close aber3420. Vorhandenes DAO-Ziel1/1 (1,05 s), aktuelles Paket/Build-ID/SQLite3.54.1 und vier Originalhashes bestanden. Letzte Vollsuite41b193c:48/48 (98,95 s), kein neuer Gesamtlauf für diese Lebensdauerkorrektur.456/384/72 unverändert.

Recordset.Fields-Fortsetzung: benannte/geprüfte Indizes, kanonische Ergebnisnamen, gehaltene Collection-/Field-Identität auch bei Enumeration, Null94/fehlend3265/Arity450 und explizite Doppelspalten-Verweigerung3251. Bestehendes DAO-Ziel1/1 (1,01 s), aktuelles Paket/Build-ID/SQLite3.54.1/vier Originalhashes bestanden; Details docs/DAO_RUNTIME.md. Kein neuer Vollsuite-Lauf dieses Inkrements.
Unveränderliche41b193c-Folgeläufe bestanden Debian12 DAO1/1 (1,61 s) und Ubuntu24.04 DAO1/1 (1,21 s), jeweils aktuelle App/Paket/Installation/SQLite3.54.1/vier Originalhashes. Diese Belege enthalten TableDef.RecordCount/GetChunk/AppendChunk, nicht die danach folgenden Field.Name-/Recordset.Fields-Inkremente. Letzte komplette Host-Suite41b193c:48/48 (98,95 s).456/384/72 unverändert; GitHub-CI weiter gesperrt.

Datenblatt-Seiteneinrichtung: funktionierender gemeinsamer Qt-Geometriedialog, Sitzungslayout und explizite versionierte Projektpersistenz, native RunCommand PageSetup32, Papier/Ränder für Vorschau/Druck/PDF. Conditional Savepoint-Speicherung bewahrt konkurrierende Änderungen und Aufrufertransaktionen; ungültige gespeicherte Geometrie blockiert PDF bis zur expliziten Reparatur. Details docs/DATASHEET_PRINT_RUNTIME.md. Ziel1/1 (3,64 s), vollständiger aktueller Desktop48/48 (89,86 s) inklusive Originalrezepte2 Tabellen/39 Zeilen, Paket/Build-ID/SQLite3.54.1/vier Originalhashes bestanden. Keine neue Distributionsmatrix/GitHub-CI;456/384/72 unverändert, F080/F087/F148 bleiben offen.

Druckexemplare: reale Datenblatt-PDFs mit1..100 Kopien, sortierter/unsortierter und vorwärts/rückwärts Seitenfolge; native Qt-Treiberkopien getrennt von manuellen Kopien. Vorhandenes Ziel1/1 (3,99 s), aktuelle App/Paket/Build-ID/SQLite/vier Originalhashes bestanden. Letzte Vollsuitee8375e8:48/48 (89,86 s), dessen GUI-freier Kern1/1 (2,20 s). Unveränderlichee837-Folgeprüfungen Debian12/Ubuntu24.04:2/2 (6,65 s /5,65 s), passende Paketinstallation/Backend/vier Hashes; noch ohne Kopien-Inkrement. Details docs/DATASHEET_PRINT_RUNTIME.md.456/384/72 unverändert, keine physische Drucker-/Windows-Parität behaupten.

Auswahldruck: vollständige markierte sichtbare Datensätze als eingefrorener eigener Text-Snapshot, reale Tabellen-/Readonly-Abfrage-PDFs mit unabhängig geprüfter Ein-/Ausschlussmenge. Leere/teilweise Zellmarkierungen verweigern3251 ohne Ganzdatenblatt-Fallback. Ziel1/1 (4,15 s), aktuelle App/Paket/Build-ID/SQLite3.54.1/vier Originalhashes und drei Uploadhashes bestanden. Details docs/DATASHEET_PRINT_RUNTIME.md. Kein neuer Vollsuite-/Matrixlauf dieses GUI-Inkrements;456/384/72 unverändert.

DAO-DDL-Identität: beobachtete eigene Database/QueryDef.Execute CREATE/DROP und SELECT INTO führen monotone Main-Tabellengenerationen; alte TableDefs/Fields/physische Recordsets/Clones verweigern3420 statt neu erstellte gleichnamige Daten zu ändern. Gleiches SQL/Rootpage/Rowid, alter Schreibpuffer, frische Wrapper, No-op/Fehl-DDL, ALTER und Workspace-Rollback unabhängig geprüft. Externes/unbeobachtetes SQL, TEMP-only und volle Jet-Locks/Lebensdauer bleiben offen. Vollständiger aktueller Desktop48/48 (89,70 s), App/Paket/Build-ID/SQLite/vier Originalhashes bestanden; schließt Kopien/Auswahldruck und Originalrezepte ein. Details docs/DAO_RUNTIME.md.456/384/72 unverändert.

Jet-Rohbinärimport: ausdrücklich gewählter MDBTools-Hexcodec (kein Raten aus CSV), exakte BLOB-/NULL-Werte und Byteanzahl-Vorschau im realen Teil-/Gesamtimport. Laufend begrenzte Child-Ausgabe20MiB/Diagnostik64KiB,30s; großer Export bricht atomar ohne Quelldateiänderung ab. Technische echte Jet4-MDBs mit Jackcess5.0.4 unabhängig geschrieben/zurückgelesen, keine Geschäftsdaten und keine neue App-Java-Abhängigkeit. Decoder-/GUI-/DAO-Bytes/Typ/Chunks/Hash/Rollback geprüft. Ziel2/2 (12,26 s), vollständiger aktueller Desktop48/48 (91,47 s), neue Kassa-Kopie Import/Zellwerte und78/100 Abfragen, App/Paket/Build-ID/SQLite/vier Originalhashes bestanden. Details docs/JET_BINARY_IMPORT.md; keine OLE-Anzeige/Anhänge/Originalformat-Schreib- oder Windows-Parität behaupten.456/384/72 unverändert; kein neuer CTest.

DAO-Katalog-Count liefert nun für Parameters/QueryDefs/TableDefs/Workspaces/Databases nach öffentlicher statisch geprüfter PIA-Int16-Signatur200006 einen kontrollierten VBA Integer, keine generischen Double/Long. Echtes VBA VarType/NativeSubtype geprüft; Ziel1/1 (1,06 s), aktuelle App/Paket/Build-ID/SQLite/vier Originalhashes bestanden. Details/PIA-Hashes docs/DAO_COUNTS_RUNTIME.md; keine COM-Ausführung/Access2003-Oraclebehauptung. Letzte Vollsuite6cd40d1:48/48 (91,47 s).456/384/72 unverändert.

DAO.Relations: lesender tatsächlicher Importkatalog mit Name/Table/ForeignTable/Attributes und geordneten Fields.Name/ForeignName; Integer Count, benannte Indexargumente, echte VBA-Typen/TypeOf/Enumeration, gehaltene Identität und kontrollierte Invalidierung. Erzwungene Beziehungen werden mit allen tatsächlichen Main-Fremdschlüsseln und CASCADE-Aktionen abgeglichen; unbekannte Schlüssel/uneindeutige Metadaten verweigern3251 statt Teilkatalog. Bestehende Ziele2/2 (11,87 s), zusätzlich neue Kassa-Uploadkopie mit acht Originaldefinitionen/acht Feldpaaren bestanden. Details docs/DAO_RELATIONS_RUNTIME.md; Append/Delete/Indexes/Linked und vollständige Jet-/Windows-Parität bleiben offen.456/384/72 unverändert, kein neuer Testkatalogpunkt.

Aktuelle App/Debian-Paketprüfung einschließlich Build-ID/SQLite3.54.1 und vier Originalfixture-Hashes bestanden. Letzte Vollsuite bleibt6cd40d1:48/48 (91,47 s); kein neuer Vollsuite-/Distributionsnachweis für Relations.

Nicht-Primärindexrichtungen: ursprünglicher Parserverlust von DESC behoben; vollständige geordnete ASC/DESC-Feldlisten für CREATE/ALTER INDEX, ungültige Zusätze verweigern statt Normalisierung. Native Indexanlage erhält Richtung, Reihenfolge und Unique in derselben Importtransaktion. Optionale echte Jet4-Referenz ohne Datensätze mit zwei unabhängig zurückgelesenen Indizes, tatsächlicher Gesamtimport und PRAGMA index_xinfo geprüft; keine App-Java-Abhängigkeit. Details docs/SOURCE_INDEX_DIRECTIONS.md. Kein echter Primärindexname aus generierten MDBTools-Namen behauptet; DAO.Indexes/Primärindexrichtungen/Locale/Originalformat weiterhin offen.456/384/72 unverändert.

Indexrichtungsnachweis: vollständiger aktueller Debian13-Desktop48/48 (89,51 s) bestanden, einschließlich tatsächlicher Jet-Binär-/Indexreferenzen, Originalrezepte für Datenblatt-PDF, Originalkassenhandler und aller originalen Kassa-Indexfeldreihenfolgen/-richtungen. Schließt DAO.Relations und Integer-Count ein. Keine neue Distributionsmatrix/GitHub-CI für diese Inkremente;456/384/72 unverändert.

Aktuelles Debian-Paket/Build-ID/SQLite3.54.1, vier Originalfixture- und drei Uploadhashes bestanden.

DAO-Single-Werte: bestätigte physische Felder mit DAO-Typ6 liefern nun bei Value/Bang/Clone und im Edit/AddNew-Puffer echte Single-Werte; Schreiben verengt vor Puffermutation und persistiert binary32-Wert, Null/Empty und explizite Konvertierungsfehler bleiben geprüft. Echte technische Jet4-Float-Referenz mit unabhängig zurückgelesenen Werten, tatsächlicher Import/Metadaten/Update/Requery/Quellhash und VBA VarType/CSng im bestehenden Ziel2/2 (1,94 s) bestanden. Keine automatische Query-/Formularlineage und keine verlustfreie allgemeine Single-Importbehauptung: MDBTools1.0.0 %.8g kann einzelne Bitmuster verändern, belegtes Gegenbeispiel in docs/DAO_SINGLE_FIELDS.md. F014/F041/F055/F105/F122 bleiben offen,456/384/72 unverändert. Letzte Vollsuite4548ce7:48/48 (89,51 s), nicht diesem Typinkrement zuschreiben.

Aktuelle App/Paket/Build-ID/SQLite3.54.1 und vier Originalfixture-/drei Uploadhashes bestanden.

Originale Indexmetadaten: separater lesender nativer MDBTools-Helfer statt generierter CLI-Primärnamen. Strikt versionierte Namen/Kind/Flags/geordnete Felder/Richtungen, gegen tatsächliche Tabellenspalten und ursprünglichen Primärschlüssel validiert; atomare main._la_access_indexes-Persistenz. Neue Importe verwenden echte Nichtprimärnamen. Zusammengesetzte und einzelne DESC-Primärschlüssel bleiben echte native DESC-Indizes; DESC-AutoNumber und unbewiesene IgnoreNulls/Required-Konfigurationen verweigern explizit. Bestehende GUI/Import/Kassa-Ziele3/3 (15,27 s), abschließende erweiterte technische Jet-Referenz1/1 (0,96 s);20 echte Primär-/48 andere Originalnamen in29 Tabellen nachgewiesen. Details docs/SOURCE_INDEX_METADATA.md. Noch kein vollständiger DAO.Indexes-/Windows-Nachweis;456/384/72 unverändert. Build benötigt mdbtools-dev/pkg-config, Paket enthält geprüften lesenden Helfer.

Indexmetadaten-Grundlage vollständig verifiziert: aktueller Debian13-Desktop48/48 (85,91 s), neue echte Jet-Referenzen einschließlich einzelner und zusammengesetzter DESC-Primärindizes, Originalrezepte, Originalkassenhandler und Originalindexnamen bestanden. Aktuelles Paket/Build-ID für App und Metadatenhelfer, tatsächliches Laden des entpackten Helfers, SQLite3.54.1 sowie vier Originalfixture-/drei Uploadhashes bestanden. Kein neuer Katalogpunkt und keine vollständige DAO.Indexes-/Windows-Parität;456/384/72 unverändert. Lokale Debian12-/Ubuntu24.04-Folgeprüfung der neuen Build-Abhängigkeit folgt separat.

DAO.Indexes: lesender originaler Indexkatalog mit nativen Typen/TypeOf, Integer Count/Item/named Index/Refresh/Enumeration, echten Namen und Boolean Primary/Unique/Required/IgnoreNulls sowie geordneten Fields.Name/Long Attributes (dbDescending). Jeder Zugriff vergleicht den vollständigen tatsächlichen Main-Indexkatalog, Primärschlüssel/Rowid-Alias, Unique, Felder/Richtungen und BINARY-Collation. Fehlende Quellen, logische kind2/IgnoreNulls/unbekannte oder zusätzliche native Indizes verweigern3251 ohne Teilkatalog. Gehaltene Definitionen/Fields invalidieren kontrolliert nach Schema/Metadaten/Rollback, TEMP bleibt wirkungslos. Bestehende Ziele3/3 (11,69 s), zusätzlich neue Kassa-Kopie mit20 Primär-/48 anderen DAO-Indexidentitäten, aktuelle App/Paket/Build-IDs/Helfer/SQLite3.54.1/vier Originalfixture-/drei Uploadhashes bestanden. Details docs/DAO_INDEXES_RUNTIME.md. Keine Mutation/Foreign/Clustered/DistinctCount/Windows-Parität;456/384/72 unverändert. Letzte Vollsuite bleibt74d75c7:48/48 (85,91 s).

Unveränderliche74d75c7-Folgeprüfung abgeschlossen: Debian12/GCC12/Qt6.4 und Ubuntu24.04/GCC13/Qt6.4 jeweils5/5 gezielte Ziele (30,64 s /27,39 s), neues mdbtools-dev/pkg-config, Paketinstallation, Backend3.54.1, Originalhashes und tatsächlich installierter Helfer mit originalen Jet-Namen/Richtungen unter MDBTools1.0.0 bestanden. Diese Snapshots enthalten die Metadaten-Grundlage, noch nicht die folgende DAO.Indexes-Anbindung; keine GitHub-CI- oder Vollsuitebehauptung für die Matrix.

Exakter Single-Import: realer MDB_FLOAT-Rohwertzugriff im vorhandenen lesenden Helfer (--single-values), IEEE754-Bitmuster statt verlustbehafteter %.8g-Speicherung. Source-SHA256 vor/nach Scan und atomarem Import, echte SINGLE-Spalten/REAL-Affinität, CSV-Zeilen/NULL/8g-Abgleich,1M-Zellen/16MiB-Grenze. Rohwerte bleiben außerhalb des angezeigten SQL-Schemas. Gegenfallbits509022793, minimale Subnormale, maximaler endlicher Single, Brüche/+0/NULL nach unabhängiger Jackcess-Jet-Erzeugung im tatsächlichen Gesamt-/Teilimport und DAO bitgenau nachgewiesen. NaN/Infinity/negatives Null oder inkonsistente CSV-Daten verweigern atomar. Bestehende relevante Ziele5/5 (27,76 s), finaler Affinitätsguard im Importziel1/1 (1,05 s); aktuelle App/Paket/Build-IDs/SQLite3.54.1/vier Originalfixture-/drei Uploadhashes und tatsächlich entpackter Helfer mit bekannten Originalbits/Quellhash bestanden. Details docs/SOURCE_SINGLE_PRECISION.md. Legacy-CSV, Double/Decimal/Currency, generische SQL-/Default-/Query-Semantik und Windows bleiben offen;456/384/72 unverändert. Letzte Vollsuite74d75c7:48/48 (85,91 s), neue Matrix noch nicht behaupten.

Codex-Call-Rückfrage vom08.11.2026 zu echter Windows/Access2003-Referenz und GitHub-Actions-Kontofreigabe: Request3ea2a1b8-75c6-411c-8f85-3cb76c4f3571 wurde mit gespeicherter Call-ID angenommen; lokales Abfragefenster endete timeout, answers/confirmed_decisions leer. Keine Zustellung/Gesprächsannahme oder Freigabe nachgewiesen, keine Antwort erfinden, nicht erneut anlegen. Privater Zustand /workspace/work/reference-blocker-call-20261108/state.json; Ergebnis result.json. GitHub-Run42497872453 (a11b95e) hat keine gestarteten Build-Schritte; Annotation bestätigt fehlgeschlagene Kontozahlungen oder Ausgabenlimit. Keine Abrechnungseinstellungen geändert; unabhängige Entwicklung läuft weiter.

Unveränderliche a11b95e-Folgeprüfung: Debian12/GCC12/Qt6.4 und Ubuntu24.04/GCC13/Qt6.4 jeweils5/5 (28,31 s /27,15 s), Paketinstallation, Originalhashes/Backend und tatsächlich installierter --single-values-Helfer mit unabhängig bekannten Originalbits/SHA bestanden. Diese Snapshots schließen DAO.Indexes und Single-Rohpräzision ein, noch keine spätere Double-Korrektur. Keine GitHub-CI-/Vollsuitebehauptung für diese Matrix.

Double-Import bitgenau: vorhandener lesender Helfer --double-values, binäre64-Bit-Quelle statt %.16g-Rückwandlung. Gemeinsamer streng validierter access_floating_metadata-Codec für Single/Double, Classic-Locale-Formatprüfung ohne Double-Overflow bei der Dezimaldarstellung von MAX_VALUE. Original-Rohwerte aus der bereits geschützten Gesamtimport-Snapshotdatei; strukturleere Anlage scannt Daten nicht doppelt. Low-level API unterscheidet ausdrücklich caller-owned Csv (kein Dateizugriff aufgrund eines Loglabels) von NativeMdbBits; reale Gesamt-/Inspector-Kopien und Originalprüfungen wählen NativeMdbBits, fehlender Nachweis fällt niemals still auf CSV zurück. Dadurch bleiben technische Tests mit echten Schemas und eigens bereitgestellten CSV-Zeilen korrekt getrennt.

Unabhängiger Jackcess-Auszug beider Kassa-Testkopien:13 Tabellen/57 Double-Spalten,69405 Nicht-NULLwerte/514 NULLs, kein negatives Null/NaN/Infinity. Der bisherige Dezimalexport reproduzierte2404 numerische Zeilentupel in zwei Tabellen nicht exakt; neue native Kopien stimmen bitgenau mit allen unabhängigen Tupeln überein. Echte Jet-Extremwertreferenz, Gesamt-/Teilimport, DAO Type7/Size8, Herkunfts-/CSV-Guards und atomare Negativnull-Verweigerung geprüft. Aktueller korrigierter Debian13-Desktop48/48 (143,10 s), aktuelle neue Uploadkopie erneut29/29 und unabhängiger Auszug exakt; dort weiterhin78/100 Originalabfragen ohne Windows-Ergebnisorakel. App/Paket/Build-IDs/SQLite3.54.1/vier Originalfixture-/drei Uploadhashes und tatsächlich entpackter --double-values-Helfer mit bekannten Originalbits/SHA bestanden. Details docs/SOURCE_DOUBLE_PRECISION.md. Zuvor entdeckte technische CSV-Herkunftsfehler behoben, keine Schutzprüfung für Originale entfernt.456/384/72 unverändert; keine neue Matrix für Double behaupten.

Nach einmaligem lesendem Resume derselben Codex-Call-Anfrage3ea2a1b8-75c6-411c-8f85-3cb76c4f3571 wurde status=unresolved und eine normale finale LiveKit-Teilantwort „Genau“ zur Windows-Referenzfrage geliefert; keine konkrete Umgebung/Protokolle/Zugriffsinformation und keine confirmed_decisions. CI-Frage bleibt unbeantwortet. Beide externen Blocker bleiben offen, keine Freigabe erfinden oder neue Anrufanforderung daraus ableiten. Privates Ergebnis resumed-result.json im bestehenden Call-Verzeichnis.


## 2026-11-08 — tatsächliche Access-2003-Referenzen und AutoWert-Quellstand

Nutzerpriorität: allgemeiner Access-Ersatz für Linux, einschließlich eigener neuer
Datenbanken/Tabellen/Abfragen/Formulare/Berichte; die Kassa-Datei bleibt nur eine
Praxisreferenz. Keine Namen-/Geschäftslogik-Sonderbehandlung. Ganze fachliche
Arbeit bündeln und anschließend bestehende gezielte Prüfungen ausführen; keine
Kleinsttest-/Zwischenabschluss-Schleife. Katalog456/384/72 bleibt unverändert.

Nutzer führte Referenzpakete v3 und v4 tatsächlich mit Access11.0/DAO3.6 aus;
Original-SHA unverändert. Geschützte Ergebnisarchive liegen nur unter
/workspace/attachments/13b07826-03ad-4205-9109-c8f99a021e6b/ und
/workspace/attachments/bd2f7462-2959-4696-89e4-31f778f6929b/.
Das v4-Archiv heißt LinAccess-Gesamtreferenz-20261108-181858-f70a7afd.zip.
Kein weiterer Anruf/Testlauf allein aufgrund alter ungeklärter Call-Statuswerte.
Die technischen, von Originalmetadaten bereinigten Counter-Beobachtungen stehen
in tests/access2003_reference/counter-observations.json. Das private Gesamtarchiv,
Originalschema/Original-SQL/VBA/Geschäftsdaten dürfen nicht committed werden.

Alle sechs v4-Worker haben Abschlussmanifeste. Der Starter meldet fälschlich
partial_modules_failed wegen nicht ausgelesener Prozess-Exitcodes; seine
observed/error-Zählung zählt zudem JSON-Arrays falsch. Workerbeobachtungen direkt
lesen. Nicht alle Beobachtungen sind verwertbare Verhaltensreferenzen: sql.setup
und dao.setup scheiterten im COM-Interop beim Zuweisen der technischen Werte
(Int32->String); spätere SELECT-Ergebnisse können daher leere Tabellen betreffen.
DAO-Navigation hatte dann3021. Diese Ergebnisse nicht als vollständige SQL-/DAO-
Oracles nutzen. Decimal-CreateField gab3259, LongBinary COM-E_INVALIDCAST;
auch dies ist kein pauschaler Access-Unfähigkeitsnachweis. Kein erneuter breiter
Windows-Testlauf statt Programmierung; weitere Referenzen erst bei tatsächlichem
Bedarf gezielt vorbereiten.

Counter-Referenzen sind dagegen verwertbar: Nummer bei AddNew bereits vorhanden;
CancelUpdate/fehlgeschlagenes Required-Update/Rollback verbrauchen2/4/6 und danach
werden3/5/7 vergeben. Gelöschte höchste/alle Zeilen ergeben danach4. COUNTER(100,7)
ergibt100/107/114. Nach explizitem2147483646 folgen2147483647,-2147483648,-2147483647.
Recordset-Feld Attributes49/DataUpdatableTrue bei bestehendem Datensatz, DAO-Änderung
3164, SQL-Änderung3113, explizites SQL-NULL3162. GUID-GenGUID-Default wurde beobachtet,
beweist aber keinen Replikationsmodus. Form-/Combo-Erzeugung bleibt wegen des
vorherigen nativen CreateControl-Absturzes aus den Paketen entfernt.

Implementiert: Hashgebundene native Jet3/4-Long-Quellzählerstände, Wiederherstellung
von sqlite_sequence/verwalteten Sequenzen einschließlich leerer Tabellen und
höchster gelöschter ID. Quellstand/Werte/Hash-Widersprüche brechen atomar ab.
Native-source UPDATE-Sperre3113 und DAO-Änderungsschutz3164; für bestätigte direkte
writable Long-Generatorfelder Field.Attributes49 und DataUpdatableTrue. Keine
automatische Legacy-Migration oder Diagnosepfad-Auswertung im CSV-Modus.
F171 bleibt offen: alter SQLite64-Bit-/Rollback-Allocator implementiert noch nicht
AddNew-Zeitpunkt, verbrauchte Nummern bei Rollback/Fehler, signed32-Wrap,
Random/Replication, vollständige Schrittweiten-/Property-/Schema-Lifecycle-Regeln.
Native Decoder verweigert negative/UUID/offengelegte Seed-/Increment-/NewValues-
Zustände; fehlende Properties beweisen nicht sämtliche möglichen Modi. Details in
AUTONUMBER_IMPORT.md. Optionale tatsächliche Jet4-Zustandsreferenz erzeugt
JetCounterStateReference.java; externe Jackcess-Dateien nicht committen/ausliefern.

Gebündelte Validierung dieses Arbeitsstands: gezielte bestehende Ziele
(table_import_contracts mit echter Jet-Counter-Zustandsreferenz und
complex_real_mdb_table_copy mit unabhängigen Double-Oracles)2/2 in81,28s.
Zusätzlich dieselbe bestehende29-Tabellen-Prüfung auf der neu hochgeladenen
Kassa-Version erfolgreich; alle4 Fixture- und3 Upload-SHAs unverändert. App und
nativer Helfer neu gebaut. Keine neue48/48-Vollsuite oder grüne GitHub-CI behaupten.
Logs privat: /workspace/work/counter-batch-{build,tests,upload,hashes}.log,
/workspace/work/counter-{app-build,upload-hashes}.log.


Nach e37de78: DAO-DDL umfasst nun Schemaaktion und Bereinigung der Zähler-/Feld-/
Index-Metadaten in einem Savepoint. DROP entfernt _la_counter_sequences,
_la_table_schema und _la_access_indexes für die tatsächlich entfernte Tabelle;
Rollback stellt Tabelle und Zustand wieder her, Wiederimport unter demselben
Namen startet korrekt mit dem Quellstand. Historische Importlogs/Abfragen bleiben.
Unbewiesene Umbenennung mit gebundenen Zähler-/Quellmetadaten wird atomar3251
verweigert; kein stiller kaputter Defaultverweis. Allgemeine Umbenennung und
Relation-Lifecycle bleiben offen. Bestehendes dao_execution1,18s bestanden;
table_import_contracts mit Jet-Zustandsfixture0,13s bestanden. Im neuen Prüffall
war BeginTrans/Rollback zunächst fälschlich am Database- statt Engine-Objekt
aufgerufen; ausschließlich diesen Prüffehler korrigiert und das betroffene Ziel
erneut ausgeführt. Kein erneuter großer Originaldaten-/Vollsuite-Lauf nötig.


## Neuer Counter-Kern nach e70d5d6 — allgemeine Datenbankfunktion

Native neue MDB-Importe verwenden für ALLE AutoNumber-Felder INT plus verwalteten
DEFAULT, einschließlich der16 einzelnen Primärschlüssel; keine INTEGER-rowid-
Aliasabhängigkeit. Das ist keine Kassa-Sonderregel. Neue Designer-Tabellen verwenden
denselben Kern; native DAO-DDL konfiguriert erkannte neue Generatoren ebenfalls.
Persistente _la_counter32_sequences mit signed32-Werten, SQL/DAO-Übergang
2147483647 -> -2147483648 -> -2147483647, Quellenzustand unverändert übernommen.
Legacy-/caller-owned-CSV bleibt64-Bit und wird nicht still migriert. Explizit
SQLite-AUTOINCREMENT-DDL behält den rowid-Alias: DAO kann über den32-Bit-Katalog
reservieren, automatisches RAW-SQL-Wrap dieses Alias bleibt unbewiesen/verweigert.

DAO AddNew reserviert vor Rückgabe, füllt den Puffer als Long und schreibt die
reservierten Felder beim Update explizit. Cancel/Close/fehlgeschlagenes Update
verbrauchen die Nummer. Workspace merkt Transaktionsreservierungen; vor Rollback
wird die aktuelle Frontier gelesen, ROLLBACK TO setzt Geschäftsdaten zurück,
verbrauchte Sequenzwerte werden VOR RELEASE unter derselben Transaktionssperre
persistiert. Kein Sidecar/neuer Dienst. Schemaaktionen nach reservierten Nummern
in derselben DAO-Transaktion werden3251 verweigert, bis die Transaktion beendet ist.
Bei Neu-Puffer-NULL bleibt eine bereits reservierte Nummer erhalten; keine darüber
hinausgehende DAO-NULL-Assignment-Parität behaupten. Counteridentität wird für DAO-
Tabellenreferenzen kanonisiert; ASCII-Groß-/Kleinschreibung ist im bestehenden Test
abgedeckt. Bookmarks/rowids bleiben von sichtbaren AutoWerten getrennt.

Neue importierte verwaltete Zähler verweigern SQL-NULL3162 und SQL-UPDATE3113,
DAO-Änderung3164; gebundene Formularsteuerungen verweigern bestehende Änderungen
2448. Normale Formularupdates lassen unveränderte Generatorfelder aus, sonst würde
die neue SQL-Sperre gültige Updates anderer Felder verhindern.32-Bit-Generatorkatalog
liefert Long-Feld-/UI-Typen; nichtzusammengesetzte Query-Projektionen mit geprüfter
physischer main-Spaltenherkunft behalten Long. Keine Typableitung aus Zelltext,
keine unbewiesene Expression-/UNION-/writable-Query-Lineage.

Gezielte gemeinsame Integration6/6 in101,55s: mainwindow_documents,
sqlite_backend_identity, dao_execution, table_import_contracts (mit tatsächlichem
Jet-Zustandsfixture), access_2003_layouts und complex_real_mdb_table_copy mit
unabhängigem Double-Oracle.29/29 Originaltabellen,20 Primär-/48 andere Indexnamen,
8 Relationen und57 Original-Double-Spalten unverändert geprüft. Anschließend nur
betroffene Fälle: Formular-Generator/gewöhnliches Feld-Update im bestehenden Layout-
Ziel1/1 (11,02s), kanonische Generatoridentität im bestehenden Import-Ziel1/1
(0,23s). App neu gebaut. Kein neuer48/48- oder GitHub-CI-Erfolg behaupten.
Prüflogs privat unter /workspace/work/counter-complete-batch-{build,tests}.log,
counter-form-guard-{build,tests}.log und counter-identity-{build,tests}.log.

F171 weiterhin offen: COUNTER(100,7) ist als echte Access-Referenz vorhanden,
aber vollständige native Seed-/Schrittweiten-Syntax/Quellmodi fehlen noch.
Random/Replikation, negative MDB-Quellstände, komplette Properties, Form-/Event-
Allokationszeitpunkt, RAW-SQL-Statementfehler-Lücken, externe Transaktionen/Crash,
konkurrierende Anwendungen, allgemeine Umbenennung und Relation-Lifecycle bleiben
zu bearbeiten. Vorhandene SQL-/DAO-v4-Setupfehler bleiben KEIN belastbares Ergebnis-
Oracle. Details/autoritative aktuelle Grenzen in docs/AUTONUMBER_IMPORT.md.
Fester Katalog weiterhin456/384/72; keine neue F-ID, kein F171-Häkchen.


## Native SQL COUNTER(seed,increment) nach0b795da

CREATE TABLE unterstützt jetzt COUNTER sowie COUNTER(seed,increment) in echten
Spaltentyp-Positionen über den bestehenden SQL-Tokenstrom, ohne Regex-Dispatch.
Lokaler Tabellenname, eine Counterdeklaration, signed-Long-Literalseed und positive
Long-Literal-Schrittweite; unbekannte/nichtliterale/Nullschritt-Fälle3251 vor
Mutation. INT NOTNULL plus vertrauenswürdiger DEFAULT ohne rowid-Alias; Literaltext
COUNTER(0,0) wird nicht als Deklaration interpretiert. AccessQueryPlan trägt die
geprüften Counterdeklarationen; DAO-DDL initialisiert sie im selben Savepoint.
Optionales main._la_counter32_steps speichert die Schrittweite, ohne das bestehende
32-Bit-Sequenzschema umzubauen. Fehlende Schritttabelle/-zeile bedeutet1, defekte
vorhandene Schrittspeicherung wird nicht als gültig behandelt. Allocation rechnet
mit signed32-Wrap und der gespeicherten positiven Schrittweite. DROP bereinigt
auch Schrittmetadaten. Keine Wiederinitialisierung existierender Tabellen.

Echte Access-Referenz100/107/114 über normalen SQL-/DAO-Neuerstellungsweg umgesetzt.
Bestehende Import-/DAO-Ziele2/2 (1,30s): SQL+DAO-Werte100/107/114, Cancel121 gefolgt
von128, Rollback135 gefolgt142, NoPK5/7, unverändertes Textliteral, Nullschritt-
Abbruch ohne Tabelle, Dropbereinigung und Projekt-Reopen mit weiterhin7er-Schritt.
Keine neue Testziel-/Punktvermehrung. Private Logs counter-step-{build,tests}.log
und counter-step-app-build.log unter /workspace/work/. Binäre MDB-Schrittweiten/
Moduserkennung, negative Source-Seeds, Random/Replikation und vollständige
Property-/ALTER-/Transaktions-/Event-Regeln bleiben F171 offen. Katalog456/384/72
unverändert; kein Abschlussanruf/kein Gesamtabschluss für diese Ergänzung.


Nach cae1a9d: nachgewiesene32-Bit-Generatorfelder im TableDef-Katalog liefern
Attributes17 (Fixed|AutoIncr) und DataUpdatableFalse; Recordset49/True bleibt
getrennt. Die21 Original-TableDef-Counter aus dem tatsächlichen Access-2003-
Metadatenlauf zeigen17/False. Keine Attribut-Ableitung für unbewiesene Query-/
Expression-/Nichtcounter-Kontexte. Katalogzuweisung bleibt383, Indexfeld-
Attributes (Descending1) bleiben getrennt. Vorhandenes Importziel mit tatsächlicher
Jet-Counterreferenz1/1 (0,26s), App gebaut. Private Logs counter-catalog-properties-
{build,tests}.log. Keine Aussage vollständiger Field.Properties-Semantik.
