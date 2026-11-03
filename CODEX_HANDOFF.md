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
