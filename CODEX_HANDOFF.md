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
