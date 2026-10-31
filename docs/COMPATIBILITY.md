# Access-Kompatibilität — ehrlicher Status und Zielarchitektur

Stand: LinAccess Native 0.7.0. **Nur ein Entwicklungsstand, keine 1:1-Migration.**

**Aktueller Regressionhinweis:** Der Gesamtimport übernimmt jetzt alle 29
Kassen-Tabellen mit NULLs, exportierten Schlüsseln/Indizes und Beziehungen atomar.
Der Schlüsselzeilen-Parserfehler mit MDBTools 1.0.1 ist behoben; MDBTools 1.0.0
benötigt zusätzliche Schema-Metadaten aus dessen MySQL-Backend. Die früheren
intermittierenden EOF-Fehler sind hier nicht reproduziert und nicht abschließend
erklärt. Siehe [Untersuchung](MDB_IMPORT_REGRESSION.md) und
[Entwicklungscheckliste](DEVELOPMENT_CHECKLIST.md).

| Bereich | Jetzt | Geplante Strategie | Grenze / Risiko |
|---|---|---|---|
| `.mdb`-Tabellen | MDBTools: typisierte INTEGER/REAL/TEXT-Kopie inklusive NULL, BLOB-Spalten ohne Nutzdaten; **29/29** Tabellen der echten Kassen-MDB in SQLite-Kopie importiert | Vollständiger Jet-Schemaimport | Standardwerte, echte OLE-Daten, exakte Jet-Typen, Encoding |
| `.accdb`-Tabellen | Nur soweit installierte MDBTools-Version liest | Zusätzlicher Importadapter | Versions- und Feature-Abhängigkeit |
| Tabellen-/Datensatzeditor | Ja, für SQLite | Constraints, Indizes, Migrationen | Noch rudimentär |
| Access-Abfragen | Jet-Adapter mit typisierten PARAMETERS/QueryDefs, grundlegenden Strings/LIKE/TOP einschließlich Sortiergleichständen, Datumswerten und eager IIf; Navigator-Ergebnisansicht. 100 Originalabfragen klassifiziert, 58 ausgeführt ohne Access-Referenzergebnis | Jet/ACE-SQL vervollständigen | Funktionen, implizite Formularparameter, Crosstab, genaue Typ-/Ausdruckssemantik; [Details](QUERY_ADAPTER.md) |
| Access-Formulare | 47 Originalformulare mit 878 geparsten Controls / 86 sichtbaren CommandButtons; unterstützte Qt-Widgets aller 47 Formulare instanziiert. Direkte Tabellen editierbar, gespeicherte Abfragen begrenzt schreibgeschützt. Gebundene Lifecycle-/Validierungsereignisse getestet | Access-Steuerelemente und Layouttreue ausbauen | ActiveX, eingebettete Unterformulare, Layoutdetails, vollständige Ereignisse; [Events](FORM_EVENTS.md) |
| Access-Berichte | Original-Layoutvorschau mit einseitigem Qt-Druck/PDF; drei originale Open-/Close-Paare automatisch ausgeführt, tatsächliche CommandBars-Aktionen angepasst; 6/9 Quellen binden 54 nicht-NULL-Textfelder | Report-Modell vervollständigen | Keine vollständige Daten-/Gruppen-/Seitenverarbeitung oder Access-Druckgleichheit; [Druckadapter](PRINT_PREVIEW_ADAPTER.md) |
| VBA-Module aus `.mdb`/`.accdb` | Anwenderexporte mit Access 365 und Access 2003 geprüft. Komplexe Kassen-MDB: 32 VBA-Dateien, 187 Prozeduren statisch erkannt, Umlautnamen unterstützt. Ein Original-Kassenhandler per Qt-Klick/DAO ausgeführt; keine vollständige Laufzeitabdeckung | Umfangreicher VBA-Parser/Interpreter | VBA-Quellcode aus geschützten Projekten, Referenzen, dokumentgebundene Steuerelemente |
| `.bas`/`.cls`-Textdateien | Import in Editor, Originalbytes erhalten, Prozedurerkennung | Vollständiger Parser | Zeichencodierung; nur sehr kleine ausführbare Teilmenge |
| VBA-Ausführung | **Experimentelle Teilmenge**: Variants, Ausdrücke, If, For/For Each/Do While, DAO, parameterisierte Sub/Function-Aufrufe, ByRef-Aliase, Label-Fehlerbehandlung/Resume und gebundene Ereignisse. SQL-Aktionen ändern die Projektkopie tatsächlich | Vollständige VBA-Syntax und Semantik | Klassen, AST/Grammatik, Modulzustand, Arrays, Optional/ParamArray, präzise Date/Currency/Decimal fehlen; [Laufzeit](DAO_RUNTIME.md) |
| `Me`, `DoCmd`, `CurrentDb`, DAO | Native Teilmengen mit tatsächlichen Recordsets/Fields/QueryDefs, typisierten Werten, Navigation und Edit/AddNew/Update/Delete; importierte Forms/Controls-Sammlungen und formulierübergreifende Änderungen mit Lebensdauerprüfung | Access-Objektmodell vervollständigen | ADO, Reports, volle Collections/Default-Member, Filter/Sort, Klassen und Access-Fenstersemantik fehlen; [Objekte](FORM_OBJECT_MODEL.md), [Typen](TYPED_FIELDS.md) |
| COM/ActiveX/WinAPI | **Nein** | Fehlerdiagnose, ggf. eingeschränkte Adapter | Echte Windows-COM-Objekte nicht nativ Linux |
| Access-Makros | Nein | Eigene Aktionsrepräsentation | Kein pauschaler 1:1-Support |

## Kompatibilitätsstufen für jedes importierte Objekt

- **native**: LinAccess-Objekt, in der App erstellt und funktionsfähig
- **imported_exact**: Originalsemantik durch Test nachgewiesen (derzeit *keine*)
- **imported_partial**: Objektdaten vorhanden, jedoch mit ausgewiesenen Verlusten
- **source_only**: Originalquelltext erhalten, nicht ausführbar
- **metadata_only**: Nur Namen/SQL/Tabellenschema erhalten, ohne Formularlayout/VBA
- **unsupported**: Bekanntes Objekt oder API wird nicht unterstützt
- **unknown**: Objektvorkommen lässt sich aus der Quelle noch nicht feststellen
- **failed**: Importversuch fehlgeschlagen, Quelle unverändert

Neue SQLite-Protokollzeilen `partial_typed_nulls` sind ausdrücklich eine
*imported_partial*-Variante; Altprojekte können noch `partial_text_only` enthalten.
Bei Formularen und Berichten können die Namen über den Jet-Katalog nachgewiesen
werden (`metadata_only`), aber VBA-Inhalte und Layout bleiben **unknown**.
Ein Bridge-Export kann Quellartefakte als **source_only**
oder als explizit fehlgeschlagen/nicht verfügbar katalogisieren. Ein separater
**bewusst verlustbehafteter** Konvertierungsschritt erzeugt aus tabellengebundenen
Formularen native LinAccess-Objekte, behält VBA-Quellen und Quellverweise bei und
reduziert das Formular auf tatsächlich vorhandene Felder. Er rekonstruiert
**nicht** das Access-Layout oder Unterformulare. Ein solcher konvertierter
Eintrag ist nicht `imported_exact`.

## Importprinzipien

1. Access-Quelle niemals beschreiben, überschreiben oder löschen.
2. Import nur in ein getrenntes LinAccess-Projekt und zunächst nach Bestätigung.
3. Fehler rollen die SQLite-Importtransaktion zurück.
4. Nicht konvertierte Objekte nie stillschweigend als funktionsfähig anzeigen.
5. VBA- und Form-/Report-Quelldateien in der Bridge mit SHA-256 prüfen und als
   unveränderte Bytes erhalten, auch wenn die Ausführung fehlschlägt.
6. Für 1:1-Ansprüche sind Regressionstests mit repräsentativen echten
   Access-Projekten notwendig. Ohne solche Tests keine Kompatibilitätszusage.

## Warum eine optionale Windows-Export-Brücke sinnvoll sein kann

Die Binärformate von Access enthalten nicht nur relationale Tabellen, sondern
auch UI-Layouts, embedded VBA und Access-spezifische Metadaten. Die vorhandenen
Linux-Tools decken das nicht durchgehend ab. Ein Exporter, der auf einem Rechner
mit installiertem Access dessen Objektmodell verwendet, kann Formulare,
Berichte, Eigenschaften, Abfragen und Module strukturierter exportieren. Danach
entwickelt und betreibt LinAccess diese Artefakte unter Linux **ohne** Windows.
Die erste Bridge-Version wurde auf der Linux-Seite mit einem vom Nutzer per
Microsoft Access 365 exportierten Access-2003-Projekt **real getestet**
([Regression](KEKSREZEPTE_BRIDGE_TEST.md)). Zusätzlich wurde ein reales,
direkt mit **Access 2003 (11.0)** exportiertes Kassenprojekt geprüft: 192
Quellartefakte mit gültigen Hashes, darunter 32 VBA-Dateien und 187 erkannte
Prozeduren ([Kassen-Regression](KASSA_ACCESS2003_BRIDGE_TEST.md)). Die
verwendete `SaveAsText`-Methode ist nicht als stabile Import-API garantiert.
Aus diesen beiden Exporten folgt weder flächendeckende
Versionenkompatibilität noch vollständige VBA-Semantik.

Native Unterformulare ohne Master-/Child-Verknüpfung, typisierte ComboBox-Spalten
und einfache berechnete Felder sind mit Originalartefakten geprüft. Ein komplexer
originaler Form_Load läuft mit 15 unabhängig geprüften dynamischen Tags und sechs
Kind-Datenquellen. Auch Form_Open läuft mit nativer Menüleiste durch.
Form_Current und zwei unabhängig geprüfte DSum-Berechnungen laufen ebenfalls.
Eine CCur-Berechnung und Aktivierungsereignisse bleiben offen; dies ist kein vollständig lauffähiger Kassen-Workflow.
Siehe [Nachweise und Grenzen](NATIVE_FORM_RUNTIME.md).
