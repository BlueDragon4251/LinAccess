# Access-Kompatibilität — ehrlicher Status und Zielarchitektur

Stand: LinAccess Native 0.7.0. **Nur ein Entwicklungsstand, keine 1:1-Migration.**

**Aktueller Regressionhinweis:** Der erneute Test der hier mitgelieferten
Kassen-MDB auf einer Linux-Umgebung mit MDBTools 1.0.0 schwankte zwischen
0 und 20 von 29 importierten Tabellen; unter anderem treten
MDBTools-Page-Offset- und Schema-/CSV-Spaltenfehler auf. Die in der
Tabelle unten genannte Zahl 29/29 bezeichnet den früheren Teststand und ist
kein garantiert reproduzierbarer Erfolg. Siehe auch `CODEX_HANDOFF.md`.

| Bereich | Jetzt | Geplante Strategie | Grenze / Risiko |
|---|---|---|---|
| `.mdb`-Tabellen | MDBTools: typisierte INTEGER/REAL/TEXT-Kopie inklusive NULL, BLOB-Spalten ohne Nutzdaten; **29/29** Tabellen der echten Kassen-MDB in SQLite-Kopie importiert | Vollständiger Jet-Schemaimport | Primärschlüssel, Indizes, Standardwerte, echte OLE-Daten, Beziehungen, Encoding |
| `.accdb`-Tabellen | Nur soweit installierte MDBTools-Version liest | Zusätzlicher Importadapter | Versions- und Feature-Abhängigkeit |
| Tabellen-/Datensatzeditor | Ja, für SQLite | Constraints, Indizes, Migrationen | Noch rudimentär |
| Access-Abfragen | Exportierte Jet-SQL-Dateien können gegen SQLite ausgeführt werden, wenn die Syntax bereits kompatibel ist; 100 Originalabfragen aus der Kassen-Bridge sind als Quelltext verfügbar | Jet/ACE-SQL-Übersetzung | Jet/ACE-Funktionen, Parameter, Crosstab |
| Access-Formulare | `SaveAsText`-Standardcontrols nun anhand ihrer TWIP-Koordinaten in Qt gezeichnet; 47 echte Kassenformulare mit 685 erkannten Controls geprüft. Einfache direkt gebundene Datensätze les-/schreibbar; andere RecordSources oft nicht kompatibel | Access-Steuerelemente und Layouttreue ausbauen | ActiveX, verschachtelte Unterformulare, Layoutdetails, gebundene Abfragen |
| Access-Berichte | Echter `SaveAsText`-Quelltext erhalten, nicht nativ ausführbar | Report-Modell | Gruppierungen, Seitenformate, Ausdrucksengine |
| VBA-Module aus `.mdb`/`.accdb` | Anwenderexporte mit Access 365 und Access 2003 geprüft. Komplexe Kassen-MDB: 32 VBA-Dateien, 187 Prozeduren statisch erkannt, Umlautnamen unterstützt. Die Kassenprozeduren **nicht** ausgeführt | Umfangreicher VBA-Parser/Interpreter | VBA-Quellcode aus geschützten Projekten, Referenzen, dokumentgebundene Steuerelemente |
| `.bas`/`.cls`-Textdateien | Import in Editor, Originalbytes erhalten, Prozedurerkennung | Vollständiger Parser | Zeichencodierung; nur sehr kleine ausführbare Teilmenge |
| VBA-Ausführung | **Experimentelle Teilmenge**: Variablen, Ausdrücke, If, For, einige Hostmethoden, parameterlose Call-Ketten und `Form_Open(Cancel As Integer/Boolean)`. `DELETE`/`DROP` u. a. werden nach erfolgreicher Interpretation tatsächlich gegen die LinAccess-Projektkopie ausgeführt | Vollständiger Interpreter für VBA-Syntax/Semantik | Nicht die Microsoft-VBA-Runtime, keine allgemeine API-Kompatibilität |
| `Me`, `DoCmd`, `CurrentDb`, DAO | `Me!Feld`, `.Visible`, `.Requery`, `.SetFocus`, `RunSQL`, `CurrentDb.Execute`, `OpenForm`, `Maximize`, `GoToRecord acLast/acNext/acNewRec`, `DoCmd.Requery`; DAO-Recordsets fehlen | Kompatibles Access-Objektmodell | Sehr umfangreiche API-Oberfläche |
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
