# Access-Kompatibilität — ehrlicher Status und Zielarchitektur

Stand: LinAccess Native 0.7.0. **Nur ein Entwicklungsstand, keine 1:1-Migration.**

`Eval`/`Application.Eval` haben einen begrenzten nativen Ausführungspfad mit
öffentlichen VBA-Funktionen, Forms-Verweisen und Access-Nulltests. Ein unverändertes
Originalmodul mit 15 Eval-Aufrufen läuft auf vorbereiteten synthetischen Controls.
Das belegt keine vollständige Ausdrucks- oder Projektkompatibilität:
[Eval-Umfang und Grenzen](EVAL_RUNTIME.md).

Native Datensatzauswahl und Löschen unterstützen begrenzte RunCommand-/alte
DoMenuItem-Pfade mit echten Transaktionen und abbrechbaren Löschereignissen.
Ein Original-Löschhandler ist direkt und per gebundenem Klick auf synthetischen
Zeilen geprüft. Copy/PasteAppend und vollständige Access-Ereignisgleichheit bleiben
offen: [Umfang und Grenzen](RECORD_DELETE_RUNTIME.md).

21 originale AutoNumber-Markierungen bleiben erhalten. Neue Access-Importe
verwenden für alle21 Zähler dieselbe verwaltete signed32-Vergabe, einschließlich
16 einzelner Primärschlüssel und fünf Zähler ohne Primärschlüssel. Quellstände,
DAO-AddNew-Reservierung, verbrauchte Cancel-/Rollback-Nummern und der Übergang
ins Negative sind umgesetzt; neue Designertabellen nutzen denselben Kern.
Legacy-/CSV-Projekte werden nicht still migriert. Das ist keine vollständige
Jet-Zählersemantik: [AutoNumber-Grenzen](AUTONUMBER_IMPORT.md).

Begrenztes Copy/PasteAppend verwendet die echte Zwischenablage und erhält
native Feldwerte. Tatsächliche AutoNumber-Felder erzeugen neue Schlüssel; manuelle
Primärschlüssel bleiben erhalten. Der originale Kopierhandler behandelt im Test
seinen Schlüsselfehler, ohne eine Zeile zu duplizieren. Vollständige Insert-Events
und Clipboard-Interoperabilität bleiben offen: [Clipboard-Laufzeit](RECORD_CLIPBOARD_RUNTIME.md).

**Aktueller Regressionhinweis:** Der Gesamtimport übernimmt jetzt alle 29
Kassen-Tabellen mit NULLs, exportierten Schlüsseln/Indizes und Beziehungen atomar.
Der Schlüsselzeilen-Parserfehler mit MDBTools 1.0.1 ist behoben; MDBTools 1.0.0
benötigt zusätzliche Schema-Metadaten aus dessen MySQL-Backend. Die früheren
intermittierenden EOF-Fehler sind hier nicht reproduziert und nicht abschließend
erklärt. Siehe [Untersuchung](MDB_IMPORT_REGRESSION.md) und
[Entwicklungscheckliste](DEVELOPMENT_CHECKLIST.md).

Der GUI-Build verwendet jetzt das offizielle, mitgelieferte SQLite 3.54.1 als
private Bibliothek. Damit bestehen die zuvor beim Vorbereiten verschachtelter
JOIN-Views abstürzenden Ubuntu-24.04-Regressionen. Qt-SQL-Verbindung und native
Adapter müssen dieselbe gepinnte Backend-Identität melden; andernfalls Fehler
3251. Das System-SQLite bleibt unverändert. Dies verbessert die Reproduzierbarkeit,
belegt jedoch keine zusätzliche pauschale Jet-/Access-Kompatibilität. Siehe
[Backend und Aktualisierung](../third_party/sqlite/README.md).

Der native DAO-Suchpfad unterstützt jetzt begrenztes Find/Clone/Bookmark und
Form.Bookmark-Navigation mit Validierung. Zwei unveränderte originale Suchhandler
funktionieren per Qt-Auswahl auf vorbereiteten Testdaten; ein dritter behält wegen
eines fehlenden Controlverweises Fehler 2465. Vollständiges editierbares
Form.Recordset und generelle Array-/Sperrsemantik bleiben offen:
[Cursor-Grenzen und Nachweise](DAO_CURSOR_RUNTIME.md).

| Bereich | Jetzt | Geplante Strategie | Grenze / Risiko |
|---|---|---|---|
| `.mdb`-Tabellen | MDBTools: typisierte INTEGER/REAL/TEXT/BLOB-Kopie inklusive NULL, begrenzte Rohbinärbytes mit explizitem Hexexport; **29/29** Tabellen der echten Kassen-MDB in SQLite-Kopie importiert | Vollständiger Jet-Schemaimport | Vollständige Standardausdrücke, OLE-Anzeige/Anhänge/Originalformat-Roundtrip, exakte Jet-Typen, Encoding |
| `.accdb`-Tabellen | Nur soweit installierte MDBTools-Version liest | Zusätzlicher Importadapter | Versions- und Feature-Abhängigkeit |
| Tabellen-/Datensatzeditor | Ja, für SQLite | Constraints, Indizes, Migrationen | Noch rudimentär |
| Access-Abfragen | Jet-Adapter mit typisierten PARAMETERS/QueryDefs, grundlegenden Strings/LIKE/TOP einschließlich Sortiergleichständen, Datumswerten und eager IIf; Navigator-Ergebnisansicht. 100 Originalabfragen klassifiziert, 77 ausgeführt (davon fünf mit Datenparametern) ohne Access-Referenzergebnis | Jet/ACE-SQL vervollständigen | Funktionen, implizite Formularparameter, Crosstab, genaue Typ-/Ausdruckssemantik; [Details](QUERY_ADAPTER.md) |
| Access-Formulare | 47 Originalformulare mit 879 geparsten Controls / 86 sichtbaren CommandButtons; unterstützte Qt-Widgets aller 47 Formulare instanziiert. Direkte Tabellen editierbar, gespeicherte Abfragen begrenzt schreibgeschützt. Gebundene Lifecycle-/Validierungsereignisse getestet | Access-Steuerelemente und Layouttreue ausbauen | ActiveX, verknüpfte Unterformulare, Layoutdetails, vollständige Ereignisse; [Events](FORM_EVENTS.md) |
| Access-Berichte | Original-Layoutvorschau; zusätzlicher nativer Druck/PDF mit wiederholten Detailbereichen, begrenztem Mehrfeld-CanGrow und zeilensicherem Umbruch langer Details, Kopf-/Fußbereichen, unterstützten Ausdrücken und Seitenzahlen; echte mehrseitige Qt-Vorschau direkt im Bericht mit Navigation und Neuabfrage; drei originale Open-/Close-Paare und drei Aggregate mit SQL-Abgleich; 6/9 Quellen binden 54 nicht-NULL-Textfelder | Report-Modell vervollständigen | Gruppen, vollständiges CanGrow/Seitenteilen, vollständige Ereignisfolge, Access-Seiteneinstellungen und Druckgleichheit bleiben offen; sechs originale komplexe Aggregataufrufe ohne verfügbaren Quellenkontext nicht ausgeführt; [Druckadapter](PRINT_PREVIEW_ADAPTER.md) |
| VBA-Module aus `.mdb`/`.accdb` | Anwenderexporte mit Access 365 und Access 2003 geprüft. Komplexe Kassen-MDB: 32 VBA-Dateien, 187 Prozeduren statisch erkannt, Umlautnamen unterstützt. Ein Original-Kassenhandler per Qt-Klick/DAO ausgeführt; keine vollständige Laufzeitabdeckung | Umfangreicher VBA-Parser/Interpreter | VBA-Quellcode aus geschützten Projekten, Referenzen, dokumentgebundene Steuerelemente |
| `.bas`/`.cls`-Textdateien | Import in Editor, Originalbytes erhalten, Prozedurerkennung | Vollständiger Parser | Zeichencodierung; nur sehr kleine ausführbare Teilmenge |
| VBA-Ausführung | **Experimentelle Teilmenge**: Variants, Ausdrücke, If, For/For Each/Do While, DAO, parameterisierte Sub/Function-Aufrufe, ByRef-Aliase, Label-Fehlerbehandlung/Resume und gebundene Ereignisse. SQL-Aktionen ändern die Projektkopie tatsächlich | Vollständige VBA-Syntax und Semantik | Klassen, AST/Grammatik, Modulzustand, vollständige Array-/Optional-/ParamArray-Semantik, Date/Decimal und vollständige Currency-Koercionen fehlen; [Laufzeit](DAO_RUNTIME.md) |
| `Me`, `DoCmd`, `CurrentDb`, DAO | Native Teilmengen mit tatsächlichen Recordsets/Fields/QueryDefs, typisierten Werten, Navigation und Edit/AddNew/Update/Delete; importierte Forms/Controls-Sammlungen und formulierübergreifende Änderungen mit Lebensdauerprüfung | Access-Objektmodell vervollständigen | ADO, Reports, volle Collections/Default-Member, vollständige Filter/Sortierung, Klassen und Access-Fenstersemantik fehlen; [Objekte](FORM_OBJECT_MODEL.md), [Typen](TYPED_FIELDS.md) |
| COM/ActiveX/WinAPI | Keine generische Ausführung; begrenzter nativer MSCAL-Kalenderadapter mit originalem VBA-/Datumsdialog-Test | Fehlerdiagnose, ggf. eingeschränkte Adapter | Echte Windows-COM-Objekte nicht nativ Linux |
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
CCur läuft mit einem nicht leeren Testwert; beim anfänglichen Null-Wert bleibt
ein expliziter Fehler 94. Native Aktivierung führt auch Form_Activate aus;
Open/Load/Activate/Current sind damit geprüft. Vollständige Ereignissemantik und
Kassen-Workflows bleiben offen; siehe [Aktivierungsgrenzen](NATIVE_ACTIVATION.md).
Siehe [Nachweise und Grenzen](NATIVE_FORM_RUNTIME.md).

Native Filter/FilterOn-Eigenschaften und Filterleiste arbeiten mit echten
Datensätzen. Zwölf gebundene Felder eines Originalformulars werden nach dem
Filtern unabhängig geprüft. Speichern/Abbruch, Rückkehr nach ungültigem Filter
und Requery sind getestet; [Umfang und Grenzen](NATIVE_FILTERS.md).

Native MsgBox execution and the bounded unqualified form-member bridge are covered
by real Qt choices, SQL effects and one unchanged Kassa missing-input handler; see
[MESSAGE_BOX_RUNTIME.md](MESSAGE_BOX_RUNTIME.md). This does not establish complete
login, VBA class scope or Access action-warning compatibility.

Bounded native DoCmd.Close and SQL-backed ComboBox.Dropdown now execute real Qt
actions, including an unchanged original Close button and a second original
missing-input branch. See [CLOSE_COMMAND_RUNTIME.md](CLOSE_COMMAND_RUNTIME.md)
for cancellation/save semantics and remaining command/combo boundaries.

Persistent scalar module state now distinguishes CodeBehind instances, procedure
locals and connection-owned standard modules; the original public declaration
artifact supplies real globals across native form instances. Visibility, recursion,
control/global precedence and reset tests are documented in
[MODULE_STATE_RUNTIME.md](MODULE_STATE_RUNTIME.md). Supported field types are
Boolean/String/Variant in the preamble; complete VBA classes/scopes remain open.

Nach expliziter Aktivierung des im Export deaktivierten Kalender-Steuerelements
führt der originale Kalender-/Datumsdialog beide Funktions-Ereignisbindungen,
Screen-Zugriffe, Projektvariablen, Monats-Handler, Default-Member-Zugriffe und
seinen Error-/Resume-Pfad tatsächlich aus. Die Auswahl beider Felder wirkt auf
eine Abfrage echter MDB-Datensätze. [Testumfang und Grenzen](CALENDAR_RUNTIME.md).


Der begrenzte [OpenForm-Laufzeitpfad](OPEN_FORM_RUNTIME.md) unterstützt vier
Positionsargumente mit ausgelassenen Vorgaben, acNormal und WhereCondition. Ein
unveränderter Kassen-Handler öffnet sein originales Zielformular; ein zunächst
roter Test belegte die dort fehlende OptionGroup-Laufzeitunterstützung. Numerische
Gruppenwerte, Standardwerte und native exklusive Auswahl sind jetzt implementiert.
Weitere Ansichten, benannte Filter, DataMode/WindowMode/OpenArgs und allgemeine
Access-Fenstersemantik sind nicht vollständig implementiert. OpenForm lässt das
native Ziel nun geöffnet und setzt VBA fort; eine eindeutige bestehende Instanz
wird wiederverwendet. Modal sperrt die anderen Fenster, ohne VBA anzuhalten.
Die Fensterlebensdauer ist vom Aufrufer unabhängig; projektweites Schließen und Quit prüfen native Unload-Abbrüche ([Nachweise](DOCUMENT_LIFETIME_RUNTIME.md)). 187/187 statisch angenommene Prozeduren bedeuten ausdrücklich
keine entsprechende Ausführungsabdeckung.


[Eckige Klammern und UTF-8-Bezeichner](BRACKET_IDENTIFIER_RUNTIME.md) haben
einen begrenzten nativen Laufzeitpfad mit echten Formular-/DAO-Tests. Ein
originaler berechneter Formularwert bleibt schreibgeschützt. Vollständiges Access
Eval, allgemeine Unicode-Namensnormalisierung und das vollständige Ausführen der
umgebenden originalen Kassenprozeduren werden dadurch nicht behauptet.


Die begrenzte [TableDefs-Sammlung](TABLEDEFS_RUNTIME.md) liefert echte
Projekttabellen und Recordsets. Die unveränderte originale Existenzfunktion führt
29 positive Prüfungen und einen negativen Fall aus; alle Recordset-Zeilenzahlen
werden mit unabhängigem SQL verglichen. Vollständige DAO-Katalogeigenschaften,
Fields/Indexes/Relations und allgemeine deklarierte Rückgabetyp-Coercion bleiben
offen. Die 187/187 angenommenen Prozeduren sind weiterhin eine Syntaxmetrik.


[With, native Tastatureingaben und Tabellenerstellungsabfragen](WITH_KEYBOARD_RUNTIME.md)
haben echte, begrenzte Laufzeitpfade. Drei unveränderte Kassen-Tastaturhandler
enden erfolgreich; ihre tatsächlichen Empfänger werden geprüft. Zwei spätere
AfterUpdate-Ereignisse scheitern im Test ohne aktiven Verkauf weiterhin mit
3075/Zeile 1023. Dies ist keine erfolgreich getestete Verkaufskette.
Option Compare ist pro Modul berücksichtigt, die vollständige Jet-Kollation
bleibt offen. SELECT INTO kopiert echte Tabellen; vollständige Jet-Metadaten
werden dabei noch nicht übernommen.


Der native OpenForm-Fortsetzungstest belegt unmittelbare Forms-Zugriffe, echte
SQL-Schreibvorgänge, Instanzwiederverwendung und sofortiges Screen.ActiveForm /
ActiveControl. Abbrechbares Schließen und Selbstschließen mit anschließender
MsgBox sind getestet; die physische Löschung wartet auf das Ende der laufenden
VBA-Aufrufe. Eine unabhängige Access-Dokumentverwaltung und vollständige
Aktivierungs-/DoCmd-/WindowMode-Semantik bleiben offen.


[Globale Befehle und Tastatureingaben](ACTIVE_DOCUMENT_COMMANDS.md) treffen nun
das aktive native Dokument desselben Projekts. Fokus, Speichern, Next/Last/New,
Maximieren, Requery und unterstützte Application.DoCmd-Pfade haben echte
Widget-/SQL-Regressionen; Me bleibt an seine CodeBehind-Instanz gebunden.
Ein Originalformular mit tatsächlichen MDB-Datensätzen wird für Fokus und
Last/Requery unabhängig verglichen; sein CodeBehind ist in diesem Bindungstest
bewusst ausgeblendet. Das ist keine vollständig bestandene Verkaufskette.
SetWarnings, vollständige Befehls-/Unterformularsemantik und Access-Dialogparität
bleiben offen.


Unabhängige Dokumentlebensdauer und der unveränderte originale Kassenöffner sind
in [Dokumentlebensdauer](DOCUMENT_LIFETIME_RUNTIME.md) beschrieben. Dynamische
ComboBox-Eigenschaften verwenden echte mehrspaltige Qt-Modelle, getrennte
gebundene Werte und ausgeblendete Spalten; [Umfang und Grenzen](DYNAMIC_COMBO_RUNTIME.md).
Die vollständige ursprüngliche Verkaufskette bleibt unbestätigt.


Native [keyboard commands and VBA key events](KEYBOARD_RUNTIME.md) have bounded
Qt execution paths, record persistence/undo and cancellation tests. This is not
full Access 2003 shortcut, layout, designer or event parity. An original login
column-index mismatch and two missing exported key-handler bindings are tracked
explicitly; no original credentials or source are copied into public evidence.

### Bounded joined action updates

Native DAO Execute/QueryDef.Execute supports one INNER JOIN, or LEFT/RIGHT JOIN
with the target on its preserved side, and a single direct physical-table target,
qualified assignments, aliases, WHERE and parameters. A
materialized candidate set freezes self-join inputs. Multiple matches are accepted
only when every assigned SQLite value has identical type and value; conflicting
candidates raise 3251 and roll back the statement. Nullable outer-join targets, FULL joins, saved-query targets,
multiple targets and tables without an accessible physical row identity remain
unsupported (3251). This conservative boundary is not Jet updateability parity.

The original INNER JOIN action is exercised unchanged against its actual table
schemas with synthetic, unique matching input and changes one expected value.
The full original-data copy has conflicting candidates and is rejected; the other
original action requires an outer join and saved-query target. Corpus coverage
therefore remains 77/100, without an Access reference oracle.

### Active-document record navigation

DoCmd.GoToRecord supports acFirst/acLast/acNext/acPrevious/acGoTo/acNewRec
for the active native form, including positive relative offsets and absolute
saved-record numbers. acActiveDataObject is accepted explicitly. Both direct VBA
statements and Application.DoCmd dispatch address the active document while Me
continues to identify the caller. Saving/validation precedes movement; cancellation
returns 2501, invalid/out-of-range positions 2105, and disabled additions 2046.
Next from the last saved row enters the new row; Previous from that row returns
to the last saved record. Named table/query targets remain unsupported (3251);
this is not full Access navigation or datasheet parity.

Named GoToRecord with acDataForm now resolves an already opened same-project
form by case-insensitive name, including hidden windows. It does not activate or
reopen the destination. Missing/closed/foreign forms raise 2450; ambiguous form
instances raise 5. Explicitly named navigation preserves the caller Me and the
active form. Table/query/report targets remain unsupported.

### Named control requery

DoCmd.Requery ControlName requeries the supported named control on the active
native form. Combo/list models reload SQL/table/saved-query row sources; existing
subform and calculated-control adapters retain their bounded behavior. A bound
non-list data control reloads the form record source after save validation, with
2501 on cancellation. Me.control.Requery continues to address the caller instance.
Native tests compare actual SQL changes and distinct caller/active combo models.
Focus-dependent implicit subform behavior and complete Access event timing still
need a reference oracle.

### Native insertion events

Direct-table form control changes now invoke bounded BeforeInsert/Dirty and,
after a successful SQL INSERT, AfterUpdate/AfterInsert. Programmatic VBA field
assignments do not invoke insertion events. Cancel, errors, Undo, initialization
and validation/constraint retries have native Qt/SQL regressions. Precise
BeforeInsert/Change/Text/Value visibility is still open: the current adapter
runs after the native editor value changes. INSERT_EVENT_RUNTIME.md records
this boundary and the still-unsupported PasteAppend event pipeline.

### Native control Change events

Imported OnChange bindings on text fields and combo boxes now dispatch through
the configured EventProcPrefix after native edits, including real SQL effects
from handlers. Bound mirror updates, load/requery and VBA Value assignment do
not dispatch Change. A native key test verifies KeyDown/KeyPress/BeforeInsert/
Change/KeyUp order, suppressed KeyPress, existing-record changes and a combo edit.
Handler errors remain visible without saving rejected dirty data automatically.
This does not implement Text-property access, exact Text/Value buffering,
cascading Text setters, LimitToList/NotInList or complete input/locale parity.

Native SelStart/SelLength/SelText now address focused text/ editable-combo
editors, including replacement, Null on no selection, focus/range validation
and UTF-16 selection. Actual Change callbacks and SQL persistence are tested.
FORM_OBJECT_MODEL.md records clamping and remaining text-buffer/format limits;
original selection handlers are not claimed (none occur in the VBA fixtures).

### Form timers

Native form TimerInterval/OnTimer now run actual VBA or supported expression
handlers on a form-owned Qt timer. Load configuration, hidden forms, dynamic
stop/restart, Unload cancellation, errors and self-close have asynchronous SQL
regressions. One unchanged original clock handler updates a native form caption and field
with independently checked timestamps on a lifecycle-isolated test copy.
TIMER_RUNTIME.md records scheduling guards, Report rejection and the absence of
an Access clock-locale/event oracle. This is not full timer/message parity.

### VBA report preview opening

DoCmd/Application.DoCmd.OpenReport supports explicit native preview (2),
WhereCondition, imported report CodeBehind, cancellation/error propagation and
independent document lifetime. An unchanged original preview handler opens its
original report through invocation and a prepared bound-button click; actual
fields are compared with a separate SQL reader. Default direct printing, other
views, named filters, window modes and OpenArgs remain unsupported (3251).
OPEN_REPORT_RUNTIME.md records renderer/pagination/oracle limitations.

### UTF-16 string positions

Len/Left/Right/Mid and the dollar substring variants now use UTF-16 units for
valid Unicode text. Native SQL Len shares the validated counter, including
embedded NUL and pure alias expansion. Tests exercise actual VBA execution,
Null/error/Long bounds and UTF-16 fields. A slice producing an isolated surrogate
is still unsupported (3251); full BSTR storage, declared-type Len sizes and
locale/code-page/string-intrinsic parity remain open (UTF16_STRINGS.md).
Leading-dot/E-exponent numeric literals and numeric Null error 94 are also tested.

Trim/LTrim/RTrim and their String variants now implement end-space removal,
Variant Null propagation and String-result error 94, while preserving other
whitespace. Actual VBA procedure/error-handler regressions cover the family;
this does not extend the remaining code-page/locale/string-type scope.

InStr/StrComp now support UTF-16 positions/binary ordering, Null propagation,
Option Compare and explicit comparison constants, with actual VBA regressions.
Text/database comparisons retain host collation limits: Qt case folding in the
GUI, ASCII folding in the standalone host, no complete Windows/Jet linguistic
matching or LCID support. See UTF16_STRINGS.md.

Native single-line text/editable combo controls now distinguish focused Text
from their prior Value until control update. Validation, cancellation/retry,
mirrors/calculations, programmatic assignment, cascading Change and actual SQL
save are tested. TEXT_VALUE_RUNTIME.md records remaining input/focus/reentrancy,
OldValue, rich-text and combo-validation limits; no full Access timing claim.

Bound data controls now expose OldValue through form AfterUpdate, using the
existing supported field representations. Native text/combo control Undo and
BeforeUpdate self-change errors are tested with real SQL and other-edit
preservation. Reentrant new drafts are retained for explicit retry, not silently
discarded. General Undo events, reentrancy, unbound OldValue and exact
Currency/Decimal/Date field representation remain open (TEXT_VALUE_RUNTIME.md).

Native project-scoped DBEngine/default Workspace transactions now execute real
SQL savepoints, commit and single-level rollback, including recordset writes,
VBA helpers and a real Qt button with handled SQL-error rollback. CurrentDb.Execute
accepts dbFailOnError through its Database object. DAO_TRANSACTIONS_RUNTIME.md
records nested rollback, clone, forced-flush, multi-database, Close and Jet cursor
limits; full DAO transaction parity is not claimed.

SaveAsText Boolean NotDefault now respects each supported property's default.
Enabled/Visible survive record reloads; source read-only protection is separate
from Enabled. Native TextBox Locked, TabStop and bounded flat-section TabIndex
reordering/traversal are tested, including a hash-verified original layout. The
original metadata declares 81 visible command buttons rather than the previously
reported 86; hidden controls remain available to VBA. CONTROL_STATE_RUNTIME.md
records focus/cycle/nesting, other Locked kinds and visual-reference limits.

Native TextBox InputMask=Password now displays immediate asterisks while Value/
Text and real SQL retain the entered value. Dynamic changes, Null/Unicode,
record navigation and one original layout with synthetic input are tested.
Other imported patterns are explicitly unsupported/read-only; the original
login failure and general input-mask semantics remain open (PASSWORD_INPUT_MASK.md).

Flat standalone forms now implement Cycle 0/1 through actual native Tab and
Shift-Tab: validated control transfer, same-row wrap or real record save/move,
cancellation with retained focus, new-row handling and handler focus override.
A close-plus-MsgBox lifetime regression protects this new entry path. Page and
nested/subform cycles plus general focus/event parity remain open
(FORM_CYCLE_RUNTIME.md).

Multiline Select Case now executes first-match list/range/relational branches,
with one-time selector evaluation, nested If/For/Exit For and bounded Resume
handling. Core and native transaction-button tests exercise actual branches and
SQL commit/rollback. Original modules contain no such blocks; no new original
execution count or complete VBA grammar claim follows (SELECT_CASE_RUNTIME.md).

Do Until, postcondition Loop While/Until, unconditional Do/Loop, While/Wend and
Exit Do now use native condition evaluation and explicit nested-frame unwinding.
Core tests exercise control flow/error retry/budget limits; a real Qt transaction
button performs its SQL updates through Do Until. No new original handler count
is claimed; full grammar/error/type parity remains open (DO_LOOP_RUNTIME.md).

Imported text-control font families, sizes, weights, italic and underline now
reach real Qt widgets, including declared control-type templates and per-instance
overrides. Editable combo editors and requeried popup views receive the same font.
SQL/mirroring/navigation and one original form/report layout are tested without
original VBA/data. Missing-font resolution, complete font semantics and exact Windows
appearance remain open (CONTROL_FONT_RENDERING.md).

Native form text-control FontName/Size/Weight/Bold/Italic/Underline now read/write real
instance fonts through VBA. Pending masked Text, old Value, selection, focus,
read-only SQL policy and data Dirty state remain intact; invalid input is checked
before mutation. Combo editors/popups and record navigation retain the changes.
Qt fallback/defaults, full property semantics and exact Windows metrics remain
open (CONTROL_FONT_RENDERING.md).

FontBold follows the documented >=700 threshold, independently of Qt's 600 bold
classification. Unexported sizes use documented 8/10-point defaults. Report font
setters stop with 3251 because the required OnPrint context is not implemented;
static rendering and getters remain available. These contracts are independently
checked against Microsoft documentation, not an Access 2003 runtime oracle.

Multiline ElseIf chains now execute in first-match order with skipped later/
inactive conditions and Null-as-false. Native Qt/SQL tests assert displayed branch
selection and independent commit/rollback balances. Full conditional grammar,
TypeOf and Access error-oracle parity remain open (ELSEIF_RUNTIME.md).

Native Form.AllowEdits now gates bound existing-record user input, preserves
unbound/new-row input and honors the documented programmatic dirty-record
exception until save/Undo. Deletion/policy-change preflight, checkbox Click and
group-only Click plus close/MsgBox lifetime are tested with real Qt/VBA/SQL.
Five original flags are retained; their native query sources are already
read-only, so no original writable-query transition is claimed (ALLOW_EDITS_RUNTIME.md).

Form-control DefaultValue now preserves templates and exposes bounded native
getters/setters. Literal defaults initialize unbound controls and writable-table
new rows without dirtying/inserting; shared fields, empty-versus-NULL SQL storage,
future metadata changes, failed-default recovery and ID nonconsumption are tested.
Original probes retain 34 entries and initialize 15 unbound literals. Computed/
Jet table defaults and full new-record typing/timing remain open (CONTROL_DEFAULT_RUNTIME.md).

Native TextBox/Label TextAlign now preserves imported templates and explicit
General/Left/Center/Right settings, with instance VBA setters and typed General
alignment. Native Qt/SQL tests preserve input buffers and read-only permissions;
24 original layouts expose 208 aligned Labels and 226 TextBoxes without supplied
CodeBehind or rows. Distributed alignment and runtime report setters explicitly
reject as unsupported. This is bounded horizontal rendering, not complete visual
parity; see [Text alignment](TEXT_ALIGNMENT_RUNTIME.md).

Native form-control ForeColor now imports/inherits signed OLE metadata, renders
RGB plus window/button-text system-color intents in Qt, and supports VBA RGB(...)
assignments without saving data. Real masked input, SQL, popup requery, read-only
permissions and native rendered pixels are tested. An empty-DB/no-CodeBehind
probe covers 45 original layouts (271 RGB / 277 system-color controls); no extra
original handler execution follows. Other OLE encodings and runtime report
setters reject explicitly; Windows palette/default/full visual identity remains
open. See [Foreground colors](FOREGROUND_COLOR_RUNTIME.md).

Native TextBox/Label BackColor and BackStyle now retain templates/overrides and
render Normal/Transparent interiors through actual palettes and label auto-fill.
Native composite-pixel/VBA/SQL/masked-input/permission regressions and empty-DB/
no-CodeBehind probes cover 55 original layouts (326 Normal / 587 Transparent /
414 retained colors). Original business coverage is unchanged. Report runtime
setters, other control/section backgrounds, disabled-state and Windows rendering
parity remain open; see [Control backgrounds](CONTROL_BACKGROUND_RUNTIME.md).

RGB(...) now also executes through the native SQL/DAO path, sharing the VBA
primitive. Quoted names, parameters and alias/view reuse are tested alongside
real DAO reads/writes and atomic direct/joined update failures. NULL/type/arity
errors propagate explicitly, including SQL error 450. Original RGB call inventory
is zero in 100 query artifacts and 32 VBA modules; original business coverage is
unchanged. This does not establish complete Jet function/coercion parity; see
[SQL RGB](SQL_RGB_RUNTIME.md).

QBColor(...) now executes the fixed classic 16-color palette in shared native
VBA and SQL/DAO paths. All palette indices, parameters/aliases, error transport,
actual Qt ForeColor/Eval effects and atomic updates through a compiled palette
view are tested. Original 100-query / 32-module QBColor call inventories are zero,
so original business coverage is unchanged. This establishes no full Variant/
Windows color parity; see [QBColor](QBCOLOR_RUNTIME.md).

Imported form/report opening from the main-window source inspector now uses
managed independent windows, not a nested exec loop. Native modal metadata,
same-instance reuse, cancelled Open/Unload, real SQL save and close/MsgBox
lifetime are tested through the production main-window/file-dialog/explorer
route on temporary projects. The legacy generated-form editor remains modal
and full F11/MDI/designer parity stays open; see
[Main-window documents](MAINWINDOW_DOCUMENT_RUNTIME.md).

Classic F11 now activates the project explorer and restores collapsed width,
after native KeyPreview/control validation, while modal forms retain focus.
An explicit database-window context prevents stale-form record commands and
Screen references; default database-window Close/SendKeys reject as unsupported
rather than editing the last form. Native SQL/focus/error/early-Current checks
pass through the actual main-window integration. Full MDI/keyboard/WM/error
identity remains open; see [Database window](DATABASE_WINDOW_RUNTIME.md).

DoCmd.SelectObject now selects real table/query/form/report explorer items or
already-open native Form/Report windows, including hidden restoration and
optional-name group selection. Imported source objects use their corresponding
kind groups; identities use stored metadata. Actual GUI/VBA/SQL tests prove
validation cancellation, unsaved-record preservation, modal/error boundaries
and explicit duplicate-name rejection. Open table/query window activation and
macro/module targets remain unsupported; see [SelectObject](SELECT_OBJECT_RUNTIME.md).

Ctrl+O/Ctrl+N now invoke actual LinAccess project open/create dialogs from the
main window and nonmodal native documents, preserving KeyPreview/control/modal/
Unload cancellation and refusing existing-file overwrite. Switching projects
from executing native SendKeys is blocked. Direct MDB/ACCDB project editing
and complete File-menu/keyboard appearance remain open; see
[Project shortcuts](PROJECT_SHORTCUTS_RUNTIME.md).

Ctrl+F6/Ctrl+Shift+F6 now cycle the actual database window and visible native
forms/reports in a bounded creation-order ring, restoring focus/minimized
windows and honoring KeyPreview, validation and modal blocking. Foreign,
hidden and closed windows are skipped; a target closed during validation is
not activated. SQL and dirty-state preservation are tested. Full Windows MDI,
exact Access window ordering and table/query/editor cycling remain open; see
[Window cycling](WINDOW_CYCLE_RUNTIME.md).

F9 now reloads the focused native combo adapter's actual row model and retains
form recalculation. Native KeyPreview can cancel F9/F4; tests cover F4 opening,
external SQL changes, preserved pending text/dirty/focus/bound IDs, failed refresh
and retry after source recovery. No record save is implied. The combo-backed
ListBox adapter and full popup/Windows timing remain limited; see
[Lookup keyboard](LOOKUP_KEYBOARD_RUNTIME.md).

Native ComboBox.ListIndex and ItemData now expose the actual selected row and
typed bound-column row data, including BoundColumn=0. Tests cover lazy loading,
read-only/index/scalar errors, report/draft restrictions, row removal and requery
without SQL changes. Typed NULL widget assignments now clear selection rather
than accidentally selecting a NULL-key row; non-null empty text remains distinct.
ListBox/multi-selection and full Windows timing/type/error parity remain open;
see [Combo list properties](COMBO_LIST_PROPERTIES_RUNTIME.md).

Single-selection ListBox now renders as a persistent native QTreeView with real
SQL/Value/Field List models, hidden columns and typed bound values. Actual
mouse/keyboard/VBA/SQL regressions cover cancellation, saving/Undo, Locked,
report read-only behavior, defaults and failed-source recovery. MultiSelect
metadata is retained; modes 1/2 show 3251 instead of a single-selection substitute.
Full headers/multi-selection/ListBox API/visual/event parity remains open; see
[Native ListBox](NATIVE_LISTBOX_RUNTIME.md).

ColumnHeads now renders actual native ComboBox/ListBox header views using SQL
projection names or the first Value List row. Header-aware ListCount follows
the documented query/value-list rule; headers remain separate from selectable
data. GUI/VBA/SQL tests cover toggling, template overrides, header clicks, empty
sources, scalar/errors and read-only report context. Field captions/Field List
headers and exact Windows index/visual/error behavior remain open; see
[Column headings](COLUMN_HEADS_RUNTIME.md).

Native ListBox now supports bounded Simple/Extended selection with Selected,
ItemsSelected, ItemData and ListIndex getters over actual Qt selection state.
Indexed VBA assignment uses an explicit setter protocol, including With and
nested expressions. Multi Value stays Null without dirtying/inserting records;
real SQL tests prove other/same-field edits preserve selected state and stored
values. Extended requery clears selection; selected Simple requery rejects while
preserving state. Full Windows event/index/type behavior and the ListIndex writer
remain open; see [Multiple selection](MULTISELECT_LISTBOX_RUNTIME.md).

ComboBox.ListRows now uses the documented default/range and controls the actual
Qt popup data-row cap. Native tests measure 3/5-row viewports, scroll/select/save
a later SQL row and preserve pending text/model identity on property changes.
Report mutation and invalid conversions reject explicitly. Full Windows heading/
theme/screen-edge/open-popup timing remains open; see
[Popup rows](COMBO_LISTROWS_RUNTIME.md).

Module-level Integer/Long fields now have numeric zero defaults and checked
16/32-bit rounded assignments. Core/native GUI/SQL tests verify failed-write
preservation, visibility, per-form/private and shared project state, reset and
actual SQL execution. This establishes storage/assignment scope, with full
numeric subtype/intermediate arithmetic/ByRef/array/class behavior still open;
see [Numeric module fields](NUMERIC_MODULE_FIELDS_RUNTIME.md).

Procedure-local Integer/Long assignments, parameters, results and integral For
counters now use checked widths. Core/native GUI/SQL tests cover conversion,
write preservation, ByVal input, known ByRef conflicts and temporaries, copyback
and pre-effect invalid targets. Fractional typed For bounds reject explicitly;
full subtype/promotion/Variant-reference/loop/compile timing remains open. See
[Numeric procedure values](NUMERIC_PROCEDURE_VALUES_RUNTIME.md).

Boolean/String procedure storage now converts values instead of retaining
incompatible numeric/text kinds. Core/native DAO tests cover typed returns,
ByVal/reference behavior, retained values after invalid writes, zero SQL effects
for invalid arguments, false=0/true=-1 and Unicode text parameter persistence.
Full Variant/object-default/locale/compiler/reference semantics remain open; see
[Scalar procedure values](SCALAR_PROCEDURE_VALUES_RUNTIME.md).

Declared Boolean SQL parameter aliases and strict pre-execution Boolean binding
are documented in SQL_BOOLEAN_PARAMETERS.md. This is bounded SQL/DAO support,
not full DAO type metadata or Variant coercion parity.

Native VBA CBool returns Boolean values with checked scalar conversion, including
shared whitespace handling for numeric Boolean text. See CBOOL_RUNTIME.md for
actual DAO persistence tests and remaining Date/locale/SQL/subtype limitations.

Checked Integer/Long storage and explicit CInt/CLng now retain subtype identity
for VarType/TypeName, including Variant copies. Standard VbVarType constants are
available. See INTEGER_VALUE_IDENTITY.md; numeric literals, arithmetic promotion,
DAO field subtype inference and storage-aware Len remain incomplete.

Checked Single/Double module/procedure values and native CSng now preserve float
precision/type identity and reject overflow before effects. Actual SQL persistence
regressions are documented in FLOATING_VALUES_RUNTIME.md. Single underflow/For,
full numeric formatting, promotion, literals and Variant semantics remain open.

Declared Currency VBA storage now retains exact four-decimal values and native
conversion/reference checks. DAO can bind Currency as canonical decimal text to
TEXT parameters, with a real SQL upper-bound regression. See CURRENCY_STORAGE_RUNTIME.md;
full Jet Currency columns/SQL arithmetic, mixed promotion and loops remain open.

SQL INTEGER parameters now use the documented 32-bit Long range, independently
of 16-bit VBA Integer. Numeric parameter aliases, Single rounding/finite checks
and declared DAO Parameter.Type/read-only/lifetime behavior have real SQL/VBA
regressions. See DAO_PARAMETER_TYPES.md; inference, Size and full coercion remain open.

Native VBA/Access expressions now execute integer division and rounded Mod with
correct bounded precedence/sign/range behavior. A fixture-independent Qt form
regression verifies calculated controls, SQL effects and save/error behavior.
See INTEGER_OPERATORS_RUNTIME.md; SQL rewriting and full numeric promotion remain open.

Native VBA/Access expressions execute caret exponentiation with documented
left-to-right chaining, sign precedence and checked Double/domain behavior.
Actual Qt/VBA/SQL audit tests are in POWER_OPERATOR_RUNTIME.md. SQL caret,
zero-negative powers, complete underflow/rounding/type/error identity remain open.

Native Xor/Eqv/Imp and their logical precedence now execute in VBA and Access
expressions, with truth-table and actual Qt/VBA/SQL audit regressions. See
LOGICAL_OPERATORS_RUNTIME.md; SQL rewriting and full mixed-type/Variant/And-Or
edge semantics remain open.

Native Fix/Sgn and corrected Abs/Int now cover documented null behavior, known
subtypes and exact Currency truncation with real Qt/VBA/SQL regressions. See
NUMERIC_FUNCTIONS_RUNTIME.md; full type/locale/promotion and SQL rewriting remain open.

Native Ctrl+apostrophe copies the previous displayed record field into writable
TextBox edit buffers with actual Qt/VBA/SQL and SendKeys regressions. See
COPY_PREVIOUS_FIELD_RUNTIME.md; exact Access selection/insertion modes, other
adapters and complete keyboard parity remain open.

Native Ctrl+Plus/Ctrl+Shift+Equals and RunCommand new/next/previous/first/last now
use bounded active-document navigation, with real Qt/VBA/SQL policy/validation/
no-INSERT regressions. See NEW_RECORD_SHORTCUT_RUNTIME.md; full keyboard/layout/
MDI/updatable-query parity remains open.

Native Ctrl+Minus/SendKeys use the existing bounded form Delete Record pipeline,
with fixture-independent real Qt No/Yes/default-No, callback/validation and
foreign-key regressions. See DELETE_KEY_RUNTIME.md; full datasheet/batch/query
and exact Windows keyboard/event behavior remain open.

Long native TextBox values no longer truncate at Qt's default editor limit;
real SQL regressions preserve Unicode/newlines through unrelated-field saves,
VBA append and keyboard copy. See LONG_TEXT_VALUES_RUNTIME.md; complete Memo
multiline UI, formatting and Access field-size limits remain open.

Ctrl+F/acCmdFind now opens native literal text search with actual record navigation,
case/match/field modes and Qt/VBA/SQL validation/draft/lifetime regressions. See
FIND_FORM_RUNTIME.md; full Find/Replace, wildcard/format/query/report/datasheet/
locale/MDI and exact Access 2003 UI parity remain open.

Shift+F4/acCmdFindNext (341)/SendKeys repeat the active form's last executed
literal search after the modeless dialog closes, with the same validation and
real SQL navigation. Search history is document-local; shared Access search
history, exact callback/empty-history identity and full Find/Replace remain open.

VBA Byte/CByte now has checked 0..255 conversion and explicit VarType 17 identity
through native modules/locals/parameters/results, with real Qt/DAO/SQL regression.
Known ByRef type mismatch and overflow preserve data; declared Byte For counters
reject before effects. Full promotion/arrays/locales/Windows identity remain open.
See BYTE_VALUES_RUNTIME.md; original business execution coverage is unchanged.

Native VBA Not now preserves Byte/Integer/Long bit width and result subtype,
Boolean/Null/Empty and scalar-once behavior, with before/after core and actual
Qt/VBA/SQL regression. Full promotion/textual Boolean/locale/Jet SQL and Windows
oracle remain open; see TYPED_NOT_RUNTIME.md.

Native VarType now returns an Integer value, including nested VarType/TypeName,
Variant-copy/calculated-control and SQL regression. This fixes the previous
Double result subtype; general object default-property/type inference and full
Variant/array/Date/Decimal semantics remain open. See VARTYPE_RESULT_RUNTIME.md.

VarType now follows explicitly advertised native scalar defaults for control
Value, DAO Field/Parameter and Err, with scalar-once/recursive/error handling
and actual Qt/VBA/SQL type-code regression. Err.Number retains Long; existing
closed recordset data guards now use 3420. Full COM/class/non-Value defaults,
TypeName class metadata, imported field subtypes and DAO identity remain open.
See DEFAULT_TYPE_INSPECTION_RUNTIME.md.

Native Asc/Chr execute ASCII0..127 and AscW/ChrW supported UTF-16 cases, with
checked types/errors, actual Qt/VBA/DAO/SQL and pending/save regression. Three
unchanged original KeyPress procedures execute over 384 direct/285 Qt input cases
on prepared native controls, with original byte hashes verified. This is wiring/
execution evidence, not Windows or full-form/business parity. ANSI/DBCS, byte
strings, dollar aliases, full BSTR and SQL rewriting remain open. See CHARACTER_FUNCTIONS_RUNTIME.md.

Unambiguous imported TINYINT/SMALLINT/INT metadata now restores native Byte/
Integer/Long field values and bounded readonly DAO Field.Type. Real Qt/VBA/SQL
regression and 127 original INT columns/471904 non-NULL values are checked against
SQL data. Generic SQLite INTEGER remains generic; raw SQL range enforcement,
query types, full Field.Type/errors and other exact Jet types remain open.
MySQL FLOAT conflates Currency/Single and is not blindly reclassified.
See INTEGER_FIELD_TYPES_RUNTIME.md.

Round now executes shared native bankers rounding in VBA and the Access SQL
adapter, replacing SQLite half-away delegation. Core/actual Qt/VBA/DAO/SQL
regression covers half cases, explicit fixed Currency, pending/save and failed
UPDATE atomicity. Precision0..15 is bounded; Single/high precision/full binary/
locale/type/Jet money/Windows behavior remain open. See ROUND_RUNTIME.md.

Explicit native Form.Visible can show/hide already activated steady-state forms,
retains hidden document lifetime and drafts on show, and clears stale hidden
active context. Default record/Close and SendKeys cannot target an all-hidden
caller; named close/Unload cancellation remain active. Real Qt/VBA/SQL and window/
original regressions pass. Init/modal/dirty-hide/embedded/report/MDI/full Windows
behavior remains open; see FORM_VISIBILITY_RUNTIME.md.

Native Ctrl+Page Down/Up now use next/previous record commands, preserving actual
text/ComboBox focus and existing validation/save/new-row/lifetime guards. Real
Qt/VBA/SQL tests include KeyDown cancellation, SendKeys and self-close in Current.
Plain Page keys/multi-page scrolling, other views and complete Windows keyboard
identity remain open. See [RECORD_NAVIGATION_KEYS_RUNTIME.md](RECORD_NAVIGATION_KEYS_RUNTIME.md).

Native VBA/calculated-control Replace supports checked start/count/compare,
UTF-16 suffix semantics and scalar-once inputs. Actual Qt/VBA/DAO parameter/SQL
regressions cover pending/save/Undo and rejection before side effects. Full
Windows linguistic/BSTR/error identity and Jet SQL Replace remain open. See
[REPLACE_RUNTIME.md](REPLACE_RUNTIME.md).

Supported native Len string measurements/InStr positions now return Long;
StrComp returns Integer, preserving Null. Actual Qt/VBA/DAO parameter/SQL tests
verify numeric values and VarType codes through pending/save/Null/Undo. Len
numeric/UDT storage sizes and full Windows/SQL subtype propagation remain open.
See [STRING_RESULT_TYPES_RUNTIME.md](STRING_RESULT_TYPES_RUNTIME.md).

Native IsObject preserves object-reference context without resolving scalar
defaults; Nothing/closed stored references remain True. Explicit invalid member
getters retain their native errors. Actual Qt/VBA/DAO parameter/SQL regressions
verify object kinds and pending-source preservation. Full class/COM/ADO/Windows
object identity remains open. See [ISOBJECT_RUNTIME.md](ISOBJECT_RUNTIME.md).

New MDB imports preserve Currency/Single/Double source metadata using the additional
MDBTools Access-schema labels, resolving the MySQL FLOAT ambiguity. All 43 Currency and 57 Double
original columns are identified; conflicting/missing mappings diagnose
before import. SQLite REAL conversion/storage and native numeric/SQL precision
are not made exact by this metadata recovery; existing projects are not migrated.
See [JET_FLOATING_METADATA.md](JET_FLOATING_METADATA.md).

Verified direct-table Currency fields now expose fixed Currency values and
Field.Type=dbCurrency through a checked lossless REAL encoding strictly inside
+/-2^39. Native text load/assignment/copy preserve exact four-decimal units.
All 121613 original non-NULL Currency field values match independent CSV decimal
row/NULL multisets. Qt/VBA/DAO tests cover pending/save/OldValue/Undo/range checks.
Full-range/SQL operators/parameters/query subtype/locale/Windows money identity
remain open. See [CURRENCY_FIELDS_RUNTIME.md](CURRENCY_FIELDS_RUNTIME.md).

Currency/Money query parameters now expose checked fixed Currency Value and
Integer dbCurrency Type, binding numerically through the lossless +/-2^39 REAL
interval. Native DAO/SQL and actual bound Qt/VBA UPDATE/requery regressions cover
rounding/ranges/Null/error retention and numeric comparison. Full-range/SQL money
operators/formatting/result typing/Windows identity remain open; Currency LIKE
conversion diagnoses 3251. See [CURRENCY_PARAMETERS_RUNTIME.md](CURRENCY_PARAMETERS_RUNTIME.md).

Unqualified SQL CCur now shares checked native Currency conversion and bounded
lossless REAL output. Actual SQL/DAO/Qt/VBA regressions cover decimal rounding,
large values/errors/aliases and a failed multi-row UPDATE without partial writes.
Full query Currency subtype/SQL arithmetic/locale/Windows identity remains open.
See [CURRENCY_SQL_RUNTIME.md](CURRENCY_SQL_RUNTIME.md).

Currency SQL object references now bind numerically through the checked encoder.
Proven root CCur/Currency-parameter projections retain Currency into DAO, query-
table display and read-only Qt form/report fields, including empty/filter/TOP/
DISTINCT/saved-query paths. Ambiguous/wildcard/set-operation results keep unknown
types. Full SQL typing/object grammar/money arithmetic/Windows parity remains
open. See [CURRENCY_QUERY_TYPES_RUNTIME.md](CURRENCY_QUERY_TYPES_RUNTIME.md).

Proven Currency SELECT sums now use native checked integer units, preserving
Currency/Null and diagnosing overflow/output range. Actual SQL/DAO/Qt grouping/
cancellation/parameter/empty/report tests and all 43 independent original Currency
CSV sums pass. General implicit field/view/complex aggregate inference and full
SQL money/Windows semantics remain open. See [CURRENCY_SUM_RUNTIME.md](CURRENCY_SUM_RUNTIME.md).

Simple single-physical-table query projections now use verified live Currency
metadata for plain/qualified/aliased fields and exact Sum(column), carried into
DAO and read-only Qt controls. Plan changes/schema refresh invalidate stale
hints. All 43 original sums pass without explicit casts. Joins/views/HAVING/
nested/complex/action SQL typing and full money/Windows identity remain open.
See [CURRENCY_SCHEMA_QUERIES_RUNTIME.md](CURRENCY_SCHEMA_QUERIES_RUNTIME.md).

Proven Currency SELECT operands now use checked fixed-point +, - and *, preserving
precedence/associativity, four-decimal multiplication rounding, Null and Currency
result types into DAO/Qt. Actual SQL/report tests include Sum(a+b) and range/
overflow checks. Mixed/division/complex/action/full-range/Windows SQL arithmetic
remains open. See [CURRENCY_ARITHMETIC_SQL_RUNTIME.md](CURRENCY_ARITHMETIC_SQL_RUNTIME.md).

Simple physical-table UPDATE now checks Currency targets before writing, with
four-decimal conversion and exact proven RHS arithmetic. DAO/VBA/Qt regressions
cover parameters, aliases, simultaneous assignments, NULL and statement-wide
rollback on conversion errors. Joined/append/complex/full Jet semantics remain
open. See [CURRENCY_UPDATE_RUNTIME.md](CURRENCY_UPDATE_RUNTIME.md).

Simple single-row physical-table INSERT VALUES now converts Currency targets,
preserving parameter order and exact proven Currency expressions. Actual DAO/
VBA/Qt tests cover explicit/positional/reordered columns, omitted/NULL values,
rounding and failed inserts without partial rows. INSERT SELECT, complex/default/
constraint/full Jet semantics remain open. See
[CURRENCY_INSERT_RUNTIME.md](CURRENCY_INSERT_RUNTIME.md).

Bound native forms now expose CurrentRecord as a read-only Long position,
including filtered navigation and the supported new-record row. Real VBA/Qt/SQL
Current-event auditing verifies types and pending edits without implicit saves.
Unbound/design/report/full Windows behavior remains outside this bounded path.
See [CURRENT_RECORD_RUNTIME.md](CURRENT_RECORD_RUNTIME.md).

Simple physical-source INSERT SELECT append now converts Currency targets and
freezes positional source rows, including exact Currency self-appends and atomic
multi-row failure. The existing insertion regression covers this extension.
Joins/TOP/nested/view/HAVING/set-operation/full Jet append semantics remain open.
See [CURRENCY_INSERT_RUNTIME.md](CURRENCY_INSERT_RUNTIME.md#bounded-insert-select-append).

Native VBA and SQL now share bounded Gregorian DateDiff and seven date/time
component functions, accepting unambiguous ISO text with Integer/Long VBA
return types and integer SQL results. Existing core/query tests cover leap days, boundaries, negative
results, Null, options, aliases and overflow. Date Variants, numeric/locale
coercion and complete calendar/Windows semantics remain open. See
[CALENDAR_FUNCTIONS_RUNTIME.md](CALENDAR_FUNCTIONS_RUNTIME.md).

Native bound forms now apply OrderBy/OrderByOn through actual table/query
execution, preserving table row identity, pending validation and filtered results.
The existing CurrentRecord test covers VBA activation, invalid-state retention,
initial imported order and query sorting. Expressions, qualified query outputs,
Jet collations/report APIs/full Windows behavior remain open. See
[FORM_ORDER_RUNTIME.md](FORM_ORDER_RUNTIME.md).

The actual app table editor now checks source types and Currency conversion,
preserves long Memo editor text and protects/allocates verified AutoNumbers.
The custom query editor uses the Access compiler/results and DAO action path,
including declared parameter prompts and saved action definitions. Existing
Qt field/main-window tests exercise both real editors on synthetic projects.
Full datasheet/formatting/constraints/Jet/designer/Windows identity remains open.
See [APP_EDITORS_RUNTIME.md](APP_EDITORS_RUNTIME.md).

The native app table view now provides Boolean checkboxes/Space and Ctrl+S with
checked live-editor completion before manual-save transactions. Existing GUI
regressions verify pending/committed -1/0, actual shortcut persistence and invalid
Currency editor retention. Full Yes/No constraints, per-record events, undo and
Access keyboard/visual identity remain open. See
[APP_EDITORS_RUNTIME.md](APP_EDITORS_RUNTIME.md#checkbox-and-keyboard-cell-completion).

Pending datasheet changes now block tab/window/project close and VBA Quit before
native forms unload. Real main-window/VBA regression checks cache/live-editor
retention, unchanged-editor allowance, Quit 2501 with retained caller and close
after revert. Full automatic record saving/events/undo/window identity remains
open. See [APP_EDITORS_RUNTIME.md](APP_EDITORS_RUNTIME.md#pending-datasheets-participate-in-close-preflight).

DAO Database/QueryDef.RecordsAffected and loaded Recordset.RecordCount now return
checked VBA Long; Fields.Count returns Integer. CREATE/ALTER/DROP report zero
rather than stale SQLite changes; DML and make-table counts remain intact.
Existing DAO/main-window regressions verify types/counts and actual SQL-editor
feedback. Full cursor/count/DAO metadata/Windows identity remains open. See
[DAO_COUNTS_RUNTIME.md](DAO_COUNTS_RUNTIME.md).

Bounded physical-table SQL UPDATE/INSERT/append now checks verified source Byte,
Integer, Long, Boolean and canonical date targets alongside Currency. Shared
rounding/ranges/Null conversion and whole-statement rollback are exercised by
actual DAO SQL; unmarked ordinary SQLite storage is preserved. Date years are
limited to 0100–9999. Date Variants, locale/complex/general expression coercion
and full Jet/Windows semantics remain open. See
[TYPED_SQL_WRITES_RUNTIME.md](TYPED_SQL_WRITES_RUNTIME.md).

## Native Variant array subset

One-dimensional `Array()` values now support Option Base 0/1, bounds, indexed
reads/writes, type inspection and whole-array ByRef/ByVal/value copies.
One-dimensional Variant `ReDim`/`ReDim Preserve` and `Erase` are also implemented.
See [precise boundaries](VARIANT_ARRAY_RUNTIME.md). Typed declarations,
multidimensional arrays and full array For Each parity remain unsupported.

## Native text lists and reverse string functions

VBA Split/Join/Filter use real String arrays; bounded array For Each includes
live later-element updates on named sources. See [array limits](VARIANT_ARRAY_RUNTIME.md).
InStrRev/StrReverse execute in both VBA and compiled native SQL, including scalar
projection aliases and typed error transport; see [string boundaries](REVERSE_STRING_RUNTIME.md).
These changes do not establish full array, Unicode BSTR or Jet function parity.

## Optional parameter defaults

Sub/Function calls support omitted trailing Optional scalar arguments with
explicit literal defaults and existing ByRef/ByVal binding for supplied values.
See [exact supported declaration and call boundaries](OPTIONAL_ARGUMENT_RUNTIME.md).
Missing/IsMissing, positional holes, named arguments and full Optional/ParamArray
semantics remain open.

## Native VBA Like

Like now executes in VBA expressions/conditions with wildcards, digit matching,
character lists/ranges/negation, module comparison mode, Null and checked invalid
patterns. See [exact Unicode/collation boundaries](VBA_LIKE_RUNTIME.md).
This does not replace the existing SQL LIKE translator or establish full locale parity.

## ParamArray evaluation

Native VBA Sub/Function calls pack variable trailing arguments into a real
zero-based Variant array for indexing, For Each, Join and ByVal forwarding.
Slot/reference mutation and full argument binding remain open; see
[precise supported ParamArray scope](PARAMARRAY_RUNTIME.md).

## Missing values and named calls

Optional arguments without explicit defaults, distinct Missing/IsMissing and
positional holes now execute natively. User-procedure named calls preserve ByRef
binding; supported native DoCmd methods bind named values through declared
signatures. See [boundaries](NAMED_ARGUMENT_RUNTIME.md). Built-in/DAO/COM method
coverage and full declaration/SCODE/Windows semantics remain open.

## Native document OpenArgs

Text/Null OpenArgs now reaches independent native form/report instances before
opening events and is readable through Me, instance scope and Forms/Reports.
Reopening after close accepts another value. Nontext values, mutation and other
opening modes remain bounded; see [OpenArgs runtime](DOCUMENT_OPENARGS_RUNTIME.md).

## Native RunSQL transaction option

DoCmd.RunSQL evaluates UseTransaction and uses an action savepoint for True;
False adds no wrapper. Existing outer DAO transactions remain owned by their
caller. See [SQLite/Jet boundaries](DOCMD_RUNSQL_TRANSACTION_RUNTIME.md).

## Native hidden forms

OpenForm acHidden initializes a real form and its data/events without showing
or activating it. The hidden instance remains readable through Forms and can
be closed by name; see [remaining mode boundaries](HIDDEN_FORM_RUNTIME.md).

## Named native builtins and Nz

Selected native builtins bind named parameters with retained optional/error
semantics. Nz(Null) now preserves Empty in VBA while query expressions keep
zero-length text. See [supported names and reference](NAMED_BUILTIN_RUNTIME.md).

## Native dialog and window continuation

Form acDialog waits for actual close/hide while retaining Unload cancellation;
acIcon uses minimized Qt windows. Hidden reveal/re-hide, read access to imported/
overridden Modal/PopUp and active-document Minimize/Restore/Maximize are native.
See [event/desktop/Quit boundaries](DIALOG_WINDOW_RUNTIME.md).

## Quit and native execution scopes

Successful native application Quit stops active VBA callers/operands and waiting
dialog frames; cancelled Quit preserves error handling. Startup Quit prevents
later Open/Load/show work in native construction scopes. Typed array function
returns remain arrays. See [remaining shutdown scope](QUIT_EXECUTION_RUNTIME.md).

Native DAO named arguments are supported for the implemented parameter subset of
Database.OpenRecordset (Name, Type), Execute (Query, Options), temporary
CreateQueryDef (Name, SQLText), QueryDef.OpenRecordset (Type), Execute (Options),
TableDef.OpenRecordset (Type), and Recordset FindFirst/Last/Next/Previous
(Criteria). This does not extend the supported recordset modes or optional DAO
options. Full DAO signatures/default/coercion parity remains open.

Recordset.Move supports signed Long Rows and optional native byte-array Start
bookmarks on the existing materialized recordset/clone model. It checks pending
edits, foreign/stale tokens, empty/current-boundary movement and Long overflow;
movement beyond available rows sets BOF/EOF. This does not implement server-side
cursors, dbForwardOnly, bookmark constants or exact Windows locking behavior.

Native Recordset.EditMode reports the actual cursor buffer state as Integer
(0/1/2), with dbEditNone/dbEditInProgress/dbEditAdd constants. Updatable reports
the existing native table/snapshot write capability. Both are read-only; this
is not a claim of full DAO server-permission/locking capability detection.

Native form CommandButton.Default is imported (including control templates),
readable/writable via VBA, and used for Enter from supported form editors after
control validation. A focused command button takes priority; hidden/disabled
defaults are skipped. Cancel/Escape, multiline EnterKeyBehavior, subform focus
and full Access 2003 keyboard/rendering equivalence remain open.

Imported labels nested beneath supported TextBox/ComboBox/ListBox controls keep
their attachment and use native Qt buddy mnemonics for Alt-letter editor focus.
Detached-label heuristics, duplicate mnemonic cycling and full Windows label
mouse/accessibility/printing behavior are not claimed.

OpenForm DataMode now supports native acFormEdit/acFormReadOnly policy overrides
before startup events for newly opened forms; see FORM_DATA_MODE_RUNTIME.md.
Writable-query DataEntry, changing mode on an already-open instance and full Windows
view/locking semantics remain open. Edit does not make read-only queries writable.

Native CommandButton.Cancel now supports imported/template metadata, VBA
read/write and real Escape-to-Click behavior with focused TextBox buffer undo;
setting a new cancel/default button clears the corresponding previous policy.
See COMMAND_BUTTON_KEYBOARD_RUNTIME.md for verified scope and remaining limits.

Native DAO dynaset/snapshot adapters expose Long AbsolutePosition: zero-based
read/write within populated rows, -1 when no current row, independent clone
positions, checked Long conversion and pending-edit protection. Invalid positions
currently raise native 3021. Full deleted-row, lazy-population, forward-only,
ODBC and exact Access 2003 error/reference semantics remain open.

Native editable DAO recordsets expose read-only LastModified as their own last
successful Update/AddNew bookmark, validated against the current native token
family. Snapshot/never-modified/stale cases raise explicit errors. Native DAO
Update now uses SQLite RETURNING rowid and preserves shared bookmark identity
when an Integer primary-key edit changes rowid. Direct external SQL/trigger-driven
identity replacement and exact Windows LastModified edge/error semantics remain
open; native form saves now also preserve committed key-change identities; see
FORM_KEY_IDENTITY_RUNTIME.md.

Native direct-table Form.DataEntry now imports/read/writes the actual entry
session with bound SQL membership, real insert/key-change persistence and
AllowAdditions dependency. DataEntry=False removes active filter/sort and restores
all rows. Writable query entry sources remain unsupported; native acFormAdd now starts
this entry session with imported edit/delete policies retained. See
DATA_ENTRY_RUNTIME.md for limits and verification.

Native Form.NewRecord now returns read-only Integer -1/0 through qualified,
unqualified and legacy field routes. True requires an actual native addable
virtual row; an empty non-addable source no longer reports a new record. Full
Windows no-current/unbound/view edge behavior remains outside the verified scope.

Native Recordset Field.OrdinalPosition exposes checked Integer metadata in actual
field/projection order, read-only with closed-owner protection. TableDef reorder
writes/persistent ordinal gaps and full DAO metadata semantics remain open.

ReDim Preserve now supports existing native String/Byte arrays held by Variant,
retaining typed values and checked bounds with correct empty String/zero Byte
new cells. Non-Preserve typed resize, typed declarations, Byte Erase and full
array/COM semantics remain open; see VARIANT_ARRAY_RUNTIME.md.

Optional Let now supports the existing native value-assignment forms, including
array elements and object properties; invalid non-assignment Let forms fail
before execution. Property Let declarations and full VBA grammar remain open.

Native numeric And/Or now enforce checked VBA Long bounds, preserve Integer
results for Integer pairs, and return Long for the other supported non-Null
numeric cases. Complete Byte/mixed/Null promotion matrices remain open.

Native expressions now parse checked &H/&O integer patterns, %/& suffixes and
supported Optional based-literal defaults, preserving Integer/Long identity and
signed bit patterns. Full literal/unary/suffix inference and raw Jet SQL based
literal translation remain open.

Native Asc/AscW/Chr/ChrW are now available in SQL through shared VBA callbacks;
Chr$/ChrW$ and named character arguments work in VBA. ANSI remains ASCII-only
and isolated surrogate output is explicitly unsupported. InStrRev Null string
arguments now return Null in VBA/SQL; StrReverse(Null) still raises 94. Full
codepage/BSTR/result-metadata equivalence remains open.


### Native repetition functions (2026-11-03)

Space/Space$ and String/String$ execute in the C++ VBA evaluator; plain Space and
String also use the registered native SQL scalar bridge. Named Number/Character
arguments, checked Long counts, first UTF-16 text unit, ASCII numeric codes with
modulo 256, embedded NUL and String Variant Null versus dollar-variant error 94 are
implemented. Negative counts raise 5, counts above 1,000,000 raise explicit 3251;
ANSI bytes above 127 and isolated surrogates remain unsupported. Full Windows
codepage and argument-evaluation edge parity is not claimed.
Existing core/DAO/query-model targets pass 3/3 (1.77 s); desktop all-target build
passes. Prior character SQL GUI-free follow-up passed 1/1 (1.37 s). GUI-free build/test passes 1/1 (1.45 s). The Debian package passes contents,
permissions, Build-ID and pinned SQLite backend smoke checks; all four original
fixture SHA-256 checks pass. Full desktop suite passes 48/48 in one run (85.23 s).


## Native Hex/Oct conversion and repetition correction (2026-11-03)

Hex/Hex$/Oct/Oct$ now convert native scalar values with checked, rounded Long
coercion, uppercase hexadecimal, zero/Empty and Variant Null versus dollar error
94. Integer inputs preserve 16-bit negative patterns, other supported numeric
inputs use 32 bits. Plain Hex/Oct execute through the native SQL bridge on real
row values, preserving Null. Named Number arguments are supported. This does not
complete decimal literal inference, 64-bit VBA or all SQL type metadata.
Space(Null) now raises 94 instead of the previously over-broad Null propagation;
String retains its documented Variant Null behavior. Existing core/DAO/query-model
checks pass 3/3 (1.67 s), including explicit Integer/Long conversions. Desktop
all-target build passes. No new test target, full-suite or package repeat for
this increment. Last grouped verification remains 68bdfd0: 48/48, GUI-free,
Debian package and four fixture hashes.


## Native VBA mathematical functions (2026-11-03)

Sqr, Exp, Log, Sin, Cos, Tan and Atn now execute through the C++ evaluator with
scalar/default-property conversion, named Number arguments and Double results.
Trigonometry uses radians. Null raises 94, invalid square-root/log domains raise
5, non-finite arguments/results raise 6, and wrong arity raises 450. Existing
core tests exercise composed identities, argument conversion and runtime errors;
core/DAO pass 2/2 (1.77 s), GUI-free passes 1/1 (1.37 s), desktop all-target build
passes. No new test target/full-suite/package repeat. This increment does not
add their Jet SQL registration or prove all Access/Windows floating-point edges.


## Native datasheet record shortcuts (2026-11-03)

The table view now handles Ctrl+Plus (including Ctrl+Shift+Equals) and Ctrl+Minus
through the same insertion/deletion paths used by its buttons. Both validate any
live editor before changing the manual-submit model; failed validation preserves
the editor and row set. Insertion selects the new row. Existing main-window GUI
regression inserts/saves/deletes/saves a record with real key events and verifies
SQL persistence. Main-window/currency-field checks pass 2/2 (1.85 s); additional
invalid-editor guard check passes 1/1 (0.26 s). Desktop affected targets build.
This retains explicit Save semantics; Access automatic record save, complete
selection/deletion confirmation behavior and the full datasheet keyboard oracle
remain open. No new test target or unrelated full-suite/package repeat.


## Mathematical SQL bridge and grouped verification (2026-11-03)

Sqr/Exp/Log/Sin/Cos/Tan/Atn are registered native SQL functions using the same
bounded evaluator and scalar binding as VBA. The shared scalar result emitter
preserves Double/Single values instead of truncating them to integers. Existing
query-model regression checks fractional square root, composed log/exponential,
trigonometry on actual row inputs and typed domain errors. DAO/query-model checks
pass 2/2 (0.33 s). Rebuilt desktop full suite passes 48/48 in one run (84.72 s).
The Debian package passes contents/permissions/Build-ID/backend checks; all four
original SHA-256 checks pass. GUI-free runtime was checked at 3c4a93a (1/1,
1.37 s); following changes are GUI/data-layer only. The catalog now explicitly
records 243/312 bounded points; these unequal items are not measured overall
Access parity. Complete SQL/Windows floating-point/Null/type metadata remains open.


## Native Val parsing (2026-11-03)

Val no longer delegates its whole behavior to std::stod. It strips documented
ASCII spacing, parses a decimal prefix with a fixed period separator, recognizes
hex/octal signed 16/32-bit patterns and checked decimal type suffixes, and reports
Null/overflow/integral-suffix mismatch instead of silently returning zero.
Named String arguments and the native SQL bridge are implemented. Non-integral
radix suffix conversions and 64-bit suffixes remain explicit 3251 boundaries;
full Access/locale/coercion/subtype edge parity is not claimed.
Existing core/DAO checks pass 2/2 (2.03 s), GUI-free passes 1/1 (1.31 s), and
SQL-model/DAO follow-up passes 2/2 (0.36 s) with real rows and fractional outputs.
Desktop affected targets build. No new target or full-suite/package repeat for
this increment; grouped baseline remains f574057 (48/48 and checked package).
Reference: Microsoft Learn Val function, including &HFFFF=-1, embedded spaces,
period-only decimal recognition and 50.5% type mismatch.


## Native bounded case conversion (2026-11-03)

LCase/UCase now preserve Variant Null; their new dollar variants reject Null with
94. Named String arguments, scalar/default-property conversion, checked arity,
embedded NUL and native SQL registration are implemented. ASCII conversion uses
explicit ranges rather than process-locale byte transforms. Valid non-ASCII text
raises 3251 because Windows locale mapping is not implemented; malformed UTF-8
raises 13. This replaces the former misleading unchanged/nonportable byte result,
not a claim of full Unicode/collation parity. Existing core/DAO/query-model checks
pass 3/3 (1.64 s). Desktop affected targets build; GUI-free build/test passes 1/1 (1.36 s). No new target/full-suite/package repeat.


## Native bounded DatePart (2026-11-03)

DatePart now returns actual Integer year/quarter/month/day-of-year/day/weekday/
week/hour/minute/second values on the existing Gregorian ISO calendar path.
Named Interval/Date/FirstDayOfWeek/FirstWeekOfYear arguments and optional holes
are supported. Week numbering implements the Jan-1 rule with explicit weekday;
other week rules/system regional settings raise 3251. Null date propagates Null;
invalid interval/options and required omissions remain typed errors. Native SQL
uses the shared evaluator. Core/DAO/query-model checks pass 3/3 (1.68 s),
GUI-free build/test passes 1/1 (1.29 s), desktop affected targets build.
No new target/full-suite/package repeat. Numeric Date coercion, actual Date
Variants, localized input, Hijri and full Access week-number oracle remain open.


## Native Switch/Choose (2026-11-03)

Switch returns the first true condition's Variant value or Null; Choose selects
by a checked rounded Long index or returns Null outside the available choices.
All argument expressions are evaluated, including unselected division-by-zero
errors. Native VBA preserves selected scalar types (including Currency); the SQL
bridge binds inputs and preserves supported scalar result kinds. Existing
core/DAO/query-model checks pass 3/3 (1.64 s), desktop affected targets build.
Object/default-property, all coercion/subtype and Windows edge behavior remain
open. No new CTest target. Grouped verification follows after these increments.


Grouped verification after native Val/case/DatePart/Switch/Choose: rebuilt desktop
suite passes 48/48 in one run (84.92 s); GUI-free build/test passes 1/1 (1.42 s).
The Debian package passes contents/permissions/Build-ID/pinned backend checks;
all four original fixture SHA-256 checks pass. Catalog 247/316 bounded items;
full Access parity remains open. No new CTest targets or public artifacts.


## Native Window menu (2026-11-03)

The main window now exposes a functional Window menu with next/previous native
document and Close All actions. Cycling reuses guarded document activation;
existing keyboard dispatch retains KeyPreview cancellation. Close All respects
modal windows, pending datasheet edits and form Unload cancellation before
removing clean tabs. Existing main-window GUI regression checks menu activation,
Unload refusal, dirty-table refusal and successful report/tab closing; passes
1/1 (1.43 s), affected desktop targets build. No new target/full-suite/package
repeat. Complete Access 2003 MDI, window lists/arrangement and pixel/menu/toolbar
parity remain open. This increment does not affect the GUI-free runtime.


## Direct native document selection (2026-11-03)

The Window menu now rebuilds a numbered, checked list of visible project-owned
nonmodal forms/reports. Entries activate the chosen real document through the
same validation/lifetime guard as keyboard cycling; minimized windows restore.
The database-window menu entry uses that guard too. Captured QPointers prevent
stale entries from accessing disposed documents, and post-event checks respect
closed targets or activation redirected by VBA. Existing main-window GUI checks
menu selection/check state, database focus, stale disposed report entries and all
prior cycling/cancellation cases; passes 1/1 (1.59 s), affected targets build.
Full MDI/arrangement, hidden/modal/datasheet window-list and Access pixel/event
oracle remain open. GUI-free runtime unchanged; no new target or full-suite repeat.


## Native bounded Date Variant foundation (2026-11-03)

The runtime now has a distinct Date Variant carrying an OLE Automation double,
checked CDate from numeric/Empty/Boolean values or unambiguous Gregorian ISO text,
and Date locals/module fields/parameters/function results. VarType/TypeName and
numeric conversion preserve Date identity; Len(Date) is 8 and IsNumeric(Date) is
false. Calendar functions accept the typed value. ISO text conversion handles
the OLE negative-fraction convention; ordinary supported Date addition/subtraction
preserves Date or yields a numeric date difference. Negative fractional arithmetic
and text/date arithmetic remain explicit 3251 boundaries.

This is a runtime foundation, not completed Date compatibility: Now/Date/Time
still use their previous ISO text path until the adapters are migrated. Qt/DAO
Date reads/writes, date literals, localized CDate/CStr, complete arithmetic and
subsecond/OLE edge rules remain open. Rebuilt desktop suite passes 48/48 in one run (84.76 s), including the latest
Window-menu/list paths; GUI-free build/test passes 1/1 (1.39 s). Debian package
contents/permissions/Build-ID/backend checks and four fixture SHA-256 checks pass.
No new test target. These checks verify the bounded foundation, not full Date parity.


## Date clocks, Qt/DAO integration and physical query lineage (2026-11-03)

Now/Date/Time now return Date Variants. Declared Date fields read/write typed
values while SQLite retains normalized ISO civil timestamps. DAO Date parameters
retain their native type when read; direct projected Date columns inherit verified
main-database physical lineage from SQLite metadata, not guesses from cell text.
Compound queries, attached sources and nonphysical Date expression inference stay
outside this inference. Format's existing bounded masks accept Date serials;
native SQL object values bind normalized timestamps. Calculated controls display
Date values and retain their typed cache; unbound assignments preserve the serial.
The ISO General-Date presentation treats epoch/midnight as time/date as applicable;
regional CDate/CStr, actual Access formatting and subsecond/OLE edges remain open.

Core/DAO initially passed; new GUI regression independently verifies bound-field
arithmetic/save, declared Date parameter update and DAO query Date readback against
actual SQL. DAO/query-model/GUI pass 3/3 (0.67 s). Desktop full run passed 47/48
(65.35 s), exposing a missing generic calculated-Date display adapter; after its
fix affected GUI plus both original targets pass 3/3 (52.73 s). All 48 targets
are covered by that run and follow-up, not a final single 48/48 run. GUI-free
build/test passes 1/1 (1.29 s). Debian package contents/permissions/Build-ID/backend checks pass, as do all four
original fixture SHA-256 checks. Catalog now 251/320 bounded items, not measured
overall Access parity.


## Native DateSerial/TimeSerial and SQL Date construction (2026-11-03)

DateSerial and TimeSerial construct real Date Variants with checked Integer
arguments and relative month/day/time normalization. Named Year/Month/Day and
Hour/Minute/Second arguments are supported. DateSerial uses Gregorian four-digit
years; two-digit regional windows raise 3251. Null/arity/range errors stay typed.
DateSerial/TimeSerial/CDate execute through the native SQL scalar bridge, which
normalizes Date storage and retains DATETIME metadata for direct constructor
projections. The projection typer now covers Currency rewriting and supported
Date functions. Date SQL arithmetic still requires a typed operator rather than
SQLite's implicit string-to-number conversion; no complete expression/UNION/
locale/Hijri or Windows edge parity is claimed.
Core passes 1/1 (1.40 s); GUI-free build/test passes 1/1 (1.36 s); query-model/DAO
pass 2/2 (0.34 s), affected desktop targets build. Existing targets only; no
full-suite/package repeat after the prior grouped Date integration checks.


## Typed Date SQL projection arithmetic (2026-11-03)

Supported SELECT projections now route Date+number, number+Date, Date+Date and
subtraction through registered native operators with bound scalar arguments.
Date differences return DOUBLE metadata; supported Date results retain DATETIME.
Simple physical-source fields and declared Date parameters participate in typing.
Ordinary function arguments and parenthesized numeric expressions retain native
Date child rewrites; Date()/Time()/Now() projections retain Date metadata too.
Focused SQL/DAO tests cover leap-day progression on actual Date rows, differences,
Null propagation, nesting and numeric-left addition: 2/2 (0.39 s).
Rebuilt desktop suite passes 48/48 in one run (93.31 s). Debian package contents,
permissions, Build-ID and pinned backend checks pass; all four original hashes
pass and the repository is private. GUI-free runtime remains the checked
constructor build (1/1, 1.36 s); this increment changes GUI/data-layer code only.
Native Date arithmetic in filters/actions/CASE/CAST/aggregates, additional Date
operators, complete dynamic promotion/comparison and Windows/OLE edge rules remain
open. Negative fractional arithmetic and unsupported typed operators remain
explicit boundaries; no complete Jet Date expression compatibility is claimed.


## Native DateAdd and numeric Date-difference SUM (2026-11-03)

DateAdd returns real Date Variants for Gregorian year/quarter/month/day/week/
hour/minute/second intervals, clamps month/year endpoints, rounds the interval
count to Long and preserves the input time fraction. Civil-day normalization
handles negative OLE fractions without treating them as ordinary signed days.
The w interval adds every calendar day. Named Interval/Number/Date arguments,
Null date propagation and typed range/argument errors are implemented in VBA;
native SQL registration retains DATETIME metadata for direct projections.
Numeric SUM expressions now preserve native Date-difference child rewrites and
DOUBLE metadata. Date-valued SUM remains unsupported and raises 3251 even on an
empty input, rather than letting SQLite sum year-shaped text. Complex source/
aggregate typing, region/Hijri settings, storage subsecond precision and complete
Windows/OLE coercion remain open.
Core/DAO/query-model targets pass 3/3 (1.82 s); GUI-free build/test passes 1/1
(1.43 s), affected desktop targets build. No new target/full-suite/package repeat
for this increment; last grouped projection baseline remains a8d3061.


## Date Variant promotion and lossless whole-second storage boundary (2026-11-03)

Supported Date Variant addition/subtraction now promotes an out-of-Date-range
finite result to Double, as documented for VBA subtraction. Declared Date
assignment remains checked. Fixed-Date SQL projections reject this overflow
instead of giving the numeric value Date metadata. Bound Date fields also accept
numeric OLE series values and normalize their storage. Date values carrying a
subsecond fraction are rejected with 3251 before field mutation, rather than
silently truncating them; complete subsecond backing remains open.
Existing core/Qt-field/query-model checks pass 3/3 (1.62 s); GUI-free build/test
passes 1/1 (1.36 s), affected desktop targets build. Real GUI checks verify numeric
Date storage and unchanged state after precision refusal. No new target/full-suite/
package repeat. README and handoff link the consolidated DATE_RUNTIME.md limits;
last grouped baseline remains a8d3061, with following focused increments recorded.


## Typed native calendar values and civil Date input (2026-11-03)

Unbound native MSCAL calendar values now return Date Variants and accept Date or
numeric OLE input in addition to supported ISO text. The date-only/time/bound
restrictions remain explicit. Both original popup-calendar event paths copy the
actual selected Date to their independently focused target and preserve SQL
object-value comparison behavior. Synthetic checks verify native/number setters,
Date type, invalid input preservation and the unchanged bound-calendar boundary.

Date-field ISO parsing now validates QDate/QTime independently instead of asking
QDateTime to resolve a local instant. This accepts valid civil times inside a
system daylight-saving gap, with canonical unchanged SQL storage. App/affected
GUI targets build; Qt-field plus both original targets pass 3/3 (49.75 s) under
TZ=Europe/Berlin. No new CTest target or GUI-free repeat (runtime unchanged).
Full regional/calendar/ActiveX/Date edge parity remains open.


## Native DateValue/TimeValue and DateDiff weekday interval (2026-11-03)

DateValue removes the time from a native Date or supported ISO value; TimeValue
returns its time as Date. Named Date/Time arguments, DateValue Null error 94 and
the documented TimeValue Null propagation are implemented. Other scalar input
families remain 3251 boundaries, not locale guesses. Unambiguous 24-hour HH:mm
and ISO date/time without seconds are accepted and normalized. Regional/month-
name/12-hour parsing and full negative-OLE representation rules remain open.
DateDiff w now counts whole origin-weekday intervals from Gregorian day distance,
including reverse spans and time-of-day-independent boundaries. System/week-year
options retain their previously documented restrictions.
Core/DAO/query-model checks pass 3/3 (1.82 s); GUI-free build/test passes 1/1
(1.42 s), affected desktop targets build. Existing targets only; no additional
full-suite/package repeat before the next grouped language check.


## Native VBA Date literals and grouped language verification (2026-11-03)

The lexer parses supported #...# literals as actual Date constants: four-digit
ISO or numeric US month/day/year dates and unambiguous 24-hour times, including
Date/time values and optional Date parameter defaults. Validation happens during
syntax inspection, before procedure side effects. Inline-branch colon splitting
preserves colons inside Date literals. Regional/month-name/two-digit forms and
complete literal/numeric-suffix grammar remain open.
Core checks pass 1/1 (1.41 s); rebuilt desktop suite passes 48/48 in one run
(87.46 s); GUI-free build/test passes 1/1 (1.43 s). This grouped run also covers
DateAdd, Date Variant promotion/storage boundaries, calendar/DST changes and
DateValue/TimeValue. Debian package contents/permissions/Build-ID/pinned backend
checks pass and all four original fixture hashes pass. Existing CTest targets
only; no full Access, Windows or new distribution-matrix parity claim.


## General native VBA colon statement chains (2026-11-03)

Procedure bodies now parse top-level colon-separated statements through the same
lexical/block/label pipeline as separate lines. Chained declarations, For/Next,
GoTo and inline labels retain their physical source line and control-flow state;
inline If branches retain their own branch scope. Quoted/bracketed/Date-literal
colons and named := arguments are preserved. Rem comments at statement boundaries
and after Then cannot turn comment text into executable branches.
Existing core/DAO tests execute chained flow, labels, Date/time literals and real
DAO method calls. GUI-free build/test passes 1/1 (1.41 s). The corpus-audit binary
was rebuilt against the changed runtime and its original syntax target passes
1/1 (0.94 s). Affected app/core/DAO targets build. No new CTest target/full-suite/
package repeat after the checked 63f55e4 language baseline. Complete module-header/
property/class grammar and all comment/label edge cases remain open.


## Qualified VBA standard-library calls (2026-11-03)

Implemented intrinsic functions/constants can be addressed through VBA., including
named arguments and Date values. Local/host objects named VBA retain precedence.
Access-only functions and unknown qualified members are rejected explicitly;
module-qualified paths such as VBA.Strings are outside this increment.
Existing core/DAO checks pass 2/2 (1.68 s); GUI-free core passes 1/1 (1.58 s).
No new CTest target or full-parity claim.


## Native DoEvents and nested Quit (2026-11-03)

Native imported form/report VBA now dispatches actual pending Qt events and
returns Integer zero. Expression/qualified calls and bare statements are
supported; direct-call preflight preserves user procedure resolution. Native
document retention spans the pump; nested accepted Quit stops the suspended
procedure even with Resume Next. Calculated expressions and hosts without an
event loop reject the operation explicitly. See DOEVENTS_RUNTIME.md for limits.
Core/layout checks pass 2/2 (16.05 s), including a queued SQL insertion and
nested Quit. GUI-free build/core passed before the final direct-call preflight
correction (1.40 s); its repeat follows with the next clock increment. Existing
CTest targets only; no full Windows message/reentrancy parity claim.


## Native VBA Timer and grouped event/clock verification (2026-11-03)

Timer/Timer()/VBA.Timer() now return actual local-midnight elapsed wall-clock
seconds as Single, including fractions. Core compares the result with the actual
local clock and checks type/arity. Windows tick-resolution and SQL Timer remain
open; see DATE_RUNTIME.md.
Rebuilt complete desktop suite passes 48/48 in one run (87.67 s), including the
DoEvents preflight correction and nested Quit. GUI-free build/core passes 1/1
(1.28 s). Debian package contents, permissions, matching Build-ID and pinned
SQLite backend smoke checks pass. All four original SHA-256 checks pass.
No new test target or new distribution-matrix/full Access parity claim.


## Native ordered VBA Collection (2026-11-03)

New Collection/New VBA.Collection construct native ordered Variant collections
with Add/Item/Remove/Count, default one-based/keyed access, named/omitted Add
arguments, Before/After, shared aliases, class inspection and For Each. ASCII
keys ignore case; duplicate/index errors preserve the collection. Native Date,
Currency, Null and DAO object references retain their identity/type. A scoped
enumeration lease rejects mutation before effects; nested collections (including
inside arrays) are rejected until cyclic lifetime is supported. See
COLLECTION_RUNTIME.md for regional/class/module/parameter limits.
Core/DAO pass 2/2 (2.02 s); GUI-free core passes 1/1 (1.41 s). Rebuilt app and
native layout/original DAO regression pass 2/2 (51.49 s). Existing test targets
only. Last grouped full suite/package/hash baseline remains aafff2b; no full
Collection/COM/Access parity claim.


## Typed Collection module/call bindings (2026-11-03)

Collection-typed module fields retain native references between procedure calls.
Required Collection parameters and function results enforce the native class.
ByVal shares mutations but keeps replacement local; ByRef replaces the caller
reference. Nothing remains a valid typed initial/reference value. Optional
object parameters and complete class/reference semantics remain unsupported.
Core/DAO targeted checks pass 2/2 (1.88 s); rebuilt GUI-free core passes 1/1
(2.28 s). Complete rebuilt desktop suite passes 48/48 in one run (92.80 s).
Debian package contents/permissions/Build-ID/pinned backend checks pass, all four
original hashes pass, and the remote repository remains private. Existing test
targets only; no complete Collection/Access or new distro-matrix claim.


## Native form date/time insertion keys (2026-11-03)

Ctrl+semicolon inserts the current local ISO date; Ctrl+Shift+semicolon/colon
inserts the current local HH:mm:ss into the focused native form TextBox,
replacing its selection or inserting at the caret. Existing control buffers,
Change events, BeforeUpdate cancellation and record saving are retained.
Read-only/disabled/calculated/unsupported bound contexts and capacity overflow
refuse mutation. The existing layout target passes 1/1 (13.58 s), exercising
actual keys, Change, cancelled save, SQL Date/time persistence and read-only
refusal. App builds. Regional Access date masks and datasheet keys remain open
in this increment; full-suite/package baseline remains 8dc6afa.


## Datasheet clock keys and keyed cell restoration (2026-11-03)

Forms and native editable datasheet text cells share one Qt clock-key mapping.
The table view starts editing the selected cell; an already open delegate editor
preserves the caret/selection. Ctrl+S retains existing validation/manual-submit
behavior and persists the actual inserted value. Read-only/non-text editors do
not mutate. ISO date/whole-second time formatting remains bounded.
Saving a table now restores the selected existing record/column by its complete
primary key after model reset, including when sort order moves that record.
Search is bounded to one million fetched rows; absent/NULL/new-generated keys
and complete large-table/no-key selection semantics remain open.
Main-window regression passes 1/1 (1.65 s), proving real date/time keys, persisted
SQL values, partial replacement and keyed restoration as sorting moves the same
record twice. The shared-helper form layout target also passed in the preceding
integration run. App builds; existing targets only. Last full suite/package/hash
baseline remains 8dc6afa; no complete datasheet/regional-format parity claim.


## Optional Collections and native editing menu (2026-11-03)

Optional Collection parameters now bind an explicit/implicit Nothing default,
with typed IsMissing=False, named supplied arguments and independent omitted
ByRef storage. Non-Nothing defaults are rejected before effects. The main-window
Bearbeiten menu routes real editor key events for undo/redo/clipboard/delete/
selection, with live enabled states and read-only/password handling. It covers
owned text editors; shared MDI menus/range clipboard/full command bars remain
open. See COLLECTION_RUNTIME.md and EDIT_MENU_RUNTIME.md.
Core/DAO targeted checks pass 2/2 (1.84 s); main-window target passes 1/1
(1.90 s). Rebuilt complete desktop suite passes 48/48 in one run (88.53 s),
including form/datasheet clock keys and primary-key-based cell restoration.
GUI-free build/core passes 1/1 (1.76 s). Debian package contents/permissions/
matching Build-ID/pinned SQLite backend checks pass; all four original hashes
pass. Repository privacy is confirmed. Existing 48 targets only; no full Access
2003/Windows interaction or new distro-matrix parity claim.


## Acyclic nested Collection graphs (2026-11-03)

Collection items can now retain other native Collections directly or through
Variant arrays. Shared acyclic graphs preserve object identity; a bounded
visited-node traversal rejects only cycles back to the receiving owner before
mutation. Nested For Each retains independent enumeration leases, so editing a
child before its own enumeration remains possible. Direct/indirect/array cycles
still report 3251; complete cyclic/class lifetime semantics remain open.
Core/DAO pass 2/2 (2.29 s); GUI-free core passes 1/1 (1.77 s). Existing targets
only. Last full suite/package/hash baseline is local 4ba1819 (48/48, 88.53 s).
That commit's CLI push was blocked by an invalid GH_TOKEN; the authenticated
GitHub connector confirms private visibility and remote head 825047c. No raw
credentials were inspected; local development continues pending renewed CLI
authentication. No claim that these pending commits are already remote.


## Native Unicode-string LenB (2026-11-03)

VBA LenB/VBA.LenB and native SQL LenB count UTF-16 storage bytes for actual
Unicode strings (without a terminator), including supplementary characters and
embedded NUL. Results are checked Long values in VBA; Null propagates and Empty
returns zero. Native scalar defaults use the existing object conversion path.
SQL uses bound input values through the same native evaluator and participates
in saved-query/alias validation. Non-string storage layout, UDT padding and
named LenB binding remain explicitly unsupported; numeric inputs report 3251.
Core/DAO/query checks pass 3/3 (2.34 s), including real SQL column/NUL input and
projection alias arithmetic. GUI-free core passes 1/1 (1.55 s). Existing targets
only. No full VBA storage-layout or new package/full-suite claim after 4ba1819.
The CLI token remains invalid; these source commits are locally secured pending
authentication renewal, rather than claimed as already pushed.


## Precise joined UPDATE refusal diagnostics (2026-11-03)

Conflicting joined source values now retain error 3251 with a dedicated, fixed
explanation through both DAO and shared-query execution. Nullable-side outer-
join writes also have a distinct refusal instead of a generic grammar failure.
The private 100-query inventory records fixed reason codes, not raw SQL/values.
Actual original boundaries are one joined_update_source_conflict and one
joined_update_nullable_target. The executed count remains 77/100; no arbitrary
source row is selected and no Access equivalence is invented. Original binding/
dependency/source-column failures remain unresolved.
DAO/query/join/inventory checks pass 4/4 (6.91 s); after the final nullable-side
classification, join/inventory pass 2/2 (7.48 s). The ambiguity regression also
checks the stable explanation and unchanged target rows. Existing targets only.
Full suite/package baseline remains local 4ba1819, and CLI push authentication
still requires renewal.


## Native datasheet selection and checked text clipboard (2026-11-03)

Shift+Space/Ctrl+Space/Ctrl+A select a row/column/all cells. Complete selection
fetches remaining Qt batches within a bound; it does not silently stop at the
initial page. Ctrl+C and Copy menu dispatch create rectangular quoted TSV with
checked model values, exact supported Currency and round-trip numeric digits.
Invalid editors, binary/NUL/hidden/disjoint/oversized cells preserve the old
clipboard. Copy validates the active edit cache without submitting SQL changes.
See DATASHEET_CLIPBOARD_RUNTIME.md for text-format/paste/locale limits.
Main-window and Currency integration checks pass (including the final Currency
run 1/1, 0.38 s): actual keys/menu, 300-row complete fetch, TSV quoting, rejected
copy and real unchanged SQL storage. The rebuilt full desktop suite passes
48/48 in one run (87.26 s), covering nested Collections, LenB and join diagnoses
as well. Latest GUI-free core passes 1/1 (1.55 s); subsequent edits affect Qt
data/GUI code only. Debian contents/permissions/Build-ID/pinned SQLite backend
checks pass; all four original fixture hashes pass. Existing 48 targets only;
no Access/Windows clipboard or complete parity claim. CLI authentication is
still unavailable; local checkpoints are secured while an authenticated private
GitHub API transfer is investigated.


## Private Git API checkpoint published (2026-11-03)

The authenticated GitHub connector published d94ec0382b637dac6d306d5af87f02e8a8756c19
on development/access-compatibility without force. It contains the exact tested
source tree 9ab113b0ba92d0021944b5ef4b6a5b2e8a13fe33 of local b96e7a6: all 25
changed blob hashes and the complete tree match. Remote head and private
visibility were verified. This publishes the previously pending optional/nested
Collection, LenB, editor/datasheet and joined-update diagnostic increments.
48/48 CTests (87.26 s), GUI-free core, Debian checks and all four fixture hashes
pass. Local individual commits are preserved; API and local metadata/history
differ until normal Git synchronization becomes available. No credential values
were inspected or changed. See CODEX_HANDOFF.md for safe continuation.


## Checked range paste and retryable native table Save (2026-11-03)

Native datasheets accept bounded strict quoted TSV into existing cells of a
clean manual-submit model. Typed targets are preflighted as a complete rectangle;
malformed/ragged/read-only/out-of-range input is refused before cache mutation.
Paste does not submit SQL. Native table Save now retains all draft updates,
inserts and deletes after a later constraint failure and database rollback.
The previous Qt submitAll cache behavior could silently omit earlier rows on
retry. Protected SQL workers now run before any cache submission; mixed draft
and row-shift regressions prove correction and full retry, defaults/AutoNumber.
See DATASHEET_CLIPBOARD_RUNTIME.md for scope and inherited submitAll limitations.
Existing 48 targets only. Full build and suite pass 48/48 (86.21 s), GUI-free
TSV/core passes 1/1 (2.63 s), Debian payload/permissions/Build-ID/pinned SQLite
checks pass and all four original fixture hashes match. The theoretical
checklist contains 270 checked and 70 open scoped points, not an Access parity
percentage. Record append/range undo/typed clipboard and complete parity remain open.


## Datasheet previous-record value key (2026-11-03)

Ctrl+apostrophe in the native table view copies the same column from the
immediately preceding displayed model row into the current manual-submit draft.
Checked Currency/date/Boolean conversion and independent SQL NULL preserve the
supported field types. Existing drafts are retained; SQL is not submitted.
First-row/no-current, hidden source/target, generated/read-only and binary
contexts do not mutate. Active text-editor insertion/caret behavior, filtered
multiuser record identity and complete Access keyboard parity remain open.
The main-window test verifies actual key dispatch and unchanged SQL before Save;
the Currency target verifies high exact Currency, Date, True stored as -1, NULL,
AutoNumber/BLOB refusal and actual persistence. A focused two-target run passes
2/2 (1.95 s); final typed case passes 1/1 (0.40 s). No new targets.
Full suite/package baseline remains source 24854bf (48/48, 86.21 s), privately
published as 33981e26e528c0d708ba2a2f3be265aedf3a01ea with identical source tree
caca8e5ef7c1de87a10390a3b14d7272a700c895. This key increment is focused-verified.


## Native editor repeat key and NULL retention (2026-11-03)

Ctrl+apostrophe also works in the native text-cell editor, using checked prior-row
values and manual-submit drafts. A shared change check preserves SQL NULL when
an untouched empty editor closes; explicit user typing/deletion can still save
empty text. Existing main-window/Currency targets pass 2/2 (2.22 s), including
actual editor dispatch and NULL versus empty SQL persistence. See
KEYBOARD_RUNTIME.md. Full suite/package baseline remains 24854bf (48/48);
this small editor increment is focused-verified, with no additional targets.


## Native table F2 navigation and multiline Zoom (2026-11-03)

F2 switches editable table cells between caret input and field navigation;
invalid values remain in the editor. Shift+F2 opens the actual native text Zoom
dialog from a table cell or its active line editor. Cancel preserves values;
unmodified OK preserves NULL and original CRLF. Edited values use manual-submit
drafts. Owner destruction/index invalidation is guarded across the nested
dialog loop. Native form Zoom and exact Windows formatting remain open.
See KEYBOARD_RUNTIME.md for scope and the official shortcut reference.
Full build and existing suite pass 48/48 (86.44 s), Debian payload/permissions/
Build-ID/pinned SQLite checks pass and all four fixture hashes match. GUI-free
core remains 1/1 (2.63 s), with subsequent Qt-only changes. Theoretical checklist:
272 checked / 70 open scoped points; this is not a complete Access parity metric.


## Native form Zoom and Application/DoCmd commands (2026-11-03)

Form TextBox Shift+F2 uses the native multiline Zoom with normal field buffers,
Change, readonly display and cancellable BeforeUpdate. Opening the dialog
suppresses only the originating field's automatic editingFinished commit.
DoCmd.RunCommand acCmdZoomBox=179 targets the active native form control; an
unsuitable focus reports 2046. Application.RunCommand forwards the existing
implemented command subset, including named Command binding, rather than adding
a separate simulated implementation. Unknown commands remain 3251.
Final existing layout target passes 1/1 (14.90 s), GUI-free core 1/1 (1.53 s),
Debian payload/permissions/Build-ID/pinned SQLite checks and all four fixture
hashes pass. Last complete suite remains local 2fcaf98 / API ab5e6cd, 48/48
(86.44 s). This increment is focused-verified; no new targets.
See KEYBOARD_RUNTIME.md for remaining ComboBox/formatting/command-set limits.
The theoretical checklist contains 275 checked / 70 open scoped points; complete
Access 2003 functionality and appearance are still not established.


## Committed-save retry and lossless native numeric editing (2026-11-03)

Two data-loss paths were reproduced on synthetic working data and corrected.
After SQL commits, Save now retires drafts before metadata/query refresh, so a
refresh failure cannot duplicate an insertion on Save retry. The failure message
explicitly distinguishes committed data from the view-refresh failure. Actual
pre-commit failures still roll back and preserve all drafts for correction.
REAL/FLOAT/DOUBLE-backed cells use checked max_digits10 text editing instead of
Qt's default two-decimal spin editor. Merely opening/closing/saving retains the
actual SQL value. Typed-null text EditRole prevents native spin/date defaults
from replacing untouched NULL fields; REAL/Currency/Date NULL cases persist.
See DATASHEET_CLIPBOARD_RUNTIME.md for exact Decimal/formatting/refresh limits.
Full build and existing suite pass 48/48 (88.41 s), Debian payload/permissions/
Build-ID/pinned SQLite checks and all four original fixture hashes pass.
GUI-free core remains 1/1 (1.53 s); subsequent changes are Qt-only. Existing 48
targets only. Theoretical checklist: 277 checked / 70 open scoped points, not
complete Access 2003 parity.


## Native field navigation and retained invalid input (2026-11-03)

Ctrl+End completes bounded query fetching before moving to the final field of
the last record. It refuses the supported cell budget instead of silently
selecting the end of Qt's initial fetch batch. Ctrl+Home retains normal native
first-record/first-field navigation. Active edits are validated first.
Tab, Backtab/Shift+Tab, Return and keypad Enter now validate into the field cache
before Qt closes a native delegate editor. Rejected Currency/type input stays
visible and focused. Valid input follows Qt field navigation while SQL remains
unsubmitted. Shared native editor commit logic also serves explicit Finish/F2.
Mouse/focus-out cancellation, all Access navigation policies and exact keyboard
parity remain open; this increment does not claim to implement them.
Existing main-window/Currency targets pass 2/2 (2.01 s): actual 300-row fetching,
Ctrl+End/Home, rejected keys including both Shift+Tab encodings, valid Tab cache
updates and unchanged SQL until explicit Save. No new targets. Full suite/package
baseline remains local e15ff33 / API c381897, 48/48 (88.41 s); this increment is
focused-verified. Theoretical checklist: 279 checked / 70 open scoped points.


## Native mouse validation and datasheet context menu (2026-11-03)

A normal native editor-close request now validates the field before disposal.
Rejected input stays alive, and mouse selection/focus returns to its original
cell after the transition. Escape's explicit RevertModelCache remains a discard
operation; model destruction/reset can still invalidate an editor. The mouse
loss was reproduced before the fix. Complete inter-window/focus policy parity
remains open.
The table view also has a real native context menu for Copy, checked range Paste,
text Zoom, record/column/all selection. Right-click within an existing selection
preserves it. Paste is enabled only for its supported clean editable target and
text clipboard; keyboard-triggered context menus anchor at the current cell.
Actions call the existing checked native operations; cut/range delete, filter,
all Access context commands and pixel/menu parity are not claimed.
Existing main-window/Currency targets pass 2/2 (2.25 s), including actual rejected
mouse navigation, retained editor/focus/current cell, Escape regressions and
real popup action dispatch into quoted TSV clipboard with dirty Paste disabled.
Last full suite/package remains e15ff33 / API c381897, 48/48 (88.41 s). These
Qt view increments are focused-verified, with the same existing 48 targets.
Theoretical checklist: 281 checked / 70 open scoped points.


## Real table documents in the native window ring (2026-11-03)

Window menu and Ctrl+F6/Ctrl+Shift+F6 now include actual native table tabs between
the database window and native forms/reports. Existing views are activated, with
weak targets rechecked across native validation; source field errors block the
switch, while valid row drafts remain unsubmitted. Mouse tab navigation is
checked before hiding the old page; pure modifier keys do not close an editor.
Native form/table mixed round trips, menu/table direction switches, unavailable
field transitions and preceding numeric/NULL/clipboard behavior are covered in
the existing targets. See WINDOW_CYCLE_RUNTIME.md for scope and remaining MDI/
query/editor-window limitations.
Full build and suite pass 48/48 (92.38 s), Debian payload/permissions/Build-ID/
pinned SQLite checks and all four original fixture hashes pass. GUI-free core
remains 1/1 (1.53 s), with subsequent Qt-only changes. Theoretical checklist:
283 checked / 70 open scoped points, not complete Access 2003 parity.


## Existing table reuse and native SelectObject (2026-11-03)

Opening the same table again activates its existing native page/model instead
of creating a duplicate. Pending row drafts remain intact; the existing model
is not replaced, refreshed or silently submitted. Initial table opening and
activation use the same source validation/modal guards as document switching.
DoCmd.SelectObject acTable,name,False now resolves an actually open table through
the owning production MainWindow and activates its real view. Lookup uses the
existing table metadata with case-insensitive name comparison. The current
source field is validated, activation can fail with 2501, and modal context
returns 2046. A missing/closed table and unsupported owner remain explicit 3251;
exact Access closed-object errors, OpenTable and query-window activation are
not implemented by this increment. The database-window True route remains
separate. Generic table DoCmd commands remain outside the native form/report
command subset; selecting a table must not retarget subsequent commands to the
old caller.
The main-window regression reopens the same dirty table, proves model identity/
one tab/draft retention and unchanged SQL, then runs interpreted SelectObject
with an uppercase table name. Actual focus moves to the table, while Me.Tag
still refers to the caller and closing that caller preserves the table page.
Existing main-window/layout targets pass 2/2 (15.56 s). Last full suite/package
remains 671d997 / private API fea77b0, 48/48 (92.38 s); this UI/runtime increment
is focused-verified. Theoretical checklist: 285 checked / 70 open scoped points.


## Native OpenTable and read-only table mode (2026-11-03)

DoCmd.OpenTable now opens real project datasheets in default/acEdit or acReadOnly
mode, with named/omitted argument binding and a model-level read-only mutation
guard. Valid editable changes persist to SQL; read-only cell/checkbox/insert/
delete attempts leave SQL unchanged. Missing differs from explicit Empty.
Design/preview/append-only and conflicting existing-mode transitions remain
explicit 3251 boundaries. Details and references: TABLE_DOCUMENT_RUNTIME.md.
Full build/suite 48/48 (92.13 s), GUI-free core 1/1 (1.68 s), Debian payload/
permissions/Build-ID/pinned SQLite checks and four fixture hashes pass. Existing
48 targets only. Theoretical checklist: 288 checked / 70 open scoped points.


## Active native table SaveRecord (2026-11-03)

DoCmd/Application RunCommand SaveRecord now uses the actual active table and the
same transactional/selection-restoring helper as its native Save button. Both
specialized interpreter and generic object-call paths are handled; readonly Save
returns 2046, and no old caller form is silently saved. Initialization remains
owned by a never-activated form unless it explicitly switches to a table.
See TABLE_DOCUMENT_RUNTIME.md. Existing trio 3/3 (16.13 s), final initialization
case 1/1 (1.93 s); real SQL persistence and unchanged background drafts verified.
Full suite/package baseline remains e054910 / API 0328119, 48/48 (92.13 s).
Theoretical checklist: 290 checked / 70 open scoped points.


## Native table Close (2026-11-03)

Explicit/default DoCmd.Close handles actual table pages with field validation,
transactional record saving and retained drafts on failure. Hidden active editor
buffers and reentrant context-menu closure are covered. Object design-save
semantics and complete Access error precedence remain open. Canonical details:
TABLE_DOCUMENT_RUNTIME.md. Full desktop 48/48 (91.15 s), GUI-free 1/1 (2.01 s),
Debian checks and four fixture hashes pass. No new test targets.
Theoretical checklist: 292 checked / 70 open scoped points.


## Native table SelectRecord/Copy commands (2026-11-03)

Interpreted DoCmd/Application table commands now select actual rows and copy the
real selection through checked TSV. Readonly access, retained dirty SQL drafts,
and invalid-field refusal with unchanged clipboard are verified. See
TABLE_DOCUMENT_RUNTIME.md. Existing main-window regression 1/1 (1.97 s);
full desktop/package baseline remains 2dd59b2 / API 32cb556, 48/48 (91.15 s).
Theoretical checklist: 294 checked / 70 open scoped points.


## Native table command Zoom (2026-11-03)

acCmdZoomBox uses the actual cell's Qt Zoom, checked cached edits and readonly
viewing. Cancel is normal completion. Whole-row selection preserves the active
column for subsequent commands. Existing main-window/Currency 2/2 (2.47 s);
full/package baseline remains 2dd59b2 / API 32cb556. TABLE_DOCUMENT_RUNTIME.md
records scope and limitations. Theoretical checklist: 295 checked / 70 open.


## Bounded native table record navigation (2026-11-03)

Named/default GoToRecord and native First/Previous/Next/Last use actual view order
and SQL fetches. A sole updated existing current record with primary key is saved
transactionally before navigation; SQL errors retain it for retry. Other dirty
rows, pending inserts/deletes and new-record placeholder behavior remain bounded
unsupported paths. See TABLE_DOCUMENT_RUNTIME.md. Full desktop 48/48 (91.33 s),
GUI-free 1/1 (2.49 s), Debian checks and four fixture hashes pass.
Theoretical checklist: 297 checked / 71 open scoped points.


## Native table SortAscending/SortDescending (2026-11-03)

Actual current-field SQL sorting retains the current record by primary key,
including readonly tables and checked sole-record updates. Unrelated drafts are
refused. Jet/Windows collation parity remains open. TABLE_DOCUMENT_RUNTIME.md
records the scope. Existing main-window 1/1 (2.08 s); last full/package baseline
e4e0900 / API 714a96e, 48/48 (91.33 s).
Theoretical checklist: 298 checked / 71 open scoped points.


## Native table Find UI and commands (2026-11-03)

Actual modeless literal-text Find supports field scope, match/case modes,
Ctrl+F/Shift+F4, Edit menu and DoCmd/Application Find/FindNext. Checked SQL fetches,
sole-record save, protected unrelated drafts and owning-page/dialog lifetime are
covered. View errors reach VBA accurately; binary Copy keeps 3251 and leaves
clipboard/payload untouched. Numeric/date/formatted/wildcard/locale/Replace and
DoCmd.FindRecord remain open. Canonical details: TABLE_FIND_RUNTIME.md.
Full desktop 48/48 (92.69 s), GUI-free 1/1 (2.45 s), Debian checks and four fixture
hashes pass; no new targets. Theoretical checklist: 301 checked / 72 open points.


## Bounded native DoCmd.FindRecord (2026-11-03)

Literal current-text-field SearchAll runs on actual active table/form records,
with named/Missing/Empty arguments, match/case modes, first/next origin and normal
no-match completion. Locked fields are searchable; canceled form validation is
2501 and preserves the draft/caller separation. Explicit table context blocks
old-caller fallback in the database window. General numeric/date/formatted/all-
field/directional/wildcard behavior remains open; see FIND_RECORD_RUNTIME.md.
Full desktop 48/48 (88.05 s), GUI-free core 1/1 (1.99 s), Debian checks and four
fixture hashes pass; no new targets. Checklist: 303 checked / 72 open points.


## Native table/form search directions (2026-11-03)

Up/Down search stops at record boundaries; SearchAll wraps. Table/form dialogs,
FindRecord and direction-preserving FindNext share these real record paths.
Existing trio 3/3 (16.70 s); full/package baseline a71acc9 / API f9367b0,
48/48 (88.05 s). General formatted/numeric/all-field/wildcard/locale search and
full error/selection equivalence remain open. See FIND_RECORD_RUNTIME.md.
Theoretical checklist: 304 checked / 72 open scoped points.


## Typed native table/form Find (2026-11-03)

The real dialogs and FindRecord now share exact Number/Currency/Date/Boolean
matching. Neighboring Double values, scaled Currency units, checked Date storage
canonicalization and NULL-versus-False stay distinct. Individual typed fields use
whole-value comparison and visible input hints; combined scope remains text-only.
Incompatible subtypes, nontext partial matching, all-field API/formatted/locale/
wildcard behavior remain open. Canonical details: TYPED_FIND_RUNTIME.md.
Full desktop 48/48 (97.33 s), GUI-free core 1/1 (1.86 s), Debian checks and four
fixture hashes pass. No new targets. Checklist: 308 checked / 72 open points.

## Scoped native table SendKeys and DoCmd.SendKeys (2026-11-03)

Active table cells and owned Find/Zoom dialogs accept actual Qt key delivery,
with immutable originating project identity and fresh focus resolution per key.
DoCmd.SendKeys supports named/omitted arguments and defaults Wait=False. FIFO
batches continue inside modal handlers; Wait=True and DoEvents drain checked
work. Real Text edit buffers precede Value/SQL commit. Database explorer,
external applications and parentless table/root dialogs remain outside scope.
TABLE_KEYBOARD_RUNTIME.md specifies cancellation, limits and regression scope.
Full Access grammar/timing/layout equivalence is not claimed.

Full rebuilt desktop suite passes 48/48 (93.79 s), GUI-free core 1/1 (2.27 s),
Debian payload/permissions/Build-ID/pinned SQLite checks and all four original
fixture hashes pass. No additional test targets.

## Native SelectAllRecords (2026-11-03)

DoCmd/Application acCmdSelectAllRecords (109) shares checked Ctrl+A table selection,
including full bounded SQL fetching and current-cell preservation without SQL
submission. Existing main-window target passes 1/1 (3.02 s), GUI-free core 1/1
(1.68 s); no new targets. Canonical scope: TABLE_DOCUMENT_RUNTIME.md.
Last full suite/package source is local 934c26d / private API 914cc78:
48/48 (93.79 s), Debian checks and all four fixture hashes passed.
Theoretical checklist: 311 checked / 72 open scoped points.

## Native table Paste command (2026-11-03)

DoCmd/Application acCmdPaste (191) uses the actual bounded Ctrl+V TSV path with
whole-block typed preflight, readonly protection and no implicit SQL submission.
Missing text MIME raises a fresh 2046 instead of reusing an old model error.
Invalid later-cell input retains the clean cache; valid two-row paste changes
only drafts; another paste refuses 2501 and retains them. Existing main-window
target 1/1 (3.12 s), GUI-free core 1/1 (1.63 s); no new targets. Canonical scope:
TABLE_DOCUMENT_RUNTIME.md. Editor-text command routing, paste-over-draft merging,
record append, binary/locale and full Access equivalence remain open.
Theoretical checklist: 312 checked / 72 open scoped points.

Full rebuilt desktop suite passes 48/48 (97.38 s), GUI-free core 1/1 (1.63 s),
Debian payload/permissions/Build-ID/pinned SQLite checks and all four original
fixture hashes pass. This includes the preceding SelectAllRecords increment.

## Native table Requery and record Save keys (2026-11-03)

Active-table DoCmd.Requery and Shift+F9 rerun actual SQL, including externally
changed source rows. A sole current existing update uses checked Save first;
unrelated drafts/insertions/deletions refuse 3251. The reloaded view starts at
the first result row, retaining its field and sort. Named-control and general
record-position/event/error-priority parity remain open. Shift+Return/keypad
Enter now share native table Save, including active cell editors and actual
SQL/primary-key/column restoration. Canonical scopes: TABLE_DOCUMENT_RUNTIME.md
and KEYBOARD_RUNTIME.md. Existing main-window target passes 1/1 (3.12 s), no new
targets. GUI-free core remains 1/1 (1.63 s); this increment changes only Qt paths.
Last full suite/package source is local 375bd82 / private API 89bbcb6:
48/48 (97.38 s), Debian checks and all four original fixture hashes passed.
Theoretical checklist: 314 checked / 72 open scoped points.

Existing main-window/layout pair passes 2/2 (17.72 s), covering the shared
form-host Requery extension and table save shortcuts. No extra targets.

## Native record commands and generated-key selection (2026-11-03)

Table acCmdRecordsGoToNew/28 and acCmdDeleteRecord/223 use actual model insert/
delete drafts with checked Save reaching transactional SQL. Multiple selected
rows refuse before mutation; existing insertion drafts protect against another
new-record command. Readonly mutation stays 2448. SelectAllRecords also works
without a current cell after a model reset. Native SQLite INSERT captures actual
primary keys inside the transaction and restores the saved row/field; unsupported
rowid bookmark cases remain explicit. Staged table deletion is not an immediate
Access commit/confirmation/event claim. Canonical scope: TABLE_DOCUMENT_RUNTIME.md.
Existing main-window/Currency insert pair 2/2 (3.35 s), no new targets.
Theoretical checklist: 317 checked / 72 open scoped points.

Final rebuilt desktop suite passes 48/48 (91.07 s), GUI-free core 1/1 (2.38 s),
Debian payload/permissions/Build-ID/pinned SQLite checks and all four original
fixture hashes pass. Actual ignored INSERT publishes no stale key; first-row
insertion into an empty table and retained generated-key/field selection pass.
No additional test targets. This includes the preceding Requery/Save-key increment.

## Explicit table acNewRec and full source fetch (2026-11-03)

Named/default table GoToRecord acNewRec shares checked record creation, including
activation from another table, empty-source creation, readonly 2448 and existing
insert-draft 3251 protection. Native insertion fetches the remaining existing
source rows before appending, shared with button/Ctrl+Plus; a 300-row partly
loaded source receives draft row 301 without adding SQL rows. Explicit row limit
and no implicit Next-past-last behavior remain. Canonical: TABLE_DOCUMENT_RUNTIME.md.
Existing main-window target passes 1/1 (3.53 s); no new targets. Qt-only change;
last GUI-free core remains 1/1 (2.38 s). Last complete suite/package source is
local a93e2ed / private API cb6a3ea: 48/48 (91.07 s), Debian checks and four
original fixture hashes passed. Theoretical checklist: 319 checked / 72 open.

## Focused table Paste (2026-11-03)

Native acCmdPaste resolves actual owned QLineEdit focus before grid TSV paste,
including cell text buffers and modeless Find queries on readonly sources. Qt
selection/caret insertion does not write the model until field validation or
SQL until Save. Other forms/projects are not fallback receivers. Clipboard
limits and checked readonly/no-text errors remain; multiline/rich command
routing stays open. Canonical scope: TABLE_DOCUMENT_RUNTIME.md. Main-window
target 1/1 (3.17 s), no new targets. The preceding new-record/fetch change also
passes its existing main-window/Currency-field/previous-value trio 3/3 (4.65 s).
Theoretical checklist: 320 checked / 72 open scoped points.

Full rebuilt desktop suite passes 48/48 (91.66 s), GUI-free core 1/1 (1.96 s),
Debian payload/permissions/Build-ID/pinned SQLite checks and all four original
fixture hashes pass. This includes named/default acNewRec, complete source fetch
and actual focused cell/readonly Find text paste. No additional test targets.

## Focused editor commands and clipboard key alternatives (2026-11-03)

Native table acCmdCopy/Cut (190/189) use actual owned QLineEdit selection,
including Find queries on readonly sources. Selected invalid numeric text can
be copied while SelectRecord/Save still reject its commit; missing selection
returns 2046 without replacing the clipboard. acCmdUndo/Redo (292/199) use the
real editor history, preserving buffered field validation/SQL boundaries.
Masked/disabled/readonly text and missing history retain checked guards; grid
Cut/Undo/Redo and general rich/multiline/form dispatch remain explicitly open.
Ctrl+Insert/Shift+Insert now share checked grid Copy/Paste with Ctrl+C/Ctrl+V.
Canonical scopes: TABLE_DOCUMENT_RUNTIME.md, KEYBOARD_RUNTIME.md. Existing
main-window/Currency-field/previous-value trio passes 3/3 (4.25 s), GUI-free core
1/1 (1.71 s), no new targets. Theoretical checklist: 323 checked / 72 open.

Full rebuilt desktop suite passes 48/48 (93.40 s), GUI-free core 1/1 (1.71 s),
Debian payload/permissions/Build-ID/pinned SQLite checks and all four original
fixture hashes pass. No additional test targets; focused Copy/Cut/Undo/Redo and
Ctrl+Insert/Shift+Insert retain the preceding real Qt/SQL regression boundaries.

## Native CurrentProject file metadata (2026-11-03)

CurrentProject/Application.CurrentProject read actual native filename, directory
and full path with per-connection weak identity caching. Caller close preserves
project references; closed connections and accepted replacement invalidate them,
including same-file reopen, while cancelled Unload preserves them. Foreign
preview connections remain separate. Readonly/unsupported properties, non-file
metadata and user variable/function shadowing retain checked boundaries.
CURRENT_PROJECT_RUNTIME.md specifies the native window scope; complete Access
CurrentProject/COM/catalog APIs remain open. Existing core/main-window pair
passes 2/2 (4.57 s), no new targets. Theoretical checklist: 324 checked / 72 open.

Final rebuilt desktop suite passes 48/48 (92.20 s), GUI-free core 1/1 (1.56 s),
Debian payload/permissions/Build-ID/pinned SQLite checks and all four original
fixture hashes pass. The main window fixes the absolute project path before
replacement; standalone relative filenames do not become guessed metadata.
No additional test targets.

## Native project catalogs and report visibility (2026-11-03)

AllForms/AllReports query actual saved native/imported definitions with Long
Count, named/indexed Item, For Each and independent readonly AccessObject
descriptors. IsLoaded follows live hidden/modal/embedded preview state, closed
owners and actual tagged native runner/designer dialogs. Name collisions refuse
5; source versions, ordering and snapshot boundaries stay explicit. Report
Visible now uses the real checked document pipeline; interpreted HideSelf/ShowSelf
continue after visibility changes and activate/deactivate SQL audit rows are
checked. Canonical scopes: PROJECT_CATALOG_RUNTIME.md, FORM_VISIBILITY_RUNTIME.md.
Existing main-window/form-visibility/layout trio 3/3 (18.72 s), no new targets.
Theoretical checklist: 327 checked / 72 open scoped points.

Final rebuilt desktop suite passes 48/48 (97.60 s), GUI-free core 1/1 (1.71 s),
Debian payload/permissions/Build-ID/pinned SQLite checks and all four original
fixture hashes pass. Includes saved-main-schema/temp-shadow isolation, genuine
empty/no-catalog distinction, native catalog lifetimes and report Visible VBA/SQL.
No additional test targets.

## Native CurrentData table/query catalogs (2026-11-03)

CurrentData/Application.CurrentData expose actual saved main-schema user tables
and native/imported query definitions through checked Long Count, Item/named
lookup, For Each and readonly Name/IsLoaded. Internal tables are excluded.
IsLoaded tracks actual table tab membership, including readonly/background tabs
and removed pages, plus real named SQL editor/source/result dialogs. Held
references outlive requesters and invalidate with the existing project generation.
User functions named CurrentData retain resolution. CURRENT_DATA_RUNTIME.md is
the canonical scope; this is metadata, not a blanket query/COM compatibility
claim. Existing core/main-window pair 2/2 (6.16 s), GUI-free core 1/1 (1.51 s);
no new targets. Theoretical checklist: 329 checked / 72 open scoped points.

Full rebuilt desktop suite passes 48/48 (101.66 s), GUI-free core 1/1 (1.51 s),
Debian payload/permissions/Build-ID/pinned SQLite checks and all four original
fixture hashes pass. Native CurrentData UI state and saved metadata included;
no additional test targets.

## Native readonly query documents (2026-11-03)

DoCmd.OpenQuery with explicit acReadOnly/normal view opens actual materialized
SELECT tabs from native/imported saved definitions. Real typed parameter prompts,
Currency results, cancellation, document reuse and readonly flags are checked.
Defaults retain acEdit and explicitly refuse unsupported editing; action/other
views do not execute. Named/default query Close, Ctrl+F4 and actual CurrentData
IsLoaded use independent root-owned tab lifetime after caller self-close.
OPEN_QUERY_RUNTIME.md specifies the bounds; general query commands/updateability
remain open. Targeted main-window 1/1 (3.35 s), no new targets.
Theoretical checklist: 331 checked / 72 open scoped points.

Full rebuilt desktop suite passes 48/48 (90.08 s), GUI-free core 1/1 (1.58 s),
Debian payload/permissions/Build-ID/pinned SQLite checks and all four original
fixture hashes pass. Actual parameterless Close targets the active readonly
query document while preserving its calling form. No additional test targets.

## Readonly query Requery (2026-11-03)

Active readonly query Requery and Shift+F9 reread saved definitions and declared
parameters, execute real SQL and replace results only after success. Errors and
cancel retain the previous model. Closed pages are not resurrected; same-query
recursive SQL refuses 3251 and project close is blocked during SQL execution.
An actual SQLite callback closes/destroys the old page while its replacement is
materialized. Existing main-window target passes 1/1 (3.44 s), no new targets.
OPEN_QUERY_RUNTIME.md defines the bounded scope. Theoretical checklist: 332
checked / 72 open scoped points, not a percentage of complete Access parity.

Full desktop suite passes 48/48 (89.91 s), GUI-free core 1/1 (1.57 s),
Debian payload/permissions/Build-ID/pinned SQLite and all four original fixture
hashes pass. No new test targets.

## Native query navigation, activation and clipboard (2026-11-03)

Readonly saved-query result tabs now handle default and explicit acDataQuery
GoToRecord, bounded record offsets, first/previous/next/last RunCommand and
case-insensitive SelectObject acQuery activation without rerunning SQL. Actual
record/all-cell selection, checked TSV Copy/Ctrl+C/Ctrl+Insert and scoped native
SendKeys operate on the query document, preserving its VBA caller. Tables and
queries share the bounded rectangle/quoting formatter. Unsupported binary,
hidden/disjoint or unrepresentable clipboard values leave existing text intact.
OPEN_QUERY_RUNTIME.md defines the scope and remaining query UI limitations.
Existing full desktop suite 48/48 (91.62 s), GUI-free core 1/1 (1.54 s), Debian
package and all four original fixture SHA-256 checks pass. No new test targets.
Theoretical checklist: 336 checked / 72 open scoped points, not total Access parity.

## Native readonly query search (2026-11-03)

Query Ctrl+F/Shift+F4, Find/FindNext and current-field DoCmd.FindRecord search real
materialized result cells. Proven per-position physical/projection metadata
supports checked text/number/Currency/Date/Boolean search without changing the
existing DAO type-map contract or guessing unknown expressions. Null and search
bounds retain checked behavior. Search-field edit commands share actual focused
text handling. New query tabs select their first result cell synchronously,
fixing immediate post-OpenQuery FindRecord without an event-loop delay.
OPEN_QUERY_RUNTIME.md specifies scope and open search semantics. Full existing
suite 48/48 (90.46 s), GUI-free core 1/1 (1.50 s), Debian package and all four
original fixture hashes pass; no new targets. Checklist: 337 checked / 72 open
scoped points, not an overall Access compatibility percentage.

## Query Edit menu and readonly Zoom (2026-11-03)

Query views now participate in actual Edit-menu Copy/Select All, with readonly
grid Cut/Paste disabled and direct forwarded Copy events using the existing
checked formatter. Selection limits are checked before per-cell index allocation.
VBA acCmdZoomBox and Shift+F2 open real readonly Qt cell Zoom; Close preserves
SQL/model values. Existing main-window/layout pair passes 2/2 (17.70 s), Debian
package and four original fixture SHA-256 checks pass. No new test targets;
previous full suite remains the query-search 48/48 baseline (90.46 s).
OPEN_QUERY_RUNTIME.md states limits. Checklist: 339 checked / 72 open scoped
points, not an overall parity percentage.

## Qualified saved-query lineage and original corpus (2026-11-03)

Three-part query.relation.field references now bind to unique actual output
labels through checked direct projections in stored view definitions, including
ranked TOP wrappers and two input views over one physical table. Physical-only
UNION metadata, ambiguous origins and deeper qualifications are not guessed.
One unchanged original query now executes: current corpus 78/100 (73 ordinary,
five data-parameter cases), still without an Access result oracle. Three remaining
column failures are checked against unchanged saved source definitions (two) or
the original MDB table schema (one). Other remaining cases: nine missing query
dependencies, eight binding/context failures, two guarded joined actions.
QUERY_ORIGINS.md is the canonical scope. Full desktop 48/48 (90.47 s), GUI-free
core 1/1 (1.67 s), Debian package and four fixture SHA-256 checks pass. No new
test targets. Checklist: 340 checked / 72 open scoped points, not total parity.

## Multidimensional Variants and DAO GetRows (2026-11-03)

Variant arrays retain checked per-dimension bounds, multidimensional reads/writes,
value-copy semantics and first-dimension-fastest For Each. ReDim supports up to
60 Variant dimensions and Preserve changes only the last upper bound, with
Option Base applied to omitted lower bounds. Join/Filter remain rank one.
Explicit-positive-NumRows DAO GetRows returns real zero-based field/record
matrices, preserves typed Currency/Date/Boolean/Null and native SQLite binary
cells, and advances to the next unread record only after successful conversion.
DAO_GETROWS_RUNTIME.md and VARIANT_ARRAY_RUNTIME.md specify defaults/typed-array/
provider boundaries. Existing full suite 48/48 (92.36 s), rebuilt GUI-free core
1/1 (1.50 s), Debian package and all four fixture hashes pass; no new targets.
Checklist: 343 checked / 72 open scoped points, not overall Access parity.

## Primitive typed ReDim in Variant containers (2026-11-03)

Existing Variant variables support ReDim As Byte/Integer/Long/Single/Double/
Currency/Date/Boolean/String/Variant with actual typed defaults, checked element
coercion, VarType/TypeName and retained type on omitted As. Non-Preserve clears
values; Preserve retains values/type and existing dimension restrictions.
Fresh typed declaration through ReDim remains 3251 rather than pretending to
implement typed variable declarations. Numeric typed Join/Filter stay outside
the bounded String/Variant path. VARIANT_ARRAY_RUNTIME.md defines the scope.
Full desktop 48/48 (89.81 s), freshly rebuilt GUI-free core 1/1 (1.45 s), Debian
package and four original fixture hashes pass; no new targets. Checklist:
344 checked / 72 open scoped points, not total Access parity.

## Array object references and Byte Erase (2026-11-03)

Set on Variant array cells stores actual native object/Nothing references.
Container copies retain independent entries with shared object identity; existing
Variant containers support ReDim As Object with Nothing defaults, metadata and
Preserve. Native DAO objects retrieved from real arrays execute SQL and cursor
methods. Invalid scalar/object index writes preserve entries, and implicit
Object-array default-member writes remain explicit 3251. Byte arrays held in
Variant variables now Erase to Empty with supported ByRef writeback/copy isolation.
VARIANT_ARRAY_RUNTIME.md specifies remaining declaration/class/COM limits.
Full desktop suite 48/48 (92.45 s), fresh GUI-free core 1/1 (1.79 s), Debian
package and all four original fixture hashes pass; no new targets. Checklist:
346 checked / 72 open scoped points, not overall parity.


## Bounded editable query documents (2026-11-03)

Default/acEdit OpenQuery now opens actual keyed single-table SELECT-star query
datasheets, using the typed transactional table model with real SQL save,
validated record close and unchanged-definition Requery. Query identity and
independent lifetime are retained. Competing writable native source tabs and
changed source structure are guarded; no general Jet concurrency claim.
OPEN_QUERY_RUNTIME.md defines the exact restrictions. Filtered/projected/joined
updateability, action-query UI and full Access query designer remain open.


## Filtered editable query subset (2026-11-03)

Single-table SELECT-star editing additionally accepts bounded compiled WHERE
predicates through the real SQL table model. Actual saves refresh membership;
excluded rows are not edited accidentally. Subqueries/functions/other clauses,
parameters/projections/joins remain outside this path. Structurally nullable
primary keys refuse; INTEGER rowid aliases are distinguished from DESC indexed
keys. OPEN_QUERY_RUNTIME.md documents the measured native behavior and limits.
Full 48/48 (89.83 s), GUI-free 1/1 (1.49 s), Debian package and four hashes pass.


## Declared dynamic VBA arrays and parameters (2026-11-03)

Local/module dynamic declarations support native typed unallocated state,
checked ReDim/Preserve, copy isolation and Erase without losing declaration
type. Module values persist across procedures. Required matching ByRef array
parameters perform actual caller-visible resize/edit/Erase; whole-array a()
expressions preserve reference identity. Direct declared-array ByVal and plain
Variant ByRef bindings remain guarded; explicit Variant copies can be inspected.
VARIANT_ARRAY_RUNTIME.md defines fixed/class/Optional/COM/enumeration boundaries.
Full desktop 48/48 (93.65 s), rebuilt GUI-free 1/1 (1.85 s), Debian package and
four original fixture hashes pass; no new test targets or full VBA claim.


## Native VBA file metadata (2026-11-04)

FileLen and FileDateTime read actual native Linux file metadata, with named/
VBA-qualified calls, UTF-8 absolute/relative paths, checked Long size overflow
and typed local-time Date results. No contents are read or written by these
functions. FILE_METADATA_RUNTIME.md documents Windows path/handle/locale/SQL
boundaries and native error mapping. Full desktop 48/48 (93.70 s), rebuilt
GUI-free 1/1 (1.65 s), Debian package and four original fixture hashes pass.


## Native SQL/DAO file metadata (2026-11-04)

FileLen/FileDateTime execute actual C++ filesystem functions from parameterized
SQL, exposing INTEGER4/DATETIME projections and real Long/Date DAO Values.
DAO Field.Type recognizes proven DATE/DATETIME as dbDate. Callbacks remain
non-deterministic; pure rewrite allowlists are not widened. Checked filesystem
errors propagate without private path text. FILE_METADATA_RUNTIME.md defines
native path/handle/concurrency and unsupported rewrite boundaries.
Full desktop 48/48 (94.77 s), GUI-free 1/1 (1.89 s; runtime unchanged), rebuilt
Debian package and four original fixture hashes pass; no new test targets.


## Literal-bound fixed VBA arrays (2026-11-04)

Local/module fixed arrays allocate real typed dimensions from decimal literal
bounds, respecting Option Base. Erase resets defaults without deallocation;
Variant snapshots remain independent. Fixed binding survives matching ByRef
array calls and protects against ReDim/shape replacement. Whole fixed-array
assignment remains an explicit 3251 boundary. VARIANT_ARRAY_RUNTIME.md specifies
constant/class/Static/alias/enumeration/error-precedence limits. Full desktop
48/48 (91.05 s), rebuilt GUI-free 1/1 (1.76 s), Debian package and four original
fixture hashes pass; no new test targets or complete VBA declaration claim.


## Native scalar constants and constant dimensions (2026-11-04)

Local/module Const declarations bind immutable checked scalar values, compose
pure native expressions and supply real folded fixed-array dimensions. Supported
ByRef calls receive temporary constant values, and direct writes are rejected
before the procedure body. Windows-target newline/control-character constants
are available. VBA_CONSTANTS_RUNTIME.md specifies scalar/type/linking/class/
conditional/default-value boundaries; full compiler parity is not claimed.
Full desktop 48/48 (95.27 s), rebuilt GUI-free 1/1 (2.11 s), Debian package and
four original fixture hashes pass; no new CTest targets.


## Folded optional defaults (2026-11-04)

Optional scalar defaults accept pure expressions and prior same-module constants,
retained as immutable parameter values. Omitted calls copy/coerce defaults;
explicit Empty, named overrides, Missing/Null and Collection Nothing retain
separate meanings. Independent ByRef defaults do not leak state between calls.
OPTIONAL_ARGUMENT_RUNTIME.md and VBA_CONSTANTS_RUNTIME.md define linking/object/
type/error boundaries. Full desktop 48/48 (92.51 s), rebuilt GUI-free 1/1 (1.84 s),
Debian package and four original hashes pass. Final test-only UTC construction
uses Qt-6.4-compatible API and the query-model target passes separately (0.04 s)
on the existing Qt 6.8.2 environment; no actual Qt 6.4 matrix is claimed.


## Source-only native VBA translation cache (2026-11-04)

Successful native procedure plans reuse exact-source/name/parser-context keys
within a bounded per-thread LRU. Mutable module/host/argument/DAO/widget state
binds per call and is not cached. Source edits get different plans; oversized
plans run uncached. VBA_PROCEDURE_CACHE.md documents budgets, retained metadata
and the limited synthetic performance comparison. This is not a complete VBA
compiler or execution-result cache. Full desktop 48/48 (81.60 s), rebuilt GUI-free
1/1 (1.18 s), Debian package and four original fixture hashes pass.


## Native Dictionary and indexed object/DAO writes (2026-11-04)

Local CreateObject/New Scripting.Dictionary uses actual ordered native storage,
checked primitive keys, bounded comparison modes, real snapshots and guarded
enumeration lifetime. Default indexed object writes retain object identity and
array value isolation. DAO Recordset/Fields indexed writes use actual buffered
Field.Value and keyed SQL Update. SCRIPTING_DICTIONARY_RUNTIME.md records
unsupported key/coercion/COM/module-binding boundaries.


Validation (2026-11-04): full desktop 48/48 (77.17 s), freshly rebuilt
GUI-free core 1/1 (1.20 s), targeted core/DAO 2/2 (1.44 s), Debian
payload/permissions/Build-ID/pinned SQLite and all four original fixture
hashes pass. Scoped checklist: 368 checked / 72 open; not total Access parity.


## Persistent scalar module Object fields (2026-11-04)

Private/Public/Global scalar Object fields initialize to Nothing and retain real
native references in per-host module state. Repeated Dictionary calls share
actual storage; separate hosts remain independent; clearing both aliases with
Set ... = Nothing releases bindings. Non-object native binding refuses 13.
Early-bound Scripting/DAO module declarations and automatic New remain open.
Existing core target covers repeated calls, identity, visibility, clear/recreate
and host isolation. Full desktop 48/48 (78.39 s), freshly rebuilt GUI-free 1/1
(1.09 s), targeted core 1/1 (1.05 s), Debian payload/permissions/Build-ID/pinned
SQLite and all four original fixture hashes pass. Scoped checklist 369 checked
/ 72 open; not total Access 2003 parity.


## Checked native Object parameters and optional Nothing (2026-11-04)

As Object binds actual references/Nothing before execution. ByVal shares object
mutation while replacement stays local; supported ByRef writes replacement back.
Optional Object Nothing slots are independent per omitted call. Set and object
collection loops use checked assignment. OPTIONAL_ARGUMENT_RUNTIME.md records
remaining default-member/compiler/early-bound boundaries. Existing core target
extended only. Full desktop 48/48 (77.73 s), freshly rebuilt GUI-free 1/1
(1.24 s), targeted core 1/1 (1.23 s), Debian payload/permissions/Build-ID/pinned
SQLite and all four original hashes pass. Scoped checklist 370 checked / 72
open; not complete Access 2003 parity.


## Native DAO identity and qualified object bindings (2026-11-04)

DAO_OBJECT_BINDINGS.md defines actual native DAO identities, qualified scalar
module/local/parameter/result binding and independently omitted Nothing
defaults. Actual SQL cursor state persists across procedure calls; wrong Set
leaves the original typed Recordset intact. Existing owner/connection/clone/
transaction boundaries and the bounded form Recordset adapter remain active.
Full desktop 48/48 (77.79 s), freshly rebuilt GUI-free core 1/1 (1.23 s),
targeted core/DAO 2/2 (1.59 s), Debian payload/permissions/Build-ID/pinned SQLite
and all four original fixture hashes pass. No extra CTest target. Scoped
checklist 371 checked / 72 open, not complete Access 2003 parity.


## Explicit persistent procedure-local Static declarations (2026-11-04)

STATIC_LOCAL_RUNTIME.md specifies actual per-host/procedure/source static
scalar/object/array state, nested-call alias refresh and current initialization
limits. Existing core cases cover persistence, native Dictionary/fixed/dynamic
arrays, independent scopes and recursion. Static procedure modifiers now refuse
incomplete implicit-local behavior instead of silently executing transient
locals. Full desktop 48/48 (79.78 s), freshly rebuilt GUI-free 1/1 (1.45 s),
targeted core 1/1 (1.26 s), Debian payload/permissions/Build-ID/pinned SQLite and
all four original hashes pass. No new CTest target. Scoped checklist 372 checked
/ 72 open, not complete Access 2003 parity.


## Native whole-text-field datasheet replacement (2026-11-04)

Ctrl+H, native edit menu and acCmdReplace/29 route to an actual modeless
whole-field text replacement path. It edits the real manual-submit model,
uses keyed transactional Save and retains drafts on SQL failure. Captured
primary-key/source/field checks cancel reentrant record changes before mutation.
TABLE_REPLACE_RUNTIME.md defines search/text/readonly/new-record limits and
remaining partial/bulk/form/undo/oracle work. Existing main-window target only.
Full desktop 48/48 (77.04 s), rebuilt GUI-free 1/1 (1.36 s), targeted GUI/SQL
1/1 (2.15 s), Debian payload/permissions/Build-ID/pinned SQLite and all four
original fixture hashes pass. Scoped checklist 373 checked / 72 open, not
complete Access 2003 parity.


## Current-record pending-edit discard with Esc (2026-11-04)

Native datasheet Esc/context action uses actual per-row manual-submit revert
without saving or discarding other drafts. First Esc inside the cell editor
retains prior cached drafts; subsequent record Esc discards that row, including
a pending insertion. Committed SQL remains unchanged. TABLE_RECORD_UNDO_RUNTIME.md
defines remaining history/deletion/form/oracle boundaries. Existing main-window
regression only. Full desktop 48/48 (79.77 s), GUI-free build confirmed and core
1/1 (1.74 s), targeted GUI/SQL 1/1 (2.10 s), Debian payload/permissions/Build-ID/
pinned SQLite and all four original fixture hashes pass. Scoped checklist
374 checked / 72 open; not complete Access 2003 parity.


## Qualified native Scripting.Dictionary bindings (2026-11-04)

Native Dictionary local/module/parameter/function-result declarations initialize
to Nothing and preserve checked actual references. Typed ByRef mutation,
qualified returns, optional Nothing slots and wrong Collection replacement
refusal execute in the existing core target. Generic type-library/COM loading,
unqualified type lookup, As New and typed Dictionary arrays remain open.
SCRIPTING_DICTIONARY_RUNTIME.md specifies the bounded native path. Full desktop
48/48 (77.95 s), freshly rebuilt GUI-free 1/1 (1.26 s), targeted core 1/1
(1.20 s), Debian payload/permissions/Build-ID/pinned SQLite and all four original
hashes pass. Scoped checklist 375 checked / 72 open, not complete Access parity.


## Native scalar object default-member value assignments (2026-11-04)

Declared object value/Let assignments invoke actual Field/Parameter/Control
Value setters and preserve their reference binding; Set continues reference
replacement. Real DAO edit/SQL and bound form/save paths execute in existing
cases, including explicit Let syntax (already parsed previously).
OBJECT_DEFAULT_ASSIGNMENT_RUNTIME.md records Nothing/no-default and remaining
Variant/coercion/oracle boundaries. Full desktop 48/48 (79.37 s), freshly rebuilt
GUI-free core 1/1 (1.49 s), targeted core/DAO/forms 3/3 (10.76 s), Debian payload/
permissions/Build-ID/pinned SQLite and four original fixture hashes pass. No new
CTest target. Scoped checklist 376 checked / 72 open, not complete Access parity.


## Actual Byte Variant For Each execution (2026-11-04)

Packed and multidimensional Byte arrays enumerate true Byte Variants in native
linear order. Opaque allocation identity preserves value copies/live supported
ByRef element updates and detects source replacement without reporting Byte()
as Object. VARIANT_ARRAY_RUNTIME.md documents ownership, source and alias
boundaries. Existing core target checks 0/128/255, empty arrays, typed order,
copy isolation and ReDim refusal. Full desktop 48/48 (78.05 s), freshly rebuilt
GUI-free 1/1 (1.26 s), targeted core 1/1 (1.22 s), Debian payload/permissions/
Build-ID/pinned SQLite and all four original fixture hashes pass. No new CTest
target. Scoped checklist 377 checked / 72 open, not complete Access parity.


## Source-mapped multi-target Erase (2026-11-04)

Erase lists retain one native statement/source line while applying actual
dynamic deallocation, fixed defaults and Variant release in order. Constant
targets refuse before body execution; a runtime error stops the remaining
targets and Resume Next proceeds to the next statement. Existing core target
checks mixed array storage, dimensions/copies, middle-target failure and
constant side-effect refusal. VARIANT_ARRAY_RUNTIME.md defines native bounds.
Full desktop 48/48 (78.60 s), freshly rebuilt GUI-free 1/1 (1.25 s), targeted
core 1/1 (1.27 s), Debian payload/permissions/Build-ID/pinned SQLite and all four
original fixture hashes pass. Scoped checklist 378 checked / 72 open; no new
CTest target or claim of complete Access parity.


## Ordered multi-target ReDim/Preserve (2026-11-04)

ReDim lists compile source-only per-target metadata into one native instruction.
Successful module/ByRef binding is published before later dimensions, allowing
an actual bounds function to observe the preceding allocation. Preserve/type/
copy/default behavior uses the existing native arrays; runtime failure retains
earlier completed operations and skips remaining targets under Resume Next.
Cache memory accounting includes descriptors. VARIANT_ARRAY_RUNTIME.md specifies
remaining property/compiler/alias/oracle boundaries. Full desktop 48/48
(78.63 s), freshly rebuilt GUI-free 1/1 (1.47 s), targeted core 1/1 (1.31 s),
Debian payload/permissions/Build-ID/pinned SQLite and four original fixture
hashes pass. No new CTest target. Scoped checklist 379 checked / 72 open, not
complete Access 2003 parity.


## Validated Up/Down datasheet record movement (2026-11-04)

Bare vertical record keys reuse actual native current-record Save/navigation,
retaining the current column and stopping on SQL failure with drafts/selection
intact. TABLE_KEYBOARD_RUNTIME.md records destination, edit, unrelated-draft,
addition-row and general movement/oracle boundaries. Existing main-window target
checks actual keys, saved SQL, UNIQUE failure, Esc and retry. Full desktop 48/48
(77.41 s), GUI-free build confirmed and core 1/1 (1.79 s), targeted GUI/SQL
1/1 (2.30 s), Debian payload/permissions/Build-ID/pinned SQLite and all four
original hashes pass. Scoped checklist 380 checked / 72 open; no extra CTest
target or complete Access 2003 parity claim.


## Validated adjacent-row Tab/Backtab movement (2026-11-04)

View/delegate cursor requests to adjacent existing rows reuse checked native
Save/navigation, restore the proposed logical column after Save and preserve
current drafts on failure. Existing GUI cases execute view Tab, Backtab, editor
Tab, actual SQL persistence and UNIQUE failure/Esc recovery. Endpoint/addition/
hidden-row/general movement and full Access oracle limits remain in
TABLE_KEYBOARD_RUNTIME.md. Full desktop 48/48 (78.54 s), GUI-free build confirmed
and core 1/1 (2.07 s), targeted GUI/SQL 1/1 (2.48 s), Debian payload/permissions/
Build-ID/pinned SQLite and all four original hashes pass. Scoped checklist
381 checked / 72 open, no extra CTest target or complete Access parity claim.


## Checked native Ctrl+Home/Ctrl+End record boundaries (2026-11-04)

Typed native record movement requests route first/last keys through checked
Save/context/fetch behavior and actual visible-field header order. Current-first
and known fully fetched current-last field moves preserve pending edits without
SQL Save. Existing GUI checks positions, SQL persistence, same-record drafts
and UNIQUE failure. TABLE_KEYBOARD_RUNTIME.md records driver/editor/selection/
insert/delete/oracle limits. Full desktop 48/48 (80.50 s), GUI-free build
confirmed and core 1/1 (2.31 s), targeted GUI/SQL 1/1 (2.12 s), Debian payload/
permissions/Build-ID/pinned SQLite and four original hashes pass. Scoped checklist
382 checked / 72 open; no extra CTest target or complete Access parity claim.


## Validated loaded-page record navigation (2026-11-04)

PageUp/PageDown pass actual loaded-viewport record distances to checked native
Save/navigation. Existing 60-row GUI/SQL execution verifies a multi-record step,
Save, UNIQUE failure retaining current row/cache/SQL, Esc/retry and a zero-step
first-page draft without Save. TABLE_KEYBOARD_RUNTIME.md records lazy-page,
context/editor/selection/geometry and Access oracle limits. Full desktop 48/48
(78.15 s), GUI-free build confirmed and core 1/1 (1.94 s), targeted GUI/SQL
1/1 (2.30 s), Debian payload/permissions/Build-ID/pinned SQLite and all four
original fixture hashes pass. Scoped checklist 383 checked / 72 open, no extra
CTest target or complete Access 2003 parity claim.


## Native multi-counter Next source groups (2026-11-04)

Next lists lower to matched native loop steps and one source-line/error group.
Zero/Exit For, mixed For Each, Do and inline merging retain actual scope flow;
Resume Next skips remaining terminators and abandoned iteration state is
released. NEXT_LIST_RUNTIME.md records compiler/jump/error/oracle boundaries.
Existing core cases cover values, counters, empty loops, mutation after error
resumption and retained invalid-order/blocked-label diagnostics. Full desktop
48/48 (78.61 s), freshly rebuilt GUI-free 1/1 (1.39 s), targeted core 1/1
(1.31 s), Debian payload/permissions/Build-ID/pinned SQLite and all four original
fixture hashes pass. Scoped checklist 384 checked / 72 open; no extra CTest
target or complete Access 2003 parity claim.


## Fester Statuskatalog und eigene VBA-Fehler (2026-11-04)

[FIXED_WORK_CATALOG.md](FIXED_WORK_CATALOG.md) fixiert F001–F456 in fünf
Kategorien und 14 fachlichen Unterkategorien. Neue Details gehören zu bestehenden
IDs; die Gesamtzahl wächst nicht mehr. Die Verhältniszahl bestätigter
Teilumfänge ist keine Gesamtparität mit Access 2003.

Unter F051 trägt natives Err.Raise eigene signierte Long-Fehlernummern samt
Quelle/Beschreibung/Hilfeangaben durch verschachtelte native Aufrufe und
Ereignisfunktionen. Benannte/ausgelassene optionale Argumente, beschreibbare
Err-Felder sowie Clear/Resume arbeiten mit echtem Laufzeitzustand. Kein
Windows-Hilfeaufruf und keine vollständige Access-Fehlertext-/HRESULT-/Debugger-
Gleichheit werden behauptet. Details stehen in DAO_RUNTIME.md.


Unter F042 sind Enum-Mitglieder als native unveränderliche Long-Konstanten
angebunden: automatische/explizite Werte, Aliase, negative Werte, 32-Bit-Grenzen,
Public/Private und Enum.Member-Zugriffe. Übersetzung und Modulzustand nutzen
dieselben geordneten Deklarationen; lokale/Instanz-Wurzeln behalten Vorrang.
Variablen/Arrays/Parameter As Enum und vollständige Typbibliotheks-/Klassensicht-
barkeit bleiben offen. Unter F051 nutzt Error expression eine eigene native
Anweisung; ein lokales Err ersetzt diesen Fehlerpfad nicht. Unary +/- erhält
Null (F041). Diese Details gehören zu bestehenden IDs; Gesamtzahl bleibt 456.


GoSub/Return (F040/F051) führt native prozedurlokale Unterprogramme mit
gemeinsamen Variablen, geschachtelten Rücksprüngen und lebenden Aufrufer-Scopes
aus. Label-Handler können eigene GoSubs ausführen und den Fehlerzustand mit
Resume/Resume 0/Resume Next wiederherstellen oder per Resume label verlassen.
Return ohne GoSub meldet Fehler 3. Direkte Handler-Rückkehr in einen
suspendierten Fehler-GoSub (3251) und Sprünge in nicht aktive Blöcke bleiben
offen; keine vollständige VBA-Sprung-/Debugger-Parität.


Unter F040 erlauben Labeldefinitionen innerhalb von Blöcken GoTo zum selben
aktiven Scope oder einem Vorfahren. Nur verlassene If/Select-/For-/With-Scopes
werden freigegeben; aktive Enumeration und GoSub-Aufrufer bleiben erhalten.
Inline-Labelnamen werden beim Zusammenführen behalten und global gebunden.
Numerische Labels 0–65529 funktionieren mit/ohne Doppelpunkt; kanonische
Zielnamen unterscheiden sich nicht durch führende Nullen. On Error GoTo 0
und Resume 0 behalten ihre Sonderbedeutung. Handler-/GoSub-/Resume-Labelziele
müssen weiterhin außerhalb von Blöcken liegen; Erl-/Debugger-Gleichheit bleibt
offen. Die feste Katalogzahl bleibt 456.


Unter F043 binden native Aufrufe Arrayelemente benannter Wurzeln als echte
ByRef-Adressen: Indexauswertung einmalig, typisierte sofortige Schreibwirkung,
gleiche Elementaliase, Weiterleitung und benannte Argumente. ByVal behält den
vor späteren Argumenten ausgewerteten Wert; alte Argumentkopien überschreiben
keine lebenden Quellen. Aktive Bindungen sperren Quellersetzung/ReDim/Erase.
Native Alias-Frames veröffentlichen Änderungen und lösen veraltete Objekt-
Snapshots auf. Compound-/ParamArray-Adressen und verschachtelte Arraywurzeln
über bestehende Zellaliase bleiben mit 3251 gesperrt. Vollständige COM-/Klassen-/
ByRef-Fehlergleichheit bleibt offen. Details: VARIANT_ARRAY_RUNTIME.md.


## Stored native forms use the document runtime (2026-11-04)

Under existing F080, owned `_la_objects` form definitions now open through
AccessFormPreview, sharing actual SQL binding, VBA instance state, document
registration, DoCmd, saving and close/Unload/Quit handling with imported forms.
The former modal dialog and separate manual INSERT/UPDATE path are removed.
DoCmd.OpenForm resolves owned definitions and rejects ambiguous imported/native
names. Generated default New/Save handlers call actual DoCmd commands; they are
runtime additions and never replace the saved user VBA or original source bytes.
The existing editor persists field order and editable captions, restores both,
and closes a loaded matching document before editing, respecting close refusal.

The existing GUI target exercises real designer persistence, rendered order and
captions, actual SQL saves, persistent VBA state, opening a child then closing
its caller, IsLoaded and reopening the designer without altering source.
Targeted mainwindow_documents passes (2.24 s). Full validation follows.
No new CTest targets or catalog IDs. Fixed 456 remains. This is a bounded native
form editor, not full Access visual design, arbitrary layout editing, report
parity or an original Access result oracle. Owned subform loading remains open.

Completion: desktop 48/48 (82.61 s), GUI-free build and core 1/1 (1.52 s),
Debian payload/permissions/Build-ID/pinned SQLite and four original fixture
hashes pass. Fixed inventory remains 384/456 scoped entries, 72 open.


### Follow-up: owned embedded forms and catalog resolution

Owned native forms also load as real unlinked embedded subforms. They share the
same SQL/VBA instance and parent close traversal as imported subforms, while
LinkMasterFields/LinkChildFields remain explicitly unsupported. Native catalog
and definition queries now qualify main storage, matching project metadata and
preventing temporary same-name SQL tables from shadowing saved definitions.
Existing GUI exercises native-source lookup under a temporary shadow, an actual
embedded form SQL save through Dirty=False, IsLoaded and parent-close cleanup.
Targeted mainwindow_documents and access_2003_layouts pass 2/2 (11.47 s).
The preceding complete desktop run passed 48/48 (82.61 s); it predates this
follow-up. No new test target, catalog entry or full-parity claim.


### Owned form field geometry

The existing native form editor now persists per-field left/top/width/height in
Twips (`field_layout`). The runtime adapter validates integral bounded geometry,
uses the actual dimensions and expands the canvas/button area to contain fields.
The property row follows selection, checked-field order and drag/drop moves;
default positions count included fields rather than unchecked list entries.
Existing GUI verifies saved metadata, actual rendered position/size and restored
properties alongside unchanged user VBA. GUI/layout 2/2 (11.23 s) passed before
the final property-row refresh adjustment; complete validation follows.
This remains a bounded form editor: arbitrary controls/bands, full visual design,
master/child links and Access rendering/event oracle remain open. Fixed 456.

Completion of owned subforms and field geometry: full desktop 48/48 (78.56 s),
Debian package payload/permissions/Build-ID/pinned SQLite and all four original
fixture hashes pass. The GUI-free build/core last passed before these GUI-only
changes (1.52 s). Fixed 456 IDs, 384 confirmed scoped entries and 72 open remain.


Native form TextBoxes execute Click and retained OnMouseDown/OnMouseUp with
real typed button/modifier/Twip arguments. Existing GUI covers event order and
VBA cursor changes without implicit SQL saves. Full mouse/focus/scale/oracle
parity remains open; see KEYBOARD_RUNTIME.md. Fixed 456 remains unchanged.

Completion: desktop 48/48 (79.27 s), freshly rebuilt GUI-free core 1/1
(1.56 s), Debian package payload/permissions/Build-ID/pinned SQLite and all four
original fixture hashes pass. Fixed 456 entries and 48 CTest targets unchanged.


### Follow-up: native button mouse dispatch

Command/Toggle buttons and standalone CheckBox/OptionButton controls now use the
same typed MouseDown/MouseUp binding as TextBoxes. MouseUp runs before native
button activation; handler failures clear the pressed state and stop that click.
Native button double-click dispatch avoids a fabricated second MouseDown.
Option-group member mouse dispatch is excluded pending complete group semantics.
Owned native form action buttons retain declared MouseDown/MouseUp procedures
alongside their actual generated or user Click procedure.

The existing stored-form GUI case writes a field in MouseUp and verifies the
subsequent real Save action persists that exact value. Existing AllowEdits GUI
writes an actual SQL down/up/click trace while a read-only checkbox leaves its
bound record unchanged. Final targeted GUI/layout 2/2 (11.23 s) pass, with no
new CTest target or catalog ID. Last full desktop 48/48 (79.27 s) and fresh
GUI-free core (1.56 s) precede this GUI-only follow-up. Full focus/group/move/
scale/Windows event parity remains open. Fixed 456 entries remain unchanged.

Button follow-up packaging: Debian payload/permissions/Build-ID/pinned SQLite
and all four original fixture hashes pass. Full desktop suite last passed on
the preceding TextBox commit; final button changes passed targeted GUI/layout
2/2 (11.23 s). The test target inventory remains 48, fixed work catalog 456.


## Computed native On branches (2026-11-04)

Under existing F040/F045, On expression GoTo and On expression GoSub now execute
through native jump/return machinery. Selectors evaluate once and Let-coerce to
Integer before range checks (banker rounding, Null/type/Integer-overflow errors).
Zero or a selector beyond the destination list falls through; a coerced value
outside 0-255 raises 5. All named/numeric destinations bind before execution,
including unselected entries. Missing/invalid destinations cannot run prior body
effects. Numeric labels retain canonical leading-zero handling.

Computed GoTo carries per-destination lexical unwind counts; selected ancestor
jumps release only exited condition/With/For scopes and real enumeration leases.
Computed GoSub shares the bounded native return stack, caller scope floor,
Return and handled Resume snapshots. Inline wrappers bind targets only after
merging into the final procedure. Existing inactive/sibling inbound-jump and
nested GoSub/error/Resume destination restrictions remain explicit. The source-
only LRU plan cost now includes destination vectors and label storage.

Existing core combines counted/fractional selectors, inline nested GoSubs,
resume/retry after selector conversion, fall-through, 5/6/94 errors, numeric
labels and real Collection lease release. Core/DAO 2/2 (1.94 s) passed before
changing the existing DAO clone escape to a computed GoTo; that change and full
validation follow. No new CTest target or fixed-catalog ID. Full parser/classes/
control flow/Access oracle parity remains open.

References: [On...GoSub/GoTo](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/ongosub-ongoto-statements)
and Microsoft [MS-VBAL](https://learn.microsoft.com/en-us/openspecs/microsoft_general_purpose_programming_languages/ms-vbal/),
computed branch runtime semantics (Integer Let coercion and range error 5).

Completion: desktop 48/48 (81.46 s), freshly rebuilt GUI-free core 1/1
(1.65 s), Debian payload/permissions/Build-ID/pinned SQLite and all four
original fixture hashes pass. The existing real DAO With/clone case now exits
through a computed GoTo and successfully starts its real transaction afterward.
This full run also covers the preceding native button mouse follow-up. Fixed
456 work entries and 48 CTest targets remain unchanged.


## Explicit native TypeOf class identity (2026-11-04)

Under existing F042/F043/F122, TypeOf reference Is class now queries explicit
native class/interface metadata, without retrieving scalar defaults or guessing
identity from displayed TypeName strings. Native Collection, Dictionary, Err,
DAO objects and supported Access form/report/control/application/collection
wrappers expose bounded identity. Control and document type metadata is captured
independently of mutable field values and weak document availability.

The operator keeps its operand in object-reference context, supports existing
property/function/index/bang/With reference paths, produces Boolean and handles
Not/And through native expression precedence. A non-object Variant value compares
False; a Nothing object reference raises VBA error 91. Unknown class/reference
bindings, unsupported literal/value operand forms and calculated-expression
TypeOf remain explicit 3251 rather than fabricated matches. Full declared-type/
static expression classification, UDTs, user class inheritance, reference order,
COM interfaces and Access oracle remain open.

Existing core verifies real Collection/Dictionary references, indexed object
items, a counted object-returning function, Variant scalar, Nothing/Err and
unknown-class rejection before procedure effects. Existing DAO inspects actual
Recordset/Fields/Field references without dereferencing scalar field defaults.
Existing GUI inspects Form, TextBox, CommandButton, Controls and Report/Reports
from real native event procedures. Targeted core/DAO/GUI/layout 4/4 (13.14 s)
passed. The existing DAO case was then strengthened to inspect a temporary real
clone and begin a transaction immediately afterward, checking clone release;
this final addition and complete validation follow. Fixed 456/48 unchanged.

Primary reference: Microsoft [MS-VBAL](https://learn.microsoft.com/en-us/openspecs/microsoft_general_purpose_programming_languages/ms-vbal/),
section 5.6.7 (TypeOf...Is). This is native type identity, not a COM adapter.

TypeOf validation completion: the complete desktop run passed 47/48 (79.35 s).
The added DAO probe had already released its temporary clone and successfully
begun a transaction, then reused a cursor invalidated by the existing Rollback
contract (3420). Corrected that probe to commit its read-only transaction so
subsequent existing cursor checks remain valid; the corrected DAO target passes
1/1 (0.44 s). Runtime code did not change after the complete run: all 48 existing
targets are therefore verified across those runs, without repeating the other 47.
Fresh GUI-free build/core 1/1 (1.69 s), Debian payload/permissions/Build-ID/pinned
SQLite and four original fixture hashes pass. Last single clean full 48/48 run
was the preceding computed-branch commit (81.46 s). No new CTest target or fixed
catalog ID. Full static/UDT/user-class/COM/Access type binding remains open.


## Typed native Access object slots (2026-11-04)

Under existing F050/F043/F122, a shared closed native object vocabulary now
normalizes Access class aliases for locals, module fields, parameters and
function results. Native Form/Report/Control, supported concrete controls,
Application and Forms/Reports/Controls wrappers initialize as Nothing and use
explicit class/interface metadata for assignment/conversion. Incorrect native
classes fail with 13 before replacing a valid slot or entering the callee body.
Optional native Access object defaults support Nothing. Equal qualified/short
names share a canonical ByRef declaration identity. VBA.Collection declarations
normalize to the existing Collection type. Unknown parameter classes remain
unsupported before body execution; broadening signature syntax does not admit
unimplemented class bindings.

TypeOf uses the same vocabulary. Typed control Let/default assignments continue
through actual native field setters, without treating the object reference as a
scalar typed variable. Existing GUI exercises typed function results, module
Form/Control state, optional Nothing, canonical ByRef aliases, rejected dynamic
class assignments/arguments, a real TextBox edit followed by record Undo, and
actual Report references. It explicitly checks successful Form_Load completion.
Core/DAO/mainwindow/layout targeted 4/4 (13.84 s) pass. Complete validation follows.

These are scalar native object slots, not full static interface-member binding,
context-sensitive direct property actuals, class arrays, project classes,
reference-order resolution or COM coercion. Exact existing ByRef declaration
checks remain; arbitrary interface/COM argument conversions stay open. Fixed
456 work entries and 48 CTest targets remain unchanged.

Completion: full desktop 48/48 (85.72 s), freshly rebuilt GUI-free core 1/1
(2.15 s), Debian payload/permissions/Build-ID/pinned SQLite and all four
original fixture hashes pass. This is a clean single complete suite run and
also covers the preceding TypeOf/temporary-clone validation correction.
Fixed 456 work entries, 384 confirmed scoped entries/72 open and 48 CTest
targets remain unchanged. No full Access/interface-binding claim.


## Object-reference actuals for native typed calls (2026-11-04)

Under existing F043/F122, same-module procedure signatures now advertise which
arguments require object-reference context. Sub calls, function expressions and
positional event-function expressions preserve actual Me/control/DAO member
references for Object/native class parameters, instead of prematurely reading
scalar defaults. Named function/Sub actuals choose context by parameter name;
ordinary scalar, array and intrinsic evaluation paths retain their behavior.
Native object methods can advertise the same context in argument signatures.
The callable resolver keeps legacy callback construction and argument-reference
transport, with optional signature metadata. Same-module per-frame signature
caches contain names/flags only, not host values, widgets, cursors or arrays.
Intrinsic dispatch keeps its existing precedence and evaluation paths.

Existing DAO passes actual recordset bang Fields directly to typed Sub/function
parameters and reads their real SQL values. Existing GUI passes Me!first directly
to a Control helper, including reversed named function actuals, and verifies
successful Form_Load. Existing event-prefix GUI executes a function-bound click
with an actual Control argument, changes its real Qt Tag and writes an SQL trace.
Core/DAO/mainwindow/layout targeted 4/4 (13.58 s) pass. Full validation follows.

This does not complete external-module/qualified-method binding, general property
ByRef setters, full parentheses/coercion/static-class rules, Variant default-
member context or named event-function argument ordering. Existing ByRef array
cell/allocation and native declaration guards remain. Fixed 456/48 unchanged.

Completion: full desktop 48/48 (85.57 s), freshly rebuilt GUI-free core 1/1
(1.99 s), Debian payload/permissions/Build-ID/pinned SQLite and all four
original fixture hashes pass. Existing live ByRef array cells and DAO clone/
transaction lifetime cases pass with optional call-context metadata. Fixed
456 work entries and 48 CTest targets remain unchanged.


## Same-module function precedence (2026-11-04)

Under existing F042/F118, parenthesized function expressions bind declared
same-module procedures before native library functions. Explicit VBA-qualified
calls keep their existing native library path. Procedure metadata records function
classification and result type; using a Sub as a value is explicitly unsupported
before its body executes. Private calls retain their own optional/named argument
handling, so a user Round helper receives Missing rather than builtin Empty holes.
Typed object argument context also follows the chosen private function.

A source-only procedure-name inventory lives with module state, avoiding repeated
whole-module searches for ordinary intrinsic calls. Custom module-state providers
without an inventory use a per-frame source-name fallback. Declared routines with
unsupported bodies/signatures do not quietly fall through to a library function.
MS-VBAL's two-argument Left exception is implemented for a zero-parameter scalar
function; its normal zero-argument call remains the user's function.

Existing core verifies user Len/Round, qualified VBA.Len, optional/named arguments,
Missing holes, counted bodies, scalar Left fallback and refusal to execute a Sub
as a value. Existing GUI passes a real Control to a private Len and executes a
function-bound Len click with actual Qt/SQL effects. Core/DAO/mainwindow/layout
4/4 (14.05 s) pass. Full validation follows. This does not complete bare-name
calls, scalar/indexed-variable and property/default-member binding, external/
project/reference precedence or full class scope. Fixed 456/48 remain unchanged.

Reference: Microsoft MS-VBAL, simple-name default binding namespace tiers and
its Left exception; [specification](https://learn.microsoft.com/en-us/openspecs/microsoft_general_purpose_programming_languages/ms-vbal/).

Completion: full desktop 48/48 (88.44 s), freshly rebuilt GUI-free core 1/1
(2.34 s), Debian payload/permissions/Build-ID/pinned SQLite and all four
original fixture hashes pass. A transient execution transport disconnect left
the workspace and running checks intact; no source was edited during checks.
Fixed 456 work entries and 48 CTest targets remain unchanged.

Begrenztes natives VBA-/SQL-Partition mit benannten Argumenten und tatsächlicher
SQL-Gruppierung ist implementiert: [Umfang und Grenzen](NATIVE_PARTITION.md).
Vollstand danach48/48 (84.06 s), GUI-freier Core1/1, Paket und vier Originalfixture-
Hashes bestanden; fester Arbeitskatalog unverändert456/384/72.

Native MkDir/RmDir-Anweisungen verändern echte Linux-Verzeichnisse im begrenzten
Umfang: [Umfang und Grenzen](NATIVE_DIRECTORY_STATEMENTS.md). Vollständige
VBA-Datei-I/O und Windows-Pfad-/Fehlergleichheit bleiben offen.

Begrenzte native FV-/Pmt-Berechnung läuft auch in tatsächlichen SQL-Abfragen:
[Umfang und Grenzen](NATIVE_FINANCIAL_FUNCTIONS.md). Vollständige Finanzbibliothek,
alle numerischen Randfälle und Windows-Oracle bleiben offen.

Extras/Optionen bietet echte gespeicherte UI-Schrift-/Stileinstellungen;
[Umfang und Grenzen](NATIVE_APPLICATION_OPTIONS.md). Vollständige Access-Menüs,
Optionen und Windows-Pixelgleichheit bleiben offen.

TableDef.Fields liest echte Main-Schema-Feldmetadaten ohne Nutzdatenzeilen:
[Umfang und Grenzen](DAO_TABLEDEF_FIELDS.md). Vollständiges DAO-Katalogmodell,
alle Feldtypen/-eigenschaften und Access-Lebensdauer-/Refresh-Gleichheit sind offen.

Neue Importe erhalten bestätigte Textlängen und Memo-Herkunft und erzwingen
Textgrenzen bei tatsächlichen SQL-/DAO-Schreibzugriffen:
[Umfang und Grenzen](SOURCE_TEXT_CAPACITY.md). Alte Projekte werden nicht
stillschweigend umgeschrieben; vollständige Jet-/UI-Textgleichheit bleibt offen.

Neue MDB-Importe erhalten96 bestätigte literale Original-Standardvorgaben;
native SQL-Defaults, DAO-AddNew-Puffer und neue gebundene Formularfelder führen
sie tatsächlich aus: [Umfang und Grenzen](SOURCE_LITERAL_DEFAULTS.md).
Allgemeine Ausdrücke und vollständige Windows-Default-/Eigenschaftsgleichheit
bleiben offen; keine Migration alter Projekte wird erfunden.

## Public standard-module functions in native SQL

The first native SQL/VBA scalar bridge executes actual public standard-module
bodies in SELECT and UPDATE expressions, sharing module/scalar Static state with
the existing per-connection runtime. Private/class/ambiguous sources, GUI/DAO host
operations, pre-existing object/array state and unproven result transports are
refused. Prepared source replacement, savepoint rollback, reopened connections,
TEMP shadows, builtin protection and private error-output boundaries are covered.
This does not add Currency/Date/object/array transport, complete Jet result metadata,
reentrant DAO/GUI or a Windows Access2003 oracle. Original coverage stays78/100.
See [native SQL/VBA scope](NATIVE_VBA_SQL.md). Production code full48/48 (87.35 s),
subsequent actual UPDATE DAO1/1 (0.74 s), current package/SQLite/four fixture hashes
pass. No new target/catalog ID;456/384/72 unchanged, F040/F068/F121 remain open.

Declared SQL/VBA Currency/Date returns now use native compiler declaration metadata,
checked CURRENCY_REAL/DATETIME transport and live result-declaration guards. Actual
DAO fields retain Currency5/Date8 and values; metadata preparation has no body
counter effects. Typed arithmetic/INSERT and stale prepared-signature rejection are
covered. Unknown Variant transports, complete compound/nested/argument type
propagation and complete make-table metadata remain open (complex typed shapes3251,
no target created). Unaliased function labels show the original expression instead
of internal guard calls. Full48/48 (85.94 s), fresh core1/1 (2.36 s), subsequent label
DAO/query/original3/3 (8.36 s), updated package/SQLite/four hashes pass. Original78/100
unchanged; [scope](NATIVE_VBA_SQL.md),456/384/72, F014/F068/F121 still open.

Simple native SELECT INTO now creates proven Currency/Date output declarations and
copies actual source rows atomically. Output metadata preparation does not evaluate
VBA bodies; failed source execution rolls back the destination. Unproven columns
remain untyped, and typed UNION/TOP/ranked/RIGHT JOIN shapes or duplicate names are
explicitly refused. AutoNumber/keys/defaults/validation/complete Jet metadata remain
open under F014/F130; no additional catalog point. docs/NATIVE_VBA_SQL.md.

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

Native table and saved-query datasheets now offer genuine print dialogs and
paginated atomic PDF output from checked snapshots of actual visible/filter/sort
state. Dirty drafts and unsupported binary/size/geometry contexts refuse explicitly.
Scope and remaining Access print/layout/oracle limits: [datasheet output](DATASHEET_PRINT_RUNTIME.md).
F080/F148/F087 remain open; fixed456/384/72 unchanged.


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

Datenblatt-Seiteneinrichtung: funktionierender gemeinsamer Qt-Geometriedialog, Sitzungslayout und explizite versionierte Projektpersistenz, native RunCommand PageSetup32, Papier/Ränder für Vorschau/Druck/PDF. Conditional Savepoint-Speicherung bewahrt konkurrierende Änderungen und Aufrufertransaktionen; ungültige gespeicherte Geometrie blockiert PDF bis zur expliziten Reparatur. Details docs/DATASHEET_PRINT_RUNTIME.md. Ziel1/1 (3,64 s), vollständiger aktueller Desktop48/48 (89,86 s) inklusive Originalrezepte2 Tabellen/39 Zeilen, Paket/Build-ID/SQLite3.54.1/vier Originalhashes bestanden. Keine neue Distributionsmatrix/GitHub-CI;456/384/72 unverändert, F080/F087/F148 bleiben offen.

Druckexemplare: reale Datenblatt-PDFs mit1..100 Kopien, sortierter/unsortierter und vorwärts/rückwärts Seitenfolge; native Qt-Treiberkopien getrennt von manuellen Kopien. Vorhandenes Ziel1/1 (3,99 s), aktuelle App/Paket/Build-ID/SQLite/vier Originalhashes bestanden. Letzte Vollsuitee8375e8:48/48 (89,86 s), dessen GUI-freier Kern1/1 (2,20 s). Unveränderlichee837-Folgeprüfungen Debian12/Ubuntu24.04:2/2 (6,65 s /5,65 s), passende Paketinstallation/Backend/vier Hashes; noch ohne Kopien-Inkrement. Details docs/DATASHEET_PRINT_RUNTIME.md.456/384/72 unverändert, keine physische Drucker-/Windows-Parität behaupten.

Auswahldruck: vollständige markierte sichtbare Datensätze als eingefrorener eigener Text-Snapshot, reale Tabellen-/Readonly-Abfrage-PDFs mit unabhängig geprüfter Ein-/Ausschlussmenge. Leere/teilweise Zellmarkierungen verweigern3251 ohne Ganzdatenblatt-Fallback. Ziel1/1 (4,15 s), aktuelle App/Paket/Build-ID/SQLite3.54.1/vier Originalhashes und drei Uploadhashes bestanden. Details docs/DATASHEET_PRINT_RUNTIME.md. Kein neuer Vollsuite-/Matrixlauf dieses GUI-Inkrements;456/384/72 unverändert.

DAO-DDL-Identität: beobachtete eigene Database/QueryDef.Execute CREATE/DROP und SELECT INTO führen monotone Main-Tabellengenerationen; alte TableDefs/Fields/physische Recordsets/Clones verweigern3420 statt neu erstellte gleichnamige Daten zu ändern. Gleiches SQL/Rootpage/Rowid, alter Schreibpuffer, frische Wrapper, No-op/Fehl-DDL, ALTER und Workspace-Rollback unabhängig geprüft. Externes/unbeobachtetes SQL, TEMP-only und volle Jet-Locks/Lebensdauer bleiben offen. Vollständiger aktueller Desktop48/48 (89,70 s), App/Paket/Build-ID/SQLite/vier Originalhashes bestanden; schließt Kopien/Auswahldruck und Originalrezepte ein. Details docs/DAO_RUNTIME.md.456/384/72 unverändert.

DAO.Relations: lesender tatsächlicher Importkatalog mit Name/Table/ForeignTable/Attributes und geordneten Fields.Name/ForeignName; Integer Count, benannte Indexargumente, echte VBA-Typen/TypeOf/Enumeration, gehaltene Identität und kontrollierte Invalidierung. Erzwungene Beziehungen werden mit allen tatsächlichen Main-Fremdschlüsseln und CASCADE-Aktionen abgeglichen; unbekannte Schlüssel/uneindeutige Metadaten verweigern3251 statt Teilkatalog. Bestehende Ziele2/2 (11,87 s), zusätzlich neue Kassa-Uploadkopie mit acht Originaldefinitionen/acht Feldpaaren bestanden. Details docs/DAO_RELATIONS_RUNTIME.md; Append/Delete/Indexes/Linked und vollständige Jet-/Windows-Parität bleiben offen.456/384/72 unverändert, kein neuer Testkatalogpunkt.

Nicht-Primärindexrichtungen: ursprünglicher Parserverlust von DESC behoben; vollständige geordnete ASC/DESC-Feldlisten für CREATE/ALTER INDEX, ungültige Zusätze verweigern statt Normalisierung. Native Indexanlage erhält Richtung, Reihenfolge und Unique in derselben Importtransaktion. Optionale echte Jet4-Referenz ohne Datensätze mit zwei unabhängig zurückgelesenen Indizes, tatsächlicher Gesamtimport und PRAGMA index_xinfo geprüft; keine App-Java-Abhängigkeit. Details docs/SOURCE_INDEX_DIRECTIONS.md. Kein echter Primärindexname aus generierten MDBTools-Namen behauptet; DAO.Indexes/Primärindexrichtungen/Locale/Originalformat weiterhin offen.456/384/72 unverändert.

Indexrichtungsnachweis: vollständiger aktueller Debian13-Desktop48/48 (89,51 s) bestanden, einschließlich tatsächlicher Jet-Binär-/Indexreferenzen, Originalrezepte für Datenblatt-PDF, Originalkassenhandler und aller originalen Kassa-Indexfeldreihenfolgen/-richtungen. Schließt DAO.Relations und Integer-Count ein. Keine neue Distributionsmatrix/GitHub-CI für diese Inkremente;456/384/72 unverändert.

DAO-Single-Werte: bestätigte physische Felder mit DAO-Typ6 liefern nun bei Value/Bang/Clone und im Edit/AddNew-Puffer echte Single-Werte; Schreiben verengt vor Puffermutation und persistiert binary32-Wert, Null/Empty und explizite Konvertierungsfehler bleiben geprüft. Echte technische Jet4-Float-Referenz mit unabhängig zurückgelesenen Werten, tatsächlicher Import/Metadaten/Update/Requery/Quellhash und VBA VarType/CSng im bestehenden Ziel2/2 (1,94 s) bestanden. Keine automatische Query-/Formularlineage und keine verlustfreie allgemeine Single-Importbehauptung: MDBTools1.0.0 %.8g kann einzelne Bitmuster verändern, belegtes Gegenbeispiel in docs/DAO_SINGLE_FIELDS.md. F014/F041/F055/F105/F122 bleiben offen,456/384/72 unverändert. Letzte Vollsuite4548ce7:48/48 (89,51 s), nicht diesem Typinkrement zuschreiben.

Originale Indexmetadaten: separater lesender nativer MDBTools-Helfer statt generierter CLI-Primärnamen. Strikt versionierte Namen/Kind/Flags/geordnete Felder/Richtungen, gegen tatsächliche Tabellenspalten und ursprünglichen Primärschlüssel validiert; atomare main._la_access_indexes-Persistenz. Neue Importe verwenden echte Nichtprimärnamen. Zusammengesetzte und einzelne DESC-Primärschlüssel bleiben echte native DESC-Indizes; DESC-AutoNumber und unbewiesene IgnoreNulls/Required-Konfigurationen verweigern explizit. Bestehende GUI/Import/Kassa-Ziele3/3 (15,27 s), abschließende erweiterte technische Jet-Referenz1/1 (0,96 s);20 echte Primär-/48 andere Originalnamen in29 Tabellen nachgewiesen. Details docs/SOURCE_INDEX_METADATA.md. Noch kein vollständiger DAO.Indexes-/Windows-Nachweis;456/384/72 unverändert. Build benötigt mdbtools-dev/pkg-config, Paket enthält geprüften lesenden Helfer.

Indexmetadaten-Grundlage vollständig verifiziert: aktueller Debian13-Desktop48/48 (85,91 s), neue echte Jet-Referenzen einschließlich einzelner und zusammengesetzter DESC-Primärindizes, Originalrezepte, Originalkassenhandler und Originalindexnamen bestanden. Aktuelles Paket/Build-ID für App und Metadatenhelfer, tatsächliches Laden des entpackten Helfers, SQLite3.54.1 sowie vier Originalfixture-/drei Uploadhashes bestanden. Kein neuer Katalogpunkt und keine vollständige DAO.Indexes-/Windows-Parität;456/384/72 unverändert. Lokale Debian12-/Ubuntu24.04-Folgeprüfung der neuen Build-Abhängigkeit folgt separat.

DAO.Indexes: lesender originaler Indexkatalog mit nativen Typen/TypeOf, Integer Count/Item/named Index/Refresh/Enumeration, echten Namen und Boolean Primary/Unique/Required/IgnoreNulls sowie geordneten Fields.Name/Long Attributes (dbDescending). Jeder Zugriff vergleicht den vollständigen tatsächlichen Main-Indexkatalog, Primärschlüssel/Rowid-Alias, Unique, Felder/Richtungen und BINARY-Collation. Fehlende Quellen, logische kind2/IgnoreNulls/unbekannte oder zusätzliche native Indizes verweigern3251 ohne Teilkatalog. Gehaltene Definitionen/Fields invalidieren kontrolliert nach Schema/Metadaten/Rollback, TEMP bleibt wirkungslos. Bestehende Ziele3/3 (11,69 s), zusätzlich neue Kassa-Kopie mit20 Primär-/48 anderen DAO-Indexidentitäten, aktuelle App/Paket/Build-IDs/Helfer/SQLite3.54.1/vier Originalfixture-/drei Uploadhashes bestanden. Details docs/DAO_INDEXES_RUNTIME.md. Keine Mutation/Foreign/Clustered/DistinctCount/Windows-Parität;456/384/72 unverändert. Letzte Vollsuite bleibt74d75c7:48/48 (85,91 s).

Unveränderliche74d75c7-Folgeprüfung abgeschlossen: Debian12/GCC12/Qt6.4 und Ubuntu24.04/GCC13/Qt6.4 jeweils5/5 gezielte Ziele (30,64 s /27,39 s), neues mdbtools-dev/pkg-config, Paketinstallation, Backend3.54.1, Originalhashes und tatsächlich installierter Helfer mit originalen Jet-Namen/Richtungen unter MDBTools1.0.0 bestanden. Diese Snapshots enthalten die Metadaten-Grundlage, noch nicht die folgende DAO.Indexes-Anbindung; keine GitHub-CI- oder Vollsuitebehauptung für die Matrix.

Exakter Single-Import: realer MDB_FLOAT-Rohwertzugriff im vorhandenen lesenden Helfer (--single-values), IEEE754-Bitmuster statt verlustbehafteter %.8g-Speicherung. Source-SHA256 vor/nach Scan und atomarem Import, echte SINGLE-Spalten/REAL-Affinität, CSV-Zeilen/NULL/8g-Abgleich,1M-Zellen/16MiB-Grenze. Rohwerte bleiben außerhalb des angezeigten SQL-Schemas. Gegenfallbits509022793, minimale Subnormale, maximaler endlicher Single, Brüche/+0/NULL nach unabhängiger Jackcess-Jet-Erzeugung im tatsächlichen Gesamt-/Teilimport und DAO bitgenau nachgewiesen. NaN/Infinity/negatives Null oder inkonsistente CSV-Daten verweigern atomar. Bestehende relevante Ziele5/5 (27,76 s), finaler Affinitätsguard im Importziel1/1 (1,05 s); aktuelle App/Paket/Build-IDs/SQLite3.54.1/vier Originalfixture-/drei Uploadhashes und tatsächlich entpackter Helfer mit bekannten Originalbits/Quellhash bestanden. Details docs/SOURCE_SINGLE_PRECISION.md. Legacy-CSV, Double/Decimal/Currency, generische SQL-/Default-/Query-Semantik und Windows bleiben offen;456/384/72 unverändert. Letzte Vollsuite74d75c7:48/48 (85,91 s), neue Matrix noch nicht behaupten.

Double-Import bitgenau: vorhandener lesender Helfer --double-values, binäre64-Bit-Quelle statt %.16g-Rückwandlung. Gemeinsamer streng validierter access_floating_metadata-Codec für Single/Double, Classic-Locale-Formatprüfung ohne Double-Overflow bei der Dezimaldarstellung von MAX_VALUE. Original-Rohwerte aus der bereits geschützten Gesamtimport-Snapshotdatei; strukturleere Anlage scannt Daten nicht doppelt. Low-level API unterscheidet ausdrücklich caller-owned Csv (kein Dateizugriff aufgrund eines Loglabels) von NativeMdbBits; reale Gesamt-/Inspector-Kopien und Originalprüfungen wählen NativeMdbBits, fehlender Nachweis fällt niemals still auf CSV zurück. Dadurch bleiben technische Tests mit echten Schemas und eigens bereitgestellten CSV-Zeilen korrekt getrennt.

Unabhängiger Jackcess-Auszug beider Kassa-Testkopien:13 Tabellen/57 Double-Spalten,69405 Nicht-NULLwerte/514 NULLs, kein negatives Null/NaN/Infinity. Der bisherige Dezimalexport reproduzierte2404 numerische Zeilentupel in zwei Tabellen nicht exakt; neue native Kopien stimmen bitgenau mit allen unabhängigen Tupeln überein. Echte Jet-Extremwertreferenz, Gesamt-/Teilimport, DAO Type7/Size8, Herkunfts-/CSV-Guards und atomare Negativnull-Verweigerung geprüft. Aktueller korrigierter Debian13-Desktop48/48 (143,10 s), aktuelle neue Uploadkopie erneut29/29 und unabhängiger Auszug exakt; dort weiterhin78/100 Originalabfragen ohne Windows-Ergebnisorakel. App/Paket/Build-IDs/SQLite3.54.1/vier Originalfixture-/drei Uploadhashes und tatsächlich entpackter --double-values-Helfer mit bekannten Originalbits/SHA bestanden. Details docs/SOURCE_DOUBLE_PRECISION.md. Zuvor entdeckte technische CSV-Herkunftsfehler behoben, keine Schutzprüfung für Originale entfernt.456/384/72 unverändert; keine neue Matrix für Double behaupten.

Nach einmaligem lesendem Resume derselben Codex-Call-Anfrage3ea2a1b8-75c6-411c-8f85-3cb76c4f3571 wurde status=unresolved und eine normale finale LiveKit-Teilantwort „Genau“ zur Windows-Referenzfrage geliefert; keine konkrete Umgebung/Protokolle/Zugriffsinformation und keine confirmed_decisions. CI-Frage bleibt unbeantwortet. Beide externen Blocker bleiben offen, keine Freigabe erfinden oder neue Anrufanforderung daraus ableiten. Privates Ergebnis resumed-result.json im bestehenden Call-Verzeichnis.

Nativer DAO-Schemaaufbau: eigene detached Tabellen/Felder bis zur atomaren
physischen TableDefs.Append-Anlage mit tatsächlich wirksamen Pflichtfeldern,
literalen Defaults, Textkapazität und Long-AutoNumber. Echter VBA-Durchlauf und
DAO/Import2/2 (1,25s) bestanden; App gebaut. Umfang/Grenzen:
[DAO-Felder](DAO_TABLEDEF_FIELDS.md). F055/F121 offen,456/384/72 unverändert.

Eigene DAO-Tabellen erhalten detached Primär-/Unique-/ASC-/DESC-Indizes über
CreateIndex/Index.CreateField/Indexes.Append. Tatsächliche Kataloge/Constraints und
atomarer Kollisions-Rollback/Retry geprüft; bestehende DAO/Import2/2 (1,22s), App
gebaut. Details DAO_INDEXES_RUNTIME.md. Keine vollständige F121/F130-Parität;
456/384/72 unverändert.

DAO-Nichtprimärindizes können auf bereits vorhandenen vollständig geprüften
Tabellen tatsächlich angelegt/gelöscht werden. Unique-Fehleratomizität, geordnete
ASC/DESC-Felder, Workspace-Rollback und immutable gehaltene Objekte geprüft;
DAO/Import2/2 (1,31s), App gebaut. Umfang/Primär-/Fremdschlüsselgrenzen in
DAO_INDEXES_RUNTIME.md; F121/F130 offen,456/384/72 gleich.

Grafischer Tabellenentwurf verwendet gemeinsamen DAO-Schemaweg einschließlich
nativer Text-/Integer-Typgrenzen und vollständiger neu deklarierter Primärindex-
Identität. Eigene UI-Tabelle und anschließende DAO-Indexanlage im bestehenden
Mainwindow/DAO2/2 (4,95s) geprüft, App gebaut. DAO_TABLEDEF_FIELDS.md;
keine Bestandsmigration,456/384/72 unverändert.

Grafische Nichtprimär-Indexverwaltung über Ansicht → Tabellenindizes bearbeiten:
vollständiger geprüfter Katalog, geordnete ASC/DESC-Felder, Unique/Required, Anlage
und bestätigte Löschung über DAO. Realer GUI-Create/Delete-Durchlauf mit
unveränderten Datensätzen; Mainwindow/DAO2/2 (5,11s), App gebaut. Grenzen in
DAO_INDEXES_RUNTIME.md.456/384/72 unverändert.

Native Felddefinitionen schützen primitive Speichergrenzen auch bei INSERT SELECT
und bewahren explizite Required-/Attributes-Deklarationen. Kombinierte Wertebereichs-
und GUI-Prüfung2/2 (5,00s); abschließender primitiver DAO-Fall1/1 (1,21s), App
gebaut. Details/Coercion-Grenzen DAO_TABLEDEF_FIELDS.md. Vier Fixture-/drei Upload-
hashes unverändert, Repo privat;456/384/72 gleich, keine Gesamtparität.

Grafischer Tabellenentwurf: Textgröße1..255 wählbar, Defaultgrößenprüfung vor DDL,
retinierter fehlerhafter Entwurf und tatsächliche korrigierte Größe12 geprüft.
Mainwindow1/1 (3,91s), App gebaut. DAO_TABLEDEF_FIELDS.md;456/384/72 unverändert.

DAO-Normalfelder vorhandener Tabellen: echte atomare CreateField/Fields.Append/
Delete-Operationen mit Feld-/Indexmetadaten, Bestandswerten/Defaults und Dependency-
Schutz. Mainwindow/DAO2/2 (5,38s), App gebaut; genauer Umfang und offene Counter-/
Primär-/Linked-/Artefaktgrenzen DAO_TABLEDEF_FIELDS.md.456/384/72 unverändert.

Grafische normale Feldanlage/-löschung für bestehende Tabellen mit echten
Bestandswerten/Defaults, Bestätigung und Abhängigkeitsgrenzen; eigene Struktur-
änderungen invalidieren alte DAO-Cursor/Fields/Clones bis Requery. Tatsächlicher
GUI-Durchlauf und Clone-Pufferschutz: Mainwindow/DAO2/2 (5,11s), App gebaut.
DAO_TABLEDEF_FIELDS.md/README;456/384/72 unverändert.

TableDef.Fields/Indexes.Refresh lädt gehaltene Sammlungen nach eigener DDL oder
Rollback kontrolliert neu; alte Items und gleichnamig ersetzte Tabellen bleiben
invalid. DAO1/1 (1,23s), App gebaut; COM-/readonly-Kontextgrenzen in
DAO_TABLEDEF_FIELDS.md.456/384/72 unverändert.

Native DAO-Beziehungsbuilder und atomare Katalogmutation für ausdrücklich
nicht erzwungene Beziehungen umgesetzt; Persistenz, Abhängigkeiten, Same-name-
Generation/Rollback im bestehenden DAO/Import2/2 (1,33s) geprüft, App gebaut.
Keine physische RI behauptet; erzwungene/cascade Mutation bleibt offen. Umfang
DAO_RELATIONS_RUNTIME.md.456/384/72 unverändert.

Reale native numerische FK-/Cascade-Migration einschließlich Entfernung umgesetzt:
FK-Checks, strenger Schemanachweis, rowid-/Wertdigest, Index-/Trigger-/View- und
Zählererhalt, geschützte PRAGMA-Wiederherstellung. Finales bestehendes DAO1/1
(1,33s), App gebaut; weiterer Kindverweis und unbekannte CHECK-Regel geschützt.
DAO_RELATIONS_RUNTIME.md enthält Herkunfts-/Typ-/Transaktionsgrenzen; keine
vollständige Jet-/Original-MDB-Parität.456/384/72 unverändert.

Grafische Beziehungen auf gemeinsamem DAO-Weg: echte numerische Integritätsprüfung/
Kaskaden oder explizites DontEnforce, geordnete Feldzuordnung und bestätigte
Entfernung bei erhaltenen Daten. Reale modale GUI-Anlage/physische FK-Klauseln/
Entfernung Mainwindow/DAO2/2 (5,33s), App gebaut. DAO_RELATIONS_RUNTIME.md;
Quell-/Typ-/Transaktions-/1:1-Grenzen bleiben offen,456/384/72 unverändert.

Native dbRelationUnique1: echte geprüfte Unique-Schlüssel beider Seiten und
Indexabhängigkeitsschutz, kombinierbar mit FK/Cascades oder DontEnforce3.
Backend-/reale GUI-Prüfung Mainwindow/DAO2/2 (5,64s), App gebaut.
DAO_RELATIONS_RUNTIME.md; keine Jet-NULL-/automatische Flag-Inferenz-Parität,
456/384/72 unverändert.

Vorhandene native Feldproperties Required/DefaultValue/Text.Size jetzt über DAO
und grafischen atomaren Batch-Save änderbar. Geprüfter Schema-/Bestandswert-/
Countererhalt, Fehlerrollback und Cursorinvalidierung bis Requery; Original-/
Counter-/unbekannte Schemaänderungen bleiben begrenzt. Unabhängige unbenannte
Fremdschlüssel blockieren den Feldumbau nicht; betroffene unbekannte Schlüssel
bleiben geschützt. Bestehendes GUI/DAO2/2 (5,92s), App gebaut. Details/Grenzen
DAO_TABLEDEF_FIELDS.md.456/384/72 unverändert. Weitere offene Anforderungen
fortlaufend implementieren; kein Zwischenabschluss oder Abschlussanruf.

Native Nicht-Counter-Felder sind jetzt per DAO Field.Name und grafisch atomar
umbenennbar, einschließlich Index-/Beziehungsmetadaten, echten FK-Verweisen und
Digestprüfung der Werte/rowids. Access-Abfrageartefakte/VBA/Formularbindungen
werden nicht refaktoriert. Echtes GUI/DAO2/2 (5,66s), App gebaut. Umfang/Grenzen
DAO_TABLEDEF_FIELDS.md.456/384/72 unverändert. Weiter native Feldvalidierung
und weitere offene Anforderungen implementieren, keine Zwischen-Finals.

Native AllowZeroLength für Text/Memo ist bei DAO-Anlage und vorhandenen Feldern
sowie in Neuer-Tabelle-/Feld-Dialogen implementiert, echte SQL-Prüfung/3315,
Null separat Required. Physische DAO-/Clonefelder lesen explizite Required-/
AllowZeroLength-Metadaten schreibgeschützt. Atomarer Umbau/Fehlerrücknahme;
Original-/unbewiesene Eigenschaften nicht erfunden. GUI/DAO2/2 (5,99s), App
gebaut; Details DAO_TABLEDEF_FIELDS.md.456/384/72 unverändert. Nächste Arbeit:
Primärindizes vorhandener nativer Tabellen wirklich anlegen/entfernen, samt
Daten-/Counter-/Abhängigkeitsschutz und grafischer Bedienung; dann weiter.

Primärindizes vorhandener nativer Tabellen sind jetzt per DAO und grafisch
wirklich anleg-/entfernbar: echter PK/NOT NULL, atomarer Daten-/Counter-/
Regelerhalt, Abhängigkeits-/Cursor-/Transaktionsschutz. GUI1/1 (4,30s), finales
DAO1/1 (1,64s), App gebaut; keinen frischen2/2-/Gesamtlauf behaupten. Details
DAO_INDEXES_RUNTIME.md.456/384/72 unverändert. Als Nächstes Required-Indizes
vorhandener nativer Tabellen mit echten Nullregeln auch auf bisher nullable
Feldern sowie korrekter Regelentfernung umsetzen; danach fortlaufend weiter.

Required-Indizes vorhandener nativer Tabellen erzwingen jetzt echte NOT-NULL-
Regeln auch auf bisher nullable Feldern. Vorhandene Nulls verweigern atomar3314;
Löschen entfernt nur implizite Indexregeln, andere Pflichtregeln bleiben erhalten.
Entfernte physische Indizes werden nicht versehentlich wiederhergestellt. GUI/
DAO2/2 (5,73s), App gebaut; DAO_INDEXES_RUNTIME.md.456/384/72 unverändert.
Nächste Arbeit: sichere physische DAO-Zeilenidentität bei normalen Nutzerfeldern
namens rowid/_rowid_/oid; danach andere offene Anforderungen ohne Zwischenfinal.

Physische DAO-Recordsets verwenden jetzt einen nachgewiesenen unverschatteten
rowid/_rowid_/oid-Alias für Lesen, Update/RETURNING und Löschen; normale gleich
benannte Nutzerfelder können keine falschen Zeilen mehr adressieren. Familie/
Clone/Requery übernehmen den Alias. Alle drei Aliasnamen oder WITHOUT ROWID
verweigern3251. SQLite3.54.1-Alias-ALTER-Grenze wird ebenfalls ausdrücklich3251,
keine vollständige Schemaunterstützung behaupten. Final DAO1/1 (1,58s), App
gebaut, DAO_TABLEDEF_FIELDS.md.456/384/72 unverändert. Nächste größere native
Katalogaufgabe: TableDef.Name für vorhandene eigene Tabellen samt tatsächlicher
Counterdefault-/Index-/Beziehungsmetadaten-Umbindung; dann weiter ohne Zwischenfinal.

TableDef.Name und grafische Tabellenumbenennung für eigene native Tabellen
implementiert: tatsächliche Counterdefault-/Frontier-/Schritt-/Index-/
Beziehungs-/Seiteneinstellungs-Umbindung, physische FK/View/Trigger-Verweise,
Digest-/Generations-/Rollbackschutz. Gespeicherte Access-Abfragen/VBA/Formulare
werden nicht refaktoriert; weitere Grenzen DAO_TABLEDEF_FIELDS.md. GUI/DAO2/2
(5,82s), App gebaut. Originaluploads3/3 und Fixtures4/4 Hashes unverändert.
456/384/72 unverändert. Nächste Aufgabe: echte Verbindungs-Schreibfähigkeit und
readonly DAO-Fehler/Metadaten sowie nachgewiesene lokale TableDef-Metadaten;
dann weitere offene Anforderungen fortlaufend, kein Zwischenabschluss.

DAO-Schreibfähigkeit prüft jetzt echtes SQLite-readonly und query_only statt
lediglich Recordsettyp. Database/TableDef/Recordset.Updatable, readonly Counter-
Metadaten und3027-Verweigerung vor Daten-/Countermutation implementiert; Cancel
bleibt möglich. Explizit eigene lokale TableDefs haben Attributes0/Connect=""/
SourceTableName=""; keine erfundenen importierten Linkdaten oder externen Provider.
Final DAO1/1 (1,66s), App gebaut, technischer Readonly-Dateivergleich unverändert.
DAO_TABLEDEF_FIELDS.md;456/384/72 gleich. Weitere offene DAO-/Schemaanforderungen
fortlaufend implementieren; keine Zwischen-Finals/Abschlussanrufe.

Native Single-Felder6/Size4 jetzt im DAO-/Tabellen-/Feldbau verfügbar, mit echter
quantisierter REAL-Speicherung/Prüfung. Gemeinsame Single-Konvertierung für DAO,
Datenblatt und begrenzte direkte SQL-Schreibpfade/physische Append-Selects,
Literaldefaults/ältere Defaultinserts und direkte Query-Typweitergabe. Negative
Null/Unterlauf bleiben3251 statt stiller Normalisierung. Final GUI/DAO2/2
(6,12s), App gebaut; vorher Import1/1 (0,06s), kein frischer3/3-Lauf. Grenzen
DAO_TABLEDEF_FIELDS.md/SOURCE_SINGLE_PRECISION.md; insb. komplexe/joined Writes,
Arithmetik/Locale/Access-Oracle und erzwungene Single-FKs noch offen.456/384/72
unverändert. Andere offene Anforderungen fortlaufend weiterprogrammieren.

Echte Single-Fremdschlüssel/Cascades jetzt für nachgewiesene eigene SINGLE_REAL-
Felder beider Seiten, einschließlich GUI-Auswahl. Nicht nachgewiesene Legacy-/
importierte REAL-Single-Schlüssel bleiben verweigert. GUI/DAO2/2 (6,26s), App
gebaut; DAO_RELATIONS_RUNTIME.md. Nur der Titelguard des GUI-Fehlerhelfers wurde
danach eingeengt, keine Produktionsänderung nach Prüfung.456/384/72 unverändert.
Fortsetzung: native Date-/Now-/Time-Defaults und Datumsnormalisierung prüfen/
implementieren, danach weitere offene Anforderungen ohne Zwischen-Final.

Native Datumsdefaults Date()/Now()/Time() jetzt für eigene Date-Felder im DAO-,
Tabellen-/Feldentwurf, manuellen Datenblatt und direkten Tabellenformular.
Neue Puffer speichern den initialisierten Uhrwert; explizit Null überschreibt,
ungebundene Datumsdefaults werden mitgespeichert. Neue Date-Felder erzwingen
kanonische ISO-Speicherung; alte Definitionen bleiben bei Migration erhalten.
Nur begrenzte eigene Uhrfunktionen, keine beliebigen SQL-/VBA-Defaults. App
gebaut, GUI/DAO2/2 (6,24s). SOURCE_LITERAL_DEFAULTS.md/DATE_RUNTIME.md;
F014/F015/DAO-/Formular-Vollanforderungen bleiben offen,456/384/72 unverändert.
Fortsetzung: eigene kanonische Date-Beziehungsschlüssel samt GUI und tatsächlichen
FK-/Kaskadenwirkungen; danach weitere offene Anforderungen, kein Zwischen-Final.

Eigene kanonische Date-Schlüssel jetzt für tatsächliche referenzielle Integrität,
Änderungs-/Löschkaskaden und Beziehungsdialog freigegeben. Beide Seiten benötigen
nachgewiesene eigene Date-Metadaten und DATETIME-Speicherung; ältere/importierte
Date-Schlüssel bleiben3251. Bestehende Uhrdefaults und binäre Nutzwerte bleiben
bei der Migration erhalten. GUI/DAO2/2 (6,11s), App gebaut. Keine vollständige
Jet-Datums-/Locale-/Untersekundenparität;456/384/72 unverändert. Fortsetzung:
native numerische Gültigkeitsregeln im DAO/Feldentwurf mit tatsächlicher
Speicherprüfung und atomarer Bearbeitung, danach weitere offene Anforderungen.

Numerische native Field.ValidationRule im DAO/Feldentwurf mit echten CHECKs
implementiert: explizite Vergleiche mit And/Or oder Between, Null separat von
Required,3317 bei Verletzung. Atomare Regeländerung prüft vorhandene Werte;
Rollback, fehlerhafte Mehrzeilenänderung, erhaltene DAO-Puffer, Löschen der Regel,
Feld-/Tabellenrename und Byte/Integer/Long/Single/Double/Boolean geprüft.
GUI/DAO2/2 (7,08s), App gebaut; DAO_TABLEDEF_FIELDS.md. Allgemeine Expressions,
ValidationText und volle Referenzsemantik bleiben offen;456/384/72 unverändert.
Fortsetzung: eigene Validierungstexte und physische Recordset-Feldmetadaten,
danach übrige offene Anforderungen ohne Zwischenabschluss.

Native ValidationText jetzt im DAO/Feldentwurf für unterstützte numerische Regeln,
mit wirklichem3317-SQL-Fehlertext; UTF-8/Umlaute, Großschreibung und Quotes bleiben
erhalten. Leeren/Bearbeiten/Migration unterstützt; Begrenzung255 UTF-16-Einheiten,
kein NUL. Physische Recordset.Fields lesen eigene ValidationRule/ValidationText
schreibgeschützt, mit bestehenden Clone/Requery-/Generationsguards. GUI/DAO2/2
(7,32s), App gebaut. Allgemeine Ausdrücke und genaue Windows-Formular-/Fehler-
Ereignissemantik weiterhin offen.456/384/72 unverändert. Fortsetzung: tatsächliche
Feldreihenfolge/OrdinalPosition im DAO und Feldentwurf mit atomarem Werterhalt.

Native TableDef.Field.OrdinalPosition verschiebt jetzt echte physische Spalten,
auch Counter, atomar mit Feld-/Indexkatalog und GUI-Positionssteuerung. Vorhandene
Werte/Rowids, Counterfrontier, Single-/Uhrdefaults, Regeln/Texte, Indizes, FK,
Trigger und Views bleiben erhalten; physische DAO-Puffer/Clones3420 bis Requery.
Unveränderte Position ohne Schemaänderung, außerhalb vorhandener Positionen3251;
keine ursprünglichen/unklaren Schemas oder aktiven Aufrufertransaktionen umbauen.
GUI/DAO2/2 (7,15s), App gebaut. DAO_TABLEDEF_FIELDS.md;456/384/72 unverändert.
Fortsetzung: Gültigkeitsregeln für eigene kanonische Date-Felder einschließlich
Uhrgrenzen im DAO/Feldentwurf, dann übrige offene Anforderungen fortlaufend.

Gültigkeitsregeln jetzt auch für eigene kanonische Date-Felder: ISO-#Datum#-
Konstanten (ganze Sekunden) und Date()/Now()/Time() mit Vergleichen, And/Or,
Between. Tatsächliche CHECK-Durchsetzung/3317 und ValidationText, atomare
Bestandsprüfung, Rename/Migration/GUI; alte/importierte Date-Speicherung bleibt
verweigert. Regionale Konstanten, Untersekunden/allgemeine Ausdrücke offen.
GUI/DAO2/2 (7,08s), App gebaut;456/384/72 unverändert. Fortsetzung: verlustfreie
numerische Datentypänderungen im Feldentwurf mit Bestands-/Defaultprüfung;
DAO Field.Type nach Append bleibt readonly. Danach offene Anforderungen weiter.

Feldentwurf unterstützt jetzt verlustfreie numerische Typänderung für eigene
Nicht-Counter: Boolean/Byte/Integer/Long/Single/Double mit Bestands-/Default-
Konvertierungsprüfung vor Umbau. Verlust/Rundung verweigert3251, Überlauf6;
Null, Indizes, Regeln/Texte und Counterfrontier geschützt. Numerische Werte
werden bei beabsichtigtem Speicherklassenwechsel bitgenau verglichen; andere
Werte/Rowids bleiben typgenau. Relationkeys vorher lösen; DAO Field.Type nach
Append weiter383. GUI/DAO2/2 (6,94s), App gebaut; DAO_TABLEDEF_FIELDS.md.
456/384/72 unverändert. Fortsetzung: doppelte Hauptfenster-Kompilierung von App
und GUI-Ziel über gemeinsame statische Desktopimplementierung beseitigen, dann
weitere offene Entwurfs-/Access-Anforderungen ohne Zwischenabschluss.

Native Datum-Standardwerte akzeptieren jetzt eindeutige ISO-#Datum#-Literale
(optional Uhrzeit mit ganzen Sekunden) und führendes=. Begrenzte allgemeine
Literaldefaults akzeptieren ebenfalls= und äußeren Leerraum; Textinhalt bleibt
unverändert. DAO-Anlage/Bearbeitung, Tabellen-/Feldentwurf und Datumsregeln nutzen
gemeinsamen Date-Parser; ungültige Tage/regionale Formen verweigern3251. GUI/DAO2/2
(7,98s), App gebaut; erster Compilefehler war fehlender Prüfcode-Include und ist
behoben. Kein Importparser-/Windowslexik-/Vollsuiteclaim;456/384/72 unverändert.
Fortsetzung: Table-Recordsets/Index/Seek für belegte native numerische Indizes,
mit tatsächlicher Indexreihenfolge und klaren unbewiesenen Grenzen.

Native physische Table-Recordsets1 jetzt explizit/dbOpenTable sowie als Default
für Main-Tabellen/TableDefs; SQL-/QueryDef-Dynasets2/Snapshots4 bleiben erhalten.
Geprüfte native nichtnullable numerische Indizes liefern tatsächliche Reihenfolge,
Index/Seek mit kompletten Schlüsseln, ASC-Bereichen, DESC-Gleichheit, NoMatch/
Bookmarks, benannten Argumenten und kontrollierten Clonegrenzen. NativeSnapshot-
Suche, kein B-tree-/Live-Jet-Claim; Date/Text/Null/Partial/DescendingRange und
unabhängiger Clone-Indexwechsel noch offen. DAO_RECORDSET_INDEX_SEEK.md mit
aktuellen Microsoft-API-Referenzen, kein2003-Oracle. Core/GUI/DAO/originalesKassa-
Ereignis4/4 (122,64s; davon Originalimport113,62s), App gebaut. Originaluploads3/3,
Fixtures4/4 Hashes unverändert. Keine erneute breite Originalprüfung ohne neuen
betroffenen Umfang.456/384/72 gleich. Fortsetzung: eigene kanonische Date-Index-
Schlüssel und typed Seek, danach offene Anforderungen ohne Zwischenabschluss.

Index/Seek jetzt auch für eigene kanonische nichtnullable Date-Felder und gemischte
Date-/Zahlenschlüssel. Gebundene Date-Variants und VBA-#Datum# bleiben typisiert;
String-Datekeys verweigern13. Datebereiche vergleichen kanonische ganze Sekunden
entsprechend tatsächlicher Indexreihenfolge. DAO1/1 (2,20s), App gebaut; keine
Wiederholung des unverändert betroffenen Originalimports. Voller Jet-/OLE-/Locale-
Oracle bleibt offen,456/384/72 gleich. Fortsetzung: dokumentierte zusammengesetzte
Teilschlüssel-Suche mit fehlenden hinteren Schlüsseln als Null für die bereits
strukturell nichtnullen unterstützten Indizes; danach weitere offene Anforderungen.

Seek unterstützt jetzt ausgelassene hintere Schlüssel zusammengesetzter
nichtnuller eigener Indizes: Suffix wird gemäß dokumentierter Access-Suche als
Null behandelt; Gleichheit NoMatch, aufsteigende Präfixbereiche passend. Numerik/
Date, explizit Missing am Suffix, Bookmark-Rettung geprüft. Explizite Nullwerte,
Lücken und absteigende Bereiche bleiben3251. DAO1/1 (2,25s), App gebaut;
456/384/72 gleich. Fortsetzung: eigene Feld-Caption/Description im Entwurf mit
wirklichen Tabellenköpfen und Formular-/Berichtslabels, dann übrige Anforderungen.

Eigene Feld-Caption/Description jetzt tatsächlich im Projekt, Entwurf (auchCounter),
Datenblattkopf/Hinweisen und kopierten Formular-/Berichtslabels. Physische Namen/
Werte/Counter bleiben gleich. DAO.CreateProperty Text10 sowie begrenztes eigenes
Field.Properties Item/Append/Delete/Refresh, Value/defaultValue, echte native
Property/Properties-Typbindung und readonly Recordset-Properties implementiert.
Persistente Property-IDs verhindern Wiederbelebung nach Delete/Recreate; keine
vollständige Count-/Ordinal-/Enum-Liste erfunden. Schema-/Rollbackguards bleiben.
Details NATIVE_FIELD_DISPLAY_PROPERTIES.md. GUI1/1 (5,65s); danach fehlende
Typregistrierung/TypeOf-Registrierung korrigiert und zu gemeinsamer Registry
vereinheitlicht: Core/DAO2/2 (4,66s), App gebaut. Kein frischer3/3-/Vollsuiteclaim.
456/384/72 unverändert. Fortsetzung: weitere tatsächlich vorhandene native Feld-
Property-Metadaten über Properties statt nur direkte Member anbinden; vollständige
originale Properties-/Windowsparität bleibt offen. Kontinuierlich weiterarbeiten.

Native Field.Properties stellt jetzt zusätzlich vorhandene intrinsische Metadaten
über echte Property-Objekte bereit: Name/Type/Size/Attributes/DataUpdatable/
OrdinalPosition/Required/AllowZeroLength/DefaultValue/ValidationRule/-Text, jeweils
nur soweit direkter Getter belegt. Kontext bleibt erhalten (Counter TableDef17,
Recordset49); Value-Schreiben nutzt vorhandene atomare Setter, readonly bleibt383.
Eigene Property-Familie folgt zulässigen Größen-/Pflicht-/Namensänderungen; kein
Count/Enum/Vollkatalog erfunden. DAO1/1 (2,46s), App gebaut;456/384/72 gleich.
Fortsetzung: tatsächliche Caption/Description-Spaltenproperties aus MDBTools
lesend übernehmen; vorhandene mdb_col_get_prop-API geprüft, keine VBA-Extraktion.

Echte Jet4-Caption/Description-Spaltenproperties werden jetzt aus vorhandener
lesender mdb_col_get_prop-API im selben Metadatenlauf übernommen (keine extra
Prozesse proTabelle). Versionierter begrenzter Abschnitt, vollständige geordnete
Spaltenbindung, Format-/SHA-/Textprüfung; fehlende Komponente bleibt unbekannt.
Import speichert belegte Werte/Abwesenheit, leere Strings und eigene Lease-IDs;
Datenblatt/Form-/Berichtserzeugung und readonly DAO-Properties lesen sie.
Andere Formate/alte Projekte/Vererbungsflag/Vollkatalog nicht geraten. Echte
unabhängige Jackcess-Jet4-Referenz ohne Datensätze bestätigt Unicode/Quotes,
Description-Zeilenumbruch, leer/fehlend, MDB→SQLite/DAO/Kopf und Quellhash.
Import/DAO2/2 (2,57s), App/lesender Helfer gebaut. Java nur Prüfdaten-Erzeugung,
keine Runtime-Abhängigkeit. SOURCE_FIELD_DISPLAY_PROPERTIES.md;456/384/72 gleich.
Fortsetzung: weitere belegbare Eigenschaften und allgemeine Entwurfsanforderungen
fortlaufend umsetzen; keine Zwischenabschlüsse oder Fertiganrufe.

Jet4 Required/AllowZeroLength werden jetzt im vorhandenen lesenden Metadatenlauf
als ausdrücklich vorhandene boolesche Properties übernommen. Tatsächliche
MDBTools-Repräsentation yes/no gegen unabhängig wieder geöffneten technischen
Jet4-Behälter geprüft; fehlende Properties bleiben unbekannt. Arbeitskopien
setzen NOT NULL und Text-/Memo-Leerstring-CHECK tatsächlich durch; DAO liest
Originalwerte readonly. GUI verwendet ausdrückliches nativeOwned statt aus
bekannten Getterwerten auf Bearbeitbarkeit zu schließen; eine Metadatenabfrage
pro Feldliste, importierte Entwurfsänderungen deaktiviert. Keine Originalmutation.
Technische Importprüfung1/1 (0,16s), anschließend bestehende DAO/GUI/echte
Kassa-Tabellenkopie3/3 (95,41s; davon Quellkopie87,28s); App/Helfer gebaut.
SOURCE_FIELD_DISPLAY_PROPERTIES.md. Keine neue Testregistrierung oder F-ID;
456/384/72 bleiben unverändert. Fortsetzung: eigene DAO-QueryDefs dauerhaft
umbenennen, Katalog-/View-/Generationseigenschaften mit berücksichtigen.

Eigene persistente DAO-QueryDef.Name lässt sich jetzt atomar umbenennen, auch bei
reiner Groß-/Kleinschreibungsänderung. Gesamte gespeicherte Definition bleibt
bytegleich; SQL, Parameterbindungen und aufrufereigene Field-/Parameterreferenzen
bleiben erhalten. Katalogcache folgt dem neuen Namen, verwaiste alte Einträge
geben nach Neuanlage nicht versehentlich die umbenannte Abfrage zurück. Andere
alte Namenshandles werden über Namespace-Generation3420 ungültig. Vorhandener
View-Installer läuft im selben Savepoint; Fehler stellt SQL und fünf Registrywerte
zurück. Kollision/ungültiger Name/readonly/aktive äußere Transaktion verweigert;
importierte/temporäre Namen und automatische Referenzkorrektur bleiben offen.
DAO1/1 (2,52s), App gebaut; keine Quellkopie-/Gesamtsuite unnötig wiederholt.
DAO_QUERYDEFS_CATALOG.md; F055/F121 offen,456/384/72 unverändert. Weiterarbeiten.

Datenbankfenster bietet jetzt Ansicht→Abfrage umbenennen für eigene Abfragen.
Offene Abfragedatenblätter vorher schließen; importierte Abfragen bleiben
unverändert/geschützt. Live-QueryDef wird über Namensdialog gehalten, damit
zwischenzeitliches Löschen/Neuanlegen nicht auf ein anderes Objekt umleitet.
Explorer/Selektion folgen tatsächlichem Namen. Tabellen-/Abfrageumbenennung
nutzen gemeinsamen Core-Helfer für vorhandene Druckseiteneinstellungen:
Schema geprüft, Layout unverändert verschoben, Zielkollision bricht gesamten
Savepoint ab. Kein QtGui im DAO-Backend. Bestehende DAO/GUI2/2 (7,84s) prüfen
auch tatsächliches Öffnen/Ausführen, Seitenbindung und atomare Teilfehler-Rücknahme.
App gebaut; langsame Originalkopie nicht ohne Importänderung wiederholt.
DAO_QUERYDEFS_CATALOG.md; F055/F080/F121 offen,456/384/72 unverändert. Weiterarbeiten.

DoCmd.Rename führt jetzt über tatsächliche eigene TableDef-/QueryDef.Name-Setter
Tabellen-/Abfrageumbenennung aus, einschließlich neu geordneter benannter
Argumente. ObjectType und OldName derzeit ausdrücklich erforderlich; Missing
bleibt auch bei Namen/Kommalücken Missing statt versehentlich acTable(Empty).
Null94, unbewiesener Kontext/Typ3251; geschlossene/beschäftigte/modale Projekte
und passende offene Dokumente werden geschützt. Aufrufer über skalare Auswertung
und Dispatch abgesichert. Reale VBA-Integrationsprüfung bestätigt beide Objekttypen,
resultierendes SQL, Seitenbindung und Fehlerfälle. DAO/GUI2/2 (7,77s), App gebaut.
Große DAO-/GUI-Prüfprogramme verwenden unter GNU/Clang Release jetzt -O1, um
Wiederübersetzungen zu verkürzen; Produkt-/Desktop-/Datenbibliotheken bleiben
nach tatsächlich geprüften flags.make bei -O3. Keine neue Testregistrierung.
DOCMD_RENAME_RUNTIME.md; F052/F055/F080/F121 offen,456/384/72 unverändert.
Fortsetzung: tatsächliches Entfernen eigener Datenbankobjekte samt abgeleiteten
Ansichten und Seitenmetadaten prüfen/umsetzen; kontinuierlich weiterarbeiten.

Eigene Tabellen/Abfragen können jetzt über Ansicht→Ausgewähltes Objekt löschen
(mit Nein als Vorgabe und über Dialog gehaltenem DAO-Handle) oder ausdrückliches
DoCmd.DeleteObject entfernt werden. Benannte Argumente/Missing/Null und passende
offene Dokumente geschützt; importierte Objekte im Fensterpfad deaktiviert.
QueryDefs.Delete entfernt Definition, vorhandene Seitenmetadaten und verwaltete
Views im selben Savepoint; Fehler stellt fünf Registrywerte wieder her, eigene
Workspace-Rücknahme bleibt erhalten. Native DROP/TableDefs.Delete räumt ebenso
Seitenmetadaten auf, damit Neuanlage keine Altzustände erbt. Unbekannte Katalogform
bricht gesamte Mutation ab; Roh-SQL-Rename mit Seitenzustand verweigert ohne
Rebinding. Keine automatische Löschung/Umschreibung abhängiger gespeicherter
Bindings; Relationenschutz bleibt. Bestehende DAO/GUI2/2 (8,06s), App gebaut:
echter VBA-Aufruf, Offen-/Missing-/Null-Schutz, GUI Nein/Ja, Views/Seiten/Objekte
wirklich entfernt, Namenswiederverwendung, Teilfehlerrücknahme, andere Daten gleich.
DELETE_OBJECT_RUNTIME.md; F052/F055/F080/F121 offen,456/384/72 unverändert.
Fortsetzung: vorhandene echte Access-Eval-Beobachtungen für allgemeine typisierte
Ausdrücke nutzen; VBA-/Eval-Kontext unterscheiden, keine pauschale Oracle-Parität.

Native Decimal-Variants (VarType14/CDec) implementiert: geprüfter96-Bit-Koeffizient,
Skala0..28, exakte Dezimalarithmetik/-Vergleiche, Vorzeichen/Abs/Fix/Int/Round,
Ganzzahl- und Currencykonvertierung ohne Double-Zwischenrunde; Variant-Arrays,
einmalige Skalarauflösung, homogene Berichtssummen/-Extrema/-Selektoren sowie
verlustfreie ungebundene Steuerelemente. Boost.Multiprecision nur Header>=1.74;
Buildabhängigkeit README/Workflow ergänzt, keine neue Shared-Library/Java.
Normales VBA, berechnete Controls und tatsächliches Application/Host.Eval bleiben
getrennte Kontexte. Eval übernimmt beobachtete Long-/Integer-/Decimal-Literale,
Typen, Präzedenz/Nz und Fehlerresultate. Acht Fehlerausdrücke über exaktenSHA256-
Fallabgleich mit sämtlichen112 technischen Collector-Ausdrücken wiedergewonnen;
keine Eingaben erfunden. CDec in beobachtetem Eval-Kontext verweigert2439,
Switch/Error-Variant6; gewöhnliche VBA-CDec-Konvertierung bleibt verfügbar.
Bereinigte echte Access11.0/DAO3.6/de-AT-Beobachtungen:104Werte plus8zugeordnete
Fehler, ohne private Quellmetadaten/Pfade/VBA. Bestehender DAO-Prüflauf bestätigt
56Zahlen-/Typbeobachtungen und8Fehler;48andere Werte bleiben ausdrücklich offen.
Core/DAO/GUI3/3 (9,49s), App/Backend3.54.1/JOIN/NULL/Counter geprüft; GUI-freie
Konfiguration erfolgreich, gelaufener Core-Binärpfad lautldd ohneQt/Boost-DSO.
Kein neuer CTest-Punkt/Katalogpunkt, keine neue CI-/Matrix-/Vollparitätsbehauptung.
NATIVE_DECIMAL_RUNTIME.md; exakte dbDecimal-Felder/Jet-SQL-Speicherung, vollständige
Locale/IEEE-/Automation-Identität/AST weiteroffen,456/384/72 unverändert.
Fortsetzung: die48weiteren echten Ausdrucksbeobachtungen und allgemeine
Sprach-/Datenbankanforderungen fortlaufend umsetzen; keine Zwischenabschlüsse.

Allgemeine Formatfunktionen erweitert: kombinierte Datum-/Zeitmasken, Monats-/
Wochentagsnamen, AM/PM, Literale/Datum-Escapes, feste/optionale Dezimalstellen,
Gruppierung/Prozent sowie FormatNumber/FormatCurrency mit benannten/ausgelassenen
Optionen und TriState. AktuelleQLocale liefert Zahlenkultur/regionale Erkennung;
Currency-Affixabstände werden als Legacy-Leerzeichen normalisiert, numerische
Gruppierungszeichen bleiben erhalten. MathematischeZahlenformatierung nutzt
Decimal, SQLiteINTEGER wird ausdrücklich ohne Double-Umweg gelesen (2^53+1 geprüft).
NativeSQL und Desktop teilen Formatter; reineCore-Hosts bekommen keinenQt-Zwang.
VBA.CDec unterstützt jetzt benanntesExpression-Argument. Eval-FormatNull liefert
beobachteten Leerstring, andere bestehendeVariant-/SQL-Nullregeln bleiben getrennt.
Gesamte vorhandene Ausdruckssuite abgeglichen:56Zahlen-/Typbeobachtungen+8Fehler
imDAO-Ziel sowie48weitereString-/Datums-/Locale-Beobachtungen über öffentliche
Application.Eval-Schnittstelle des tatsächlichenForm-/App-Objektmodells unterde-AT.
Insgesamt104beobachteteWerte+8Fehler; Doublefinanzwerte weitermitrelativerToleranz,
KalenderwertemitunabhängigerQt-Tages-/Sekundenberechnung. KeinepauschaleAST-/
OriginalVBA-/Locale-/Gesamtparität. Core/DAO/GUI/Formatter4/4 (9,49s), Appgebaut.
NATIVE_FORMAT_RUNTIME.md beschreibt Grenzen: komplexe/skalierende/wissenschaftliche/
Stringmasken, volleWindowsRegionaleinstellungen, negativeCurrencyprofile,
Originalfeld-Formatproperties und exakteNativeDecimal-/Currency-Datenbankfelder
weiteroffen. F014/F041/F068/F071/F118/F164;456/384/72 unverändert.
Kontinuierlich an allgemeinenDatenbankanforderungen weiterarbeiten, keinZwischenfinal.

Präzisionslücke in Str/Str$ geschlossen: Decimal-/Currency-Skalare werden einmal
aufgelöst und ohne Double-Zwischentext ausgegeben, einschließlich maximalem
96-Bit-Decimalwert und kleinster28erNachkommastelle. Currency-Konvertierung aus
Double wertet nun53-Bit-Mantisse/Exponent über ganzzahlige625/2er-Skalierung aus;
keine Abhängigkeit von plattformspezifischerlong-double-Präzision. Geprüft sind
exakte Halbwerte, direkt benachbarte IEEE-Werte, Vorzeichen, Subnormale sowie
signed64-Grenze/Nichtfinite/Überlauf. Die nachgewiesene REAL-Speichergrenze bleibt
unverändert; nativevolleCurrency-Datenbankfelder werden damit nicht vorgetäuscht.
Core/DAO/GUI/CurrencySQL/-Felder5/5 (10,03s), Appgebaut; bestehende112Ausdrucksfälle
weiterimDAO-/GUI-Prüfverbund. KeineOriginalkopie ohnebetroffeneImportänderung
wiederholt. Core-Prüfharness jetztwieDAO/GUI unterGNU/ClangRelease -O1; Runtime
lautflags.make weiter-O3. Keine neueTestregistrierung/F-ID;456/384/72 gleich.
Fortsetzung: Currency-Parameterwerte von begrenzterSQL-REAL-Kodierung trennen und
an allgemeinenexaktenDatenbanktypen fortlaufendweiterarbeiten.

Currency-Parameterwerte sind jetzt von der begrenzten SQL-REAL-Kodierung getrennt.
Value/defaultValue behält exakteCurrency-Werte einschließlich beider signed64-
Extrema undNull; ungültige Zuweisung ersetztvorigenWert nicht. ErstBind/Execute
prüft weiterdiebelegteREAL-Grenze undverweigertunsupportedWerte3251 ohnesieaus
demParameterzulöschen. Skalarstandardwerte werden einmalaufgelöst, danachwird
Generation/Lebensdauererneutgeprüft: einSQL-ändernderRückruf kannnichtdurchalte
ParameterinsneueSQLschreiben. VorhandeneParameter-/DAO2/2 (2,40s), Appgebaut;
Eingabe/Lesen/Überlauf-/SQL-Verweigerung/Null/Rückruf-/CloseundechteDatensätzegeprüft.
CURRENCY_PARAMETERS_RUNTIME.md. KeinexakterDatenbank-Currency-Speicherbehauptet;
456/384/72gleich. Fortsetzung: sichere reineCurrency-Parameterprojektionenohne
REAL-Umweg undallgemeineexakteDatenbanktypenweiterumsetzen.

Reine tabellenfreie Currency/Money-Parameterprojektionen liefern jetzt beide
signed64-Grenzwerte ohne REAL-Umweg: ganzzahlige Zehntausendstel mit eindeutigem
privatem Ergebnisdeskriptor. DAO behält Currency/Type5/Size8 und Null; Datenblatt
und Zwischenablage zeigen den exakten Betrag. Explizite Ausgabenamen verhindern
Verwechslungen mit SQLite-Parameternamen, auch bei Alias=Parametername.
SQL-Ausdrücke/FROM/WHERE/mehrdeutige Spaltennamen bleiben an der bisherigen
REAL-Grenze; veränderte Pläne sowie textuelle Filter/DAO-Find auf Tick-Ergebnissen
werden vor falschen Einheitenvergleichen verweigert. Bestehende Parameter-/DAO-
Prüfziele gemeinsam 2/2 bestanden (2,69s), Anwendung gebaut. Keine neue CTest-
Registrierung, kein neuer Katalogpunkt, keine native Currency-Felderstellung
oder allgemeine exakte Currency-SQL-Parität behauptet. 456/384/72 unverändert.
Details: CURRENCY_PARAMETERS_RUNTIME.md. Fortlaufend allgemeine native
Datenbank-/VBA-Anforderungen umsetzen; kein freiwilliger Zwischenabschluss.


Declared Currency For/Next counters now use exact signed64 ticks for the start,
frozen end/step and checked increment. Positive/negative four-decimal steps,
full-range boundaries, zero iterations, explicit zero step with Exit For,
body-written counters and the final post-loop value retain Currency kind.
Increment overflow raises6 before replacing the prior counter. Existing structured
loop stacks, Exit For and error recovery remain shared; no Double conversion is
used for this counter family. The existing execution budget still bounds zero-
step loops. Core/DAO checked together2/2 (3.99s), application built. Actual Access
2003 loop-oracle coverage, mixed Variant counter promotion and complete numeric
For/Next semantics remain open; no full parity assertion. Existing F041/F043/F045/
F105 only, fixed456/384/72 unchanged. Continue implementation without a final.

Native Zahlenmasken unterstützen jetzt bis zu drei Abschnitte für positive/
negative Zahlen und Zahlenwert0, leere Ersatzabschnitte, zitierte/escaped
Präfixe/Suffixe, negative Klammern und Prozent an beiden Seiten. Tausender-
Skalierung durch abschließende Kommata prüft exakte Rückrechnung und verweigert
Präzisionsverlust3251, statt still zu runden. .00 bleibt unterstützt. Gemeinsamer
Code für VBA/SQL/Formausdrücke; Masken-/SQL-Fälle und vorhandene GUI-Ausdrucks-
referenzen gemeinsam2/2 (6,01s), App gebaut. Keine zusätzlichen CTest-Ziele.
Wissenschaftliche/bedingte/Farb-/vierteNull-/interleavte Masken und vollständige
Windows-Formatgleichheit bleiben offen. Neue technische Prüfungen sind kein
zusätzlicher Windows-Access-Nachweis. NATIVE_FORMAT_RUNTIME.md; bestehende
F041/F068/F071/F118/F164,456/384/72 unverändert. Direkt weiterentwickeln.

FormatPercent ist jetzt als echte native VBA-/SQL-Funktion mit allen fünf
benannten/ausgelassenen Standardargumenten verfügbar. Die gemeinsame geprüfte
Maskenstrecke berücksichtigt Nachkommastellen, führendeNull, negative Klammern,
Gruppierung und lokale Prozentzeichen. Die benannte Format-Maske Percent liefert
zwei Nachkommastellen ohne Gruppierung. SQL-Integer bleiben vor der Skalierung
exakte Decimal-Werte; 9007199254740993 wird ohne Double-Verlust formatiert.
Null94, ungültige Optionen5 und Arity450 sind geprüft. Bestehender Formatter und
GUI gemeinsam2/2 (5,83s), App gebaut; native VBA-Positional-/Named-/Missing-Fälle
und vorhandene48 GUI-Referenzwerte enthalten. Keine neue Windows-Beobachtung
oder zusätzliche CTest-Registrierung. Decimal96/28-Grenze und vollständige
regionale Prozentprofile bleiben offen;456/384/72 unverändert.
NATIVE_FORMAT_RUNTIME.md; kontinuierlich allgemeine Datenbankfunktionen umsetzen.

Eigene native Tabellenfelder besitzen jetzt eine persistente DAO-Format-Text-
Eigenschaft mit derselben Detached-/Append-/Value-/Delete-/Identitätslogik wie
Caption/Description. Feldentwurf kann Masken setzen/löschen. Datenblatt zeigt
Formatwerte; SQL und EditRole bleiben unverändert. Zwischenablage behält den
bestehenden rohen typisierten Text. Druck/PDF nutzt eigene geprüfte formatierte
Werte: ungültige Masken zeigen #Fehler im Raster und brechen die Ausgabe vor
Erzeugung ab. Native Datum-/Zahlenmasken sind über echte DAO-/Designer-/Raster-
und PDF-Textextraktion geprüft, einschließlich beibehaltener SQL-Werte, Property-
Identität/Delete/Recreate und Rename/Schemaumbau. Fehlendes Include im GUI-
Prüfziel gezielt behoben; danach vorhandene DAO/GUI gemeinsam2/2 (8,00s), App
gebaut. Originale Format-Eigenschaften werden noch nicht als bekannt behandelt
(3251 statt erfundener Abwesenheit); Query-Lineage und Formular-/Berichtserbung
bleiben offen. NATIVE_FIELD_DISPLAY_PROPERTIES.md;456/384/72 unverändert.
Direkt als Nächstes proveniente Jet4-Formatübernahme mit technischer MDB prüfen.

Jet4-Format-Eigenschaften werden jetzt durch die bestehende schreibgeschützte
MDBTools-Grenze gelesen und mit eigener expliziter Format-Provenienz importiert.
DAO liest vorhandene Werte readonly383 bzw. meldet nachgewiesene Abwesenheit3270;
alte Metadaten ohne neue Format-Provenienz bleiben unbekannt3251. Datenblatt-
Formatierung verwendet diese tatsächlich gelesenen Masken, ohne Bearbeitungs-
oder SQL-Werte zu verändern. Die neue reine Entwicklungs-Javaquelle erzeugt eine
unabhängig wieder geöffnete technische Jet4-MDB mit null Zeilen/ohne VBA; nur die
importierte SQLite-Testkopie erhält Zahlen-/Datumsprüfwerte. Bestehende Import-/
DAO-Ziele gemeinsam2/2 (2,70s), zusätzlich frühere Caption-/Required-/AllowZeroLength-
Referenzdateien im selben Importziel; Herkunft, Typen, readonly, bekannte Abwesenheit,
Legacy unbekannt, ungültige Provenienz, Anzeige/Edit/Copy und unveränderter MDB-Hash
geprüft. App/Helper gebaut; vier Fixture-Hashes und drei Upload-Originalhashes OK.
Keine MDB oder Java-Abhängigkeit ins Produkt/Git übernommen. Keine Windows-
Render-/Vollparität;456/384/72 unverändert. SOURCE_FIELD_DISPLAY_PROPERTIES.md.
Kontinuierlich weitere allgemeine Datenbank-/VBA-Anforderungen umsetzen.

FormatDateTime unterstützt jetzt native VBA-/SQL-Aufrufe mit Date und optionalem
NamedFormat sowie vbGeneralDate/vbLongDate/vbShortDate/vbLongTime/vbShortTime.
Default und explizit ausgelassener NamedFormat wählen General Date; reine Zeit
und reines Datum werden getrennt dargestellt, sonst kombiniert. Lange Zeit nutzt
das tatsächliche QLocale-Profil (einschließlich US-PM), ohne Zeitzonentext; kurze
Zeit bleibt24-stündig. No-host-Syntaxprüfung darf geschachteltes CDate nicht
vorzeitig als Null konvertieren: gezielt behoben. CLDR-Schmal-/NBSP vor AM/PM
werden im nativen Legacy-Profil als Leerzeichen ausgegeben. Vorhandene Formatter/
GUI gemeinsam2/2 (5,57s), App gebaut; echte VBA Named/Missing/qualified-Fälle,
SQL1..2Argumente, Null94/ungültigeOption5/Arity450/ungültigesDatum13 und deutsche/
US-Zeitdarstellung geprüft. Keine zusätzliche CTest-Registrierung oder Windows-
Beobachtung. Volle regionale Datums-/Subsekunden-/Windows-Gleichheit weiter offen.
NATIVE_FORMAT_RUNTIME.md;456/384/72 unverändert. Kontinuierlich weiterentwickeln.

TextBox.Format wird jetzt aus SaveAsText übernommen und für berechnete Formular-
Textfelder sowie Berichtstextfelder tatsächlich gerendert. Getrennte typisierte
Werte bleiben maßgeblich: Anzeige-Fehler ändern Value nicht; Currency/Datum und
Long-Untertypen bleiben auch beim Löschen einer Maske erhalten. Native Berichte
erben proveniente direkte Feldmasken aus dem Berichtsentwurf und drucken sie
wirklich ins PDF. Bestehende gespeicherte Überschreibungen bleiben unabhängig.
Runtime-Formatlesen/-setzen/-löschen unterstützt diese Kontexte; ungültige Masken
lassen vorherige Eigenschaft stehen. Ausgabe verweigert Formatfehler vor falschem
PDF-Erfolg. Bearbeitbare Formularfelder behalten Masken nur mit explizitem3251-
Hinweis; deren Fokus-/Eingabepipeline ist nicht vorgetäuscht. CONTROL_FORMAT_RUNTIME.md.
GUI-Ziel bestanden (5,70s); anschließend Layoutziel einzeln bestanden (11,56s).
Der breitere Layoutverbund enthielt noch die alte Currency-For-Verweigerung:
gezielt durch echten Schleifenlauf mit zwei SQL-Schreibwirkungen und exaktem
Currency-Endzähler ersetzt. Keine neue CTest-Registrierung, keine frische gesamte
Suite/Windows-Gleichheit. App gebaut;456/384/72 unverändert. Ohne Zwischenfinal
weiter an den allgemeinen offenen Anforderungen programmieren.

General Number und Scientific sind jetzt native benannte Zahlenformate. General
Number behält Currency-/Decimal-Text exakt, trimmt überflüssige Nachkommastellen
und gruppiert nicht. Double/Single verwenden native15/7-signifikante Stellen.
Wissenschaftliche Einzelmasken 0.00E+00 bzw. e/E und +/- unterstützen optionale
Nachkommastellen, Exponentenpolsterung und exakte halbe Decimal-Rundung samt
Mantissenübertrag. Auch endliche Double-Werte außerhalb Decimal96 (1e100 und
Subnormale) können so formatiert werden. Vorhandene Formatter/GUI gemeinsam2/2
(5,79s), App gebaut; Max-Currency ohne Double-Umweg, größter Decimal, sehr kleine/
große Double, Rundung, Locale und echte SQL-Integer geprüft. Bisherige gezielte
Ausgabefehlerfälle verwenden weiterhin unsupported Farbmasken statt nun gültiger
wissenschaftlicher Masken. Keine neue Windows-Beobachtung/CTest-ID; kombinierte
wissenschaftliche Abschnitts-/Literal-/Engineering-Masken, vollständige regionale
Coercion und Access-Gleichheit bleiben offen. NATIVE_FORMAT_RUNTIME.md;456/384/72
unverändert. Ohne Zwischenabschluss direkt weiterimplementieren.

Textmasken unterstützen jetzt @ (mit Leerstellen), & (ohne Leerstellen), ! für
Linksfüllung, native Groß-/Kleinschreibung mit >/<, zitierte/escaped Literale und
einen zweiten Abschnitt für leere Strings/Empty. Null-Konventionen bleiben
getrennt. Tabellenanzeige und geerbte Berichtsformate verwenden diese Strecke;
Value und rohe Zwischenablage behalten den gespeicherten Originaltext. Native
Formatter/GUI gemeinsam2/2 (5,72s), App gebaut; echte SQL-Fälle, Füllrichtungen,
überzählige Zeichen, Leerwerte, Tabellen-/Berichtsmaske bei unveränderten Daten,
Unicode und vor unzulässigem Surrogat-Split verweigerte Ausgabe geprüft. NUL/
überlange Ausgabe sowie längenändernde Case-Konvertierung mit Platzhaltern werden
explizit verweigert; vollständige Windows-/Unicode-/Maskengleichheit bleibt offen.
Keine neue Windows-Beobachtung oder CTest-ID. NATIVE_FORMAT_RUNTIME.md;
456/384/72 unverändert. Nächste Umsetzung: Fokus-/Eingabestrecke formatierter
bearbeitbarer Formulartextfelder, ohne Anzeigen als Daten zurückzuparsen.

Bearbeitbare Formular-TextBoxen nutzen jetzt echte Formatdarstellung außerhalb
des Fokus und kanonische, ungerundete Eingabetexte im Fokus. Maßgeblich bleiben
typisierte Werte; nur tatsächliche Bearbeitung verwirft die Wertkopie. Formate
werden nicht als Daten zurückgeparst. Reine Darstellungswechsel erzeugen keine
Change-/Dirty-/SQL-Wirkungen. Doppelte Bindungen erhalten unabhängig ihre Masken;
verschiedene Werte mit identischer gerundeter Anzeige werden dennoch gespeichert.
Native Formularerstellung übernimmt proveniente Feldmasken. Bestehende Validierung,
Value/Text-Puffer, Undo und AfterUpdate bleiben wirksam; Maskenänderungen während
laufender Eingabe/BeforeUpdate verweigern3251. Skalarstandardwerte werden einmal
aufgelöst und Owner/Beschreibbarkeit erneut geprüft, geschlossener Owner2450.
Vorhandene GUI/Layout gemeinsam2/2 (17,22s), App gebaut. Tatsächliche Qt-Eingabe,
negative Abbruchvalidierung ohne Datenverlust, Undo, präzise SQL-Änderungen,
Text-Großdarstellung ohne Wertänderung und vollständige Uhrzeit hinter Datumsmaske
geprüft. Zunächst fehlende explizite Textgröße im technischen DAO-Aufbau gezielt
korrigiert; keine zusätzliche CTest-Registrierung. Regionale Eingabemasken,
beliebiges reentrantes Fokus-/IME-/Windows-Verhalten und vollständiger Access-
Formulardruck bleiben offen. CONTROL_FORMAT_RUNTIME.md;456/384/72 unverändert.
Ohne Zwischenfinal weitere allgemeine offene Funktionen umsetzen.

Formatierte native Suche ist in Tabellen-, Abfrage- und Formularansichten sowie
DoCmd.FindRecord integriert. Wörtliche String-Suche verwendet tatsächliche
Delegate-/TextBox-Darstellung, getrennten Wiederholungszustand und bei doppelten
Bindungen die gewählte Steuerelementmaske. Formatierte Suche über alle sichtbaren
gebundenen Textfelder ist möglich; Passwortdarstellung liefert keine Klartext-
Treffer. Rohwertsuche bleibt getrennt; formatierte Ersetzung verweigert3251.
Zell-/Zeitgrenzen, Null-/Binärausschluss und vorhandene Validierung gelten.
Vorhandene GUI/Layout gemeinsam2/2 (16,95s), App gebaut; tatsächliche Navigation,
Datumsmasken, doppelte Bindungen, Allfeldsuche und Passwortanzeige geprüft.
Ein veralteter geschlossener Dialogzeiger im Prüffall wurde gezielt korrigiert.
Keine neue CTest-ID oder Windows-Beobachtung; volle Jet-/Windows-Suchgleichheit
bleibt offen. FORMATTED_FIND_RUNTIME.md;456/384/72 unverändert.
Ohne Zwischenabschluss nächste Umsetzung: Bereichseinfügen mit neuen Datensätzen.

Native Tabellen-Bereichseinfügung kann jetzt über das Ende vorhandener Datensätze
hinaus neue Entwurfszeilen anlegen. Vorher werden noch nicht geladene vorhandene
Zeilen abgeholt und alle Zellen/Typen vorgeprüft; AutoWert-/Binärziele bleiben
geschützt. Neue Zeilen nutzen vorhandene Defaults, Datensatz-Undo, Verwerfen und
transaktionales Speichern mit echter Schlüsselerfassung. SQL-Constraintfehler
rollen alle SQL-Änderungen zurück; Entwürfe bleiben korrigierbar. Bestehender
Currency-/Feld-/Clipboard-Prüflauf1/1 (0,44s), App gebaut; Typfehler ohne Teiländerung,
zusätzliche Zeilen, Undo/Verwerfen, SQL-Defaults, erzeugte Schlüssel und Unique-
Rollback geprüft. Technische Schutzprüfung auf explizites AUTOINCREMENT korrigiert;
manuelle INTEGER-Primärschlüssel sind keine behaupteten AutoWerte. Kein neuer
CTest. Noch kein Einfügen in leere Tabellen/gesonderter Tabellen-Anfügebefehl;
weitere F340-Anforderungen bleiben offen. CLIPBOARD_APPEND_RUNTIME.md;
456/384/72 unverändert. Direkt weiterimplementieren.

Leere native Tabellen akzeptieren jetzt Texteingabe aus der Zwischenablage ab der
ersten sichtbaren beschreibbaren Spalte. Ein ausdrücklicher Kontextmenübefehl
und RunCommand acCmdPasteAppend(38) legen neue Datensatzentwürfe nach dem vollständig
geladenen Tabellenende an. Normales Einfügen behält die aktuelle Zelle als Ziel;
AutoWerte werden nicht überschrieben. Typisierte Formular-MIME wird ausdrücklich
verweigert, statt ihren Textkopf als Datensatz einzufügen. Vorhandene GUI/Feld-
Prüfungen gemeinsam2/2 (6,35s), App gebaut: leere Tabelle, Anfügen ohne Überschreiben,
SQL-Persistenz und tatsächlicher Hauptfenster-Dispatcher geprüft. F340 bleibt wegen
weiterer MIME-/Regional-/Undo-/Entwurfsszenarien offen. CLIPBOARD_APPEND_RUNTIME.md;
456/384/72 unverändert. Direkt weitere Clipboard-Teilanforderungen umsetzen.

Native Tabellen-Zwischenablage überträgt jetzt zusätzlich zum TSV eine typisierte,
begrenzte Rechteck-MIME. NULL, leerer Text, leere Binärwerte, Binärbytes, volle
SQLite-64-Bit-Ganzzahlen und vorhandene Currency-/Datumsdarstellung bleiben getrennt.
Binärdaten werden nur über validierte native Nutzdaten in Binärspalten geschrieben;
Text-/AutoWert-/Readonly-/verdeckte Zielgrenzen bleiben geschützt. Fehlerhafte MIME
wird nicht still als TSV interpretiert. Tabellen ohne SQL-Defaults benötigen keine
unnötige DAO-Rekonstruktion ihrer BIGINT-/BLOB-Felder. Für einen unveränderten
ungespeicherten Einfügebereich bestehen jetzt gemeinsame Strg+Z/Y-, Bearbeiten-Menü-
und Undo/Redo(292/199)-Wege; gespeicherte Date-Defaults werden beim Redo übernommen.
Spätere Fremdänderung, Requery oder erfolgreicher Save verwirft den Bereichsverlauf.
Vorhandene GUI/Feld/DAO gemeinsam3/3 (8,59s), App gebaut. Echte SQL-/Qt-Nachweise für
Bytes inklusive leerem BLOB, NULL/leer, große 64-Bit-Ganzzahlen, datumhaltige
Bereiche, Anfügen, leere Tabelle, Undo/Redo und Verlaufverfall; ungültige MIME,
AutoWert und verdeckte Ziele verweigert. Zwei veraltete Binär-Copy-Negativannahmen
angepasst; Qt-interner Schreibschutz während geprüfter Binärübernahme begrenzt
freigegeben, außerhalb weiterhin wirksam. Keine neue CTest-ID oder Windows-
Beobachtung. Kein VBA-Binary-Variant, keine allgemeine SQL-Rücknahme; Einfügen über
fremde Entwürfe, volle Regional-/Access-/MIME-Gleichheit bleiben offen. F340/F168/
F016 unverändert; TYPED_GRID_CLIPBOARD.md;456/384/72 unverändert. Direkt weiter.

Kreuztabellen mit TRANSFORM Count(Feld)/Count(*) und festen Textüberschriften in
PIVOT ... IN(...) laufen jetzt über denselben nativen Compiler, direkte DAO-
Recordsets, parametrisierte QueryDefs und gespeicherte Abfragen. Zählungen lassen
NULL-Feldwerte aus, Count(*) zählt Zeilen; fehlende Kategorien ergeben0. Reihenfolge,
Aliases und Fallkollisionsschutz der Ausgabespalten werden erhalten, Count-Spalten
sind provenient INTEGER4/DAO Long, QueryDef.Type ist dbQCrosstab16. Simple gruppierte
Zeilenfelder, einzelner Quellname mit optionalem Alias sowie WHERE/HAVING/ORDER
bleiben begrenzt; numerische Aggregate, dynamische/numerische Überschriften und
allgemeine Join-/Ausdrucks-Kreuztabellen verweigern3251. Vorhandene Query-Resultat/
Schema-/DAO-Prüfungen gemeinsam3/3 (2,44s), App gebaut. Zusätzlich vorhandene SQL-
Dokument-/DAO-Prüfungen2/2 (2,37s); keine neue CTest-ID. Klammer-/Nichtklammernamen
werden semantisch gleich verglichen. Kein Windows-Ergebnis-/Kollationsorakel und
keine volle Kreuztabellenparität behauptet. Siehe COUNT_CROSSTAB_RUNTIME.md.
F068/F067 und456/384/72 bleiben unverändert. Vier Repository-Originale und drei
Upload-Originale aktuell per SHA unverändert. Ohne Zwischenabschluss fortsetzen.

Feste Count-Kreuztabellen unterstützen jetzt zusätzlich vorzeichenbehaftete
Long-Überschriften und reine Year(Feld)/Month(Feld)/Day(Feld)-Pivot-Ausdrücke.
Integer-Spaltennamen werden kanonisch gebildet;01/1 und Text-/Zahl-Namenskonflikte
werden vor Ausführung verweigert. Nichtintegrale, außerhalb Long liegende und
beliebige andere Pivot-Ausdrücke bleiben3251. Vorhandene Query-Resultat-/Schema-
Prüfungen gemeinsam2/2 (0,12s), App gebaut: echte Datumszeilen mit Uhrzeit,
Schalttag und NULL, vorgegebene Monats-/Jahresreihenfolge, fehlende Kategorie,
Signum und Negativfälle geprüft. Keine neue CTest-ID/Windows-Beobachtung;
COUNT_CROSSTAB_RUNTIME.md; F068/F067 und456/384/72 unverändert. Direkt weiter.

Vollständige native Tabellenzeilen können jetzt über ausdrückliches Anfügen mit
neuen AutoWerten übernommen werden. Die Grid-MIME enthält physische Feldnamen;
Schemaform, Namen und Reihenfolge müssen vollständig passen. Nur provenient
generierte Spalten werden ausgelassen, manuelle Primärschlüssel bleiben echte
Werte mit ihren SQL-Constraints. Die neue Schlüsselvergabe erfolgt ausschließlich
beim normalen transaktionalen Save. Die Regel bleibt bei Bereichs-Undo/Redo erhalten.
Vorhandene GUI/Feld-Prüfungen gemeinsam2/2 (7,05s), App gebaut: vollständiger echter
Datensatz mit Binärbytes und großer Ganzzahl angefügt, Undo/Redo, tatsächlicher
neuer SQL-Schlüssel und verweigerte fremde Feldform geprüft. Positionsbereiche und
ältere Grid-MIME bleiben getrennt; anders geordnete/partielle Namenszuordnung und
Formular-MIME-Interop bleiben offen. TYPED_GRID_CLIPBOARD.md; F340/F168 und456/384/72
unverändert. Direkt weitere Bearbeitungsanforderungen umsetzen.

Saubere native Tabellen-Datenblätter unterstützen jetzt Ausschneiden sichtbarer
zusammenhängender Zellbereiche per Strg+X/Umschalt+Entf, Bearbeiten-/Kontextmenü und
RunCommand Cut189. Die typisierte Zwischenablage erhält die geprüften Quellwerte;
danach werden nur Cache-Zellen auf SQL NULL gesetzt. SQL bleibt bis zum normalen
transaktionalen Save unverändert. Dieselbe begrenzte Bereichs-Undo/Redo-Strecke
stellt die Werte wieder her. AutoWert-/Readonly-/verdeckte Ziele und fremde
Entwürfe bleiben geschützt. Vorhandene GUI/Feld-Prüfungen gemeinsam2/2 (6,91s),
App gebaut: echtes Menü/Dispatcher/Qt-Tastatur, Binär-Cut ohne frühe SQL-Änderung,
Undo/Redo, persistentes NULL statt leerem Text und verweigerter AutoWert-Cut bei
unveränderter Zwischenablage geprüft. Keine neue CTest-ID oder Windows-Beobachtung.
Dies ist Zellbereichsleeren mit Undo, keine Datensatzlöschung oder volle Access-
Cut-/Ereignisparität. TYPED_GRID_CLIPBOARD.md; F340/F168 und456/384/72 unverändert.
Ohne Zwischenabschluss weiterimplementieren.

Bereichseinfügen kann jetzt vorhandene ungespeicherte Zelländerungen physischer
Tabellenzeilen erhalten. Vorprüfung erfasst rohe Werte und geänderte Feldpositionen;
Undo stellt diese vorherigen Entwürfe wieder her, Redo wendet nur den neuen Bereich
an. Fehlgeschlagene Bereichsänderungen nutzen dieselbe Cache-Rücknahme, ohne SQL-
Schreibvorgang oder automatisches Speichern. Vorhandene GUI/Feld-Prüfungen gemeinsam
2/2 (6,28s), App gebaut: aufeinanderfolgende echte VBA-/UI-Pastes, ungültige Eingabe
bei erhaltenem vorherigem Entwurf/Undo, Currencywerte, separater Textentwurf und
Binäränderung mit leerem BLOB vor/nach Undo/Redo geprüft. Veraltete globale Dirty-
Negativannahmen der bestehenden Prüfungen angepasst. Neue/gelöschte Datensatzentwürfe
und mehr als1M geladene Zellen bleiben3251; Cut benötigt weiter einen sauberen Cache.
Keine neue CTest-ID/Windows-Beobachtung. TYPED_GRID_CLIPBOARD.md; F340/F168 und
456/384/72 unverändert. Ohne Zwischenabschluss weiterimplementieren.

Bereichseinfügen erhält jetzt auch bestehende neue Datensatzentwürfe und vorgemerkte
Löschungen. Cache-Rücknahme rekonstruiert eingefügte Zeilen an ihren Positionen,
übernimmt ihre tatsächlich gesetzten Felder, lässt SQL-Defaults ungesetzt und
stellt Löschmarker wieder her. Direkte Ziele in gelöschten Zeilen bleiben2448.
Bereichs-Undo/Redo stellt nach dieser Rekonstruktion eine gültige aktuelle Zelle
wieder her. Vorhandene GUI/Feld-Prüfungen gemeinsam2/2 (6,19s), App gebaut: früherer
Insert plus nachfolgende Bereichseingabe, Qt-Ctrl+Z/Y bei erhaltenem Zellziel,
tatsächliche spätere SQL-Neuanlage einschließlich Default, vorgemerkte Löschung,
verweigerter gelöschter Zielbereich und gemeinsamer SQL-Save geprüft. Erste Strecke
2/2 (6,36s), danach Zellziel-/Defaultnachweis gezielt ergänzt; keine neue CTest-ID.
Kein vorzeitiger SQL-Write, keine neue Windows-Beobachtung; beliebige externe
Index-/Reentranz-/Konkurrenz-/Access-Ereignisgleichheit bleibt offen.
TYPED_GRID_CLIPBOARD.md; F340/F168 und456/384/72 unverändert. Direkt weiter.

Native einfache numerische For-Zähler unterstützen jetzt Byte und Single zusätzlich
zu Integer/Long/Double. Grenzen/Schrittweite werden einmal ausgewertet und nicht
auf den Zählertyp eingeengt; Start/Next werden regulär Let-zugewiesen, danach wird
der tatsächlich zugewiesene (ggf. Single-gerundete) Zähler verglichen. Integer-/Long-
Bruchgrenzen, Byte-Rückwärtsschritte und Step0 mit Exit For funktionieren; Overflow6
ersetzt den letzten angenommenen Byte-/Single-Wert nicht. Deklariertes Date ist
laut MS-VBAL kein numerischer Zählertyp und bleibt vor SQL-Präfix verweigert3251.
Quelle: veröffentlichte MS-VBAL v20260818 §5.4.2.3/5.5.1.2.1, PDF-SHA dokumentiert;
keine neue Windows-/Access2003-Beobachtung. Vorhandene Core/DAO/Layout gemeinsam3/3
(17,41s), App gebaut: Rundung vor Anfangsvergleich, Endtypen Single/Double,
Bruchgrenzen, eingefrorene Grenzen, Body-Zähleränderung, Zero-Step, Overflow und
tatsächliche SQL-Schleifenaufrufe. Bestehende globale Bruchgrenzen-/Single-/Byte-
Negativannahmen gezielt in reale positive Fälle geändert; keine neue CTest-ID.
Volle gemischte Variant-/Fixedpoint-/Fehler-/Compilerparität bleibt offen.
TYPED_FOR_RUNTIME.md; F045/F041/F122 und456/384/72 unverändert. Direkt weiter.

Tabellennavigation kann jetzt von einer sauberen aktuellen Zeile wechseln, ohne
unbeteiligte Cache-Entwürfe zu speichern oder zu verwerfen. Native Pfeiltasten und
GoToRecord/RunCommand nutzen denselben Pfad. Ein einzelner befüllter neuer Datensatz
kann vor dem Verlassen über den normalen transaktionalen Save mit echter
Schlüsselerfassung gespeichert werden; vollständig leere neue Entwürfe bleiben
begrenzt, neue Zeilen werden bereits für Record-Undo als pending erkannt.
Der Shared-Find-Guard erlaubt ebenfalls eine saubere aktuelle Zeile bei anderen
Entwürfen. Andere Guard-Aufrufer behalten ihre bisherigen Insert-Grenzen.
Vorhandene GUI/Feld gemeinsam2/2 (6,20s), App gebaut: tatsächliche Qt-/VBA-Navigation,
erhaltene mehrere Entwürfe, unverändertes SQL für unbeteiligte Änderungen,
verweigerter leerer Insert, echter einzelner Insert beim Wechsel mit neuem
SQL-Schlüssel; Mehrfach-Save-/Sort-/Requery-Sperren weiterhin geprüft. Keine neue
CTest-ID oder Windows-Beobachtung; TABLE_NAVIGATION_DRAFTS.md. Aktueller gelöschter
Datensatz, nötige Teil-Saves mehrerer Entwürfe, allgemeine Addition-Row-/Windows-
Ereignisse bleiben offen. F368/F168/F080 und456/384/72 unverändert. Direkt weiter.

Native Find-Dialoge von Tabellen, Abfragen und Formularen unterstützen jetzt einen
expliziten Platzhaltermodus, getrennt von wörtlicher Suche. DoCmd.FindRecord
String-Muster aktivieren ihn bei * ? # oder [. Syntax/Matching teilen den bestehenden
VBA-Like-Parser und UTF16-DP-Matcher, einschließlich Zeichenklassen/Negation,
Whole/Prefix/Substring-Grenzen, Case-Modus und gespeichertem FindNext-Zustand.
Formatierte Suche behält reale Delegate-/Kontrollmaske und Passwortdarstellung;
typisierte Gleichheit bleibt getrennt, Musterersetzung verweigert3251. Musterfehler93
und absolute Patternlänge werden vor Navigation/Save geprüft; pro Kandidat gelten
1M Input×Pattern-Units,16384 Pattern-Units und Scan-Zell-/Zeitgrenzen. Bestehende
Navigation kann beim Matchscan den angenommenen aktuellen Datensatz speichern.
Vorhandene Core/GUI/Layout gemeinsam3/3 (19,13s), App gebaut: tatsächliche VBA-/Qt-
Muster, Ziffern/Klassen/Negation, Text-/Datumsmasken, Literal-Umschaltung, Query,
Fehler-/Größengrenzen und unveränderte SQL-Daten geprüft. Zunächst kollidierenden
lokalen Testvariablennamen korrigiert; veraltete globale Wildcard-Negativannahme
gezielt ersetzt. Keine neue CTest-ID/Windows-Beobachtung. Nicht-ASCII-Textranges/
Surrogatvergleiche und volle Jet-/Locale-/Allfeld-/Error/Event-Gleichheit bleiben
offen. WILDCARD_FIND_RUNTIME.md; F373/F071/F130 und456/384/72 unverändert. Direkt weiter.

Unformatierte DoCmd.FindRecord-Suche unterstützt jetzt OnlyCurrentField:=acAll0
über die jeweils passenden unterstützten Feldtypen. Text sucht Textfelder,
native Number/Currency/Date/Boolean verwenden ihre vorhandene geprüfte Gleichheit;
typisierte Werte benötigen weiter Whole-Field-Match. In direkten Tabellenformularen
ist der gerade fokussierte Steuerelementtyp im Allfeldmodus nicht mehr maßgeblich.
Tabellen/Readonly-Query-Resultate verwenden den bestehenden Allspalten-Scanner,
Formulare ihren begrenzten Datensatzbestand; formatierte Allfeld-Strings bleiben
getrennt. Vorhandene GUI-Prüfung1/1 (5,90s), App gebaut: tatsächliche VBA-/Qt-
Textsuche vom fokussierten Zahlenfeld, numerische Suche aus Textfokus, vollständiges
Datum mit Uhrzeit sowie Tabelle mit gefundenem Zelltyp/Spaltenwechsel geprüft.
Keine neuen CTest-/Windows-Beobachtungen; keine unbekannten Typen aus Zelltext
erraten. Volle Cross-Type-Coercion, Allfeld-/Spalten-/Fokus-/Ereignis-/Locale-
Gleichheit bleibt offen. ALL_FIELD_FIND_RUNTIME.md; F373/F071/F130/F050 und456/384/72
unverändert. Ohne Zwischenabschluss weitere allgemeine Funktionen umsetzen.

Wörtliches Tabellen-Ersetzen unterstützt jetzt Whole-/Teiltext-/Präfixmodus.
Current ersetzt das erste Vorkommen im Entwurf, Alle ersetzt alle nichtüberlappenden
Vorkommen jedes passenden Feldes bzw. einen Präfix/Whole-Wert. UI übergibt den
gewählten Vergleichsmodus tatsächlich; formatierte/wildcard/typisierte Rückersetzung
bleibt verweigert. Leerer Ersetzungstext ist echte Leerzeichenfolge, kein SQL NULL.
Batch-Ausgaben werden vor Cache-Änderung vollständig vorbereitet,16Mi UTF16-Units/
1M Zeilen/5s-Scan begrenzt und gemeinsam transaktional geschrieben; SQL-Fehler rollen
alle Änderungen zurück und hinterlassen keine Batch-Entwürfe. Vorhandene GUI/Feld
2/2 (6,26s), App gebaut: echte Dialoge, Case-insensitive Mehrfachvorkommen,
Prefix/current-erstes-Vorkommen, getrennte SQL-Save-Grenze, leere Textwerte,
pro Zeile verschiedene Ausgaben, vollständiger Check-Constraint-Rollback und NUL-
Verweigerung geprüft. Keine neue CTest-ID oder Windows-Beobachtung.
LITERAL_REPLACE_RUNTIME.md; volle Match-Offset-/Highlight-/Wildcard-/Locale-/Event-
Gleichheit bleibt offen; F373/F340/F168 und456/384/72 unverändert. Direkt weiter.

Explizites natives Datensatz-Anfügen ordnet Grid-MIME-Felder jetzt nach eindeutigen
Namen zu, auch bei anderer Reihenfolge und Teilmengen. Ausgelassene Felder bleiben
ungesetzt und erhalten SQL-Defaults; ausdrücklich geliefertes NULL überschreibt
Defaults. Nur proveniente AutoWerte werden ausgelassen. Name-/Formprüfung und
Expandierungsgrenze erfolgen vor Cacheänderung; die Supplied-/Default-Regel bleibt
bei Undo/Redo erhalten. Normales Paste/alte namenlose MIME bleiben positionsbezogen.
Vorhandene GUI/Feld gemeinsam2/2 (6,21s), App gebaut; zusätzlich Default/NULL-Nachweis
gezielt1/1 (0,48s): anders geordnete Zieltabelle mit zusätzlichem Defaultfeld,
AutoWert an anderer Position und echter Folge-ID11/12, Binärbytes, voller SQLite-
Integer, vollständiges Datum, Teilrecord mit ausgelassenen NULL-Feldern, explizites
NULL bei existierendem Textdefault, Undo/Redo und verweigerte doppelte Namen.
Keine neue CTest-ID oder Windows-Beobachtung. TYPED_GRID_CLIPBOARD.md; externe/
Formular-MIME-/Access-Ereignisgleichheit bleibt offen. F340/F168/F173 und456/384/72
unverändert. Direkt weiterimplementieren.


Readonly-Abfragedatenblätter veröffentlichen jetzt neben TSV das vorhandene typisierte
Grid-MIME mit Feldnamen. Tatsächliche rohe Querywerte erhalten Binärbytes, NULL,
Leertext/leeres BLOB und volle SQLite-Ganzzahlen; TSV verwendet Hex-Binärtext und
exakte Ganzzahldarstellung. Proveniente CURRENCY_TICKS-Ergebnisse erhalten ein
getrenntes Currency-Tag mit kanonischen 64-Bit-Ticks, keine Umdeutung als Ganzzahl.
Paste wandelt Currency vor Cacheänderung in exakten Text oder verlustfrei begrenzte
bestehende Currency/Double-Speicherung um; andere Zieltypen werden verweigert.
Keine physische CURRENCY_TICKS-Unterstützung oder externe Access-MIME-Parität.
Vorhandene GUI/Feld-Prüfungen gemeinsam2/2 (6,42s), App gebaut: tatsächliches
Query→anders geordnetes Tabellenziel, SQL-Default, volle Integer/Binär-/NULL-
Unterscheidung, beide Currency-64-Bit-Grenzen nach Text, kleine Currency nach REAL,
Undo/Redo/Save und vollständige Verweigerung großer Currency→REAL ohne Entwurf/
SQL-Änderung. Fehlendes QMimeData-Include beim Build gezielt korrigiert. Keine neue
CTest-ID/Windows-Beobachtung. TYPED_GRID_CLIPBOARD.md; F340/F168/F173/F014 und
456/384/72 unverändert. Direkt weiterimplementieren.


Bestehende aktuelle Tabellenzeilen können bei Navigation/Find jetzt einzeln
transaktional gespeichert werden, während andere Update-/Insert-/Delete-Entwürfe
im Qt-Cache unverändert bleiben. Bei mehreren Entwürfen sind geänderte/NULL-
Primärschlüssel und neue/gelöschte aktuelle Zeilen weiter verweigert; alleinige
populierte Inserts behalten den bisherigen Save-/Key-Pfad. Nach Commit aktualisiert
Qt selectRow nur die gespeicherte Zeile; Cache-Undo wird an der SQL-Grenze ungültig.
Vor Commit verschwundene/trigger-umgeschlüsselte Schlüssel führen zum Rollback.
GUI/Feld gemeinsam2/2 (6,21s), App gebaut: echte VBA-Navigation/Find, andere Entwürfe
und unveränderte Fremdzeilen, gemischte Update/Insert/Delete-Caches, Constraint-
Rollback und Trigger-Rekey-Rollback geprüft. Veraltete Find-Negativannahme und
Testversuch, ein geschütztes AutoWertfeld zu bearbeiten, gezielt korrigiert;
manuelles Text-PK separat geprüft. Uploads3/3/Fixtures4/4 hashgleich. Keine neue
CTest-ID/Windows-Beobachtung. TABLE_NAVIGATION_DRAFTS.md; vollständige aktuelle
Insert/Delete-/Fokus-/Ereignis-/Addition-Row-Parität bleibt offen. F368/F168/F080 und
456/384/72 unverändert. Direkt weiter.


„Neuer Datensatz“/acNewRec spart jetzt das pauschale Speichern fremder Entwürfe:
sauberer aktueller Datensatz darf bei weiteren Entwürfen eine neue Zeile öffnen;
bestehender geänderter aktueller Datensatz nutzt den gezielten Save-Pfad. Ein
alleiniger gefüllter Insert wird regulär gespeichert, bevor die nächste neue Zeile
entsteht. Leere aktuelle Inserts oder selektive Inserts mit weiteren Entwürfen
bleiben verweigert. Bestehende GUI-Prüfung1/1 (6,42s), App gebaut: tatsächliche
RunCommand/VBA-Aufrufe, SQL-Abgrenzung, anderer Entwurf, neuer Zeilenfokus und Escape
ohne fremden Datenverlust. Keine neue CTest-ID/Windows-Beobachtung. F368/F173/F168
und456/384/72 unverändert. Direkt weiter.


F088 hat jetzt einen ersten ausführbaren nativen Makroumfang: echter +Makro-
Aktionseditor (anlegen/entfernen/umsortieren/speichern/wiederöffnen), Explorergruppe,
DoCmd.RunMacro mit benannten Parametern und begrenzter Wiederholung. Acht Aktionen
verwenden vorhandene DoCmd-Pfade; StopMacro beendet einschließlich Wiederholungen.
Gesamte Struktur/Arity/Typen/Attribute werden vor erster Aktion geprüft, unbekannte
Bedingungen verweigert. Aktionen behalten ihre jeweiligen Transaktionsgrenzen;
späterer Laufzeitfehler rollt frühere erfolgreiche Aktionen nicht zurück. Diagnostik
bewahrt Fehlernummer/-metadaten mit Makroname/Aktionsnummer. AllMacros bietet den
bewachten vorhandenen Projektkatalog; SysCmd unterstützt nativen acMacro-Zustand.
GUI/Layout gemeinsam2/2 (17,16s), App gebaut: echter Editor/Reopen, gespeicherte
Definition, tatsächlicher benannter VBA-Aufruf mit dreifacher SQL-Wirkung,
StopMacro, präventiv verweigerte spätere unbekannte/bedingte/falsch typisierte
Aktionen, fehlendes Makro2485, bestehende SQL-Fehlernummer3078 mit Aktion2 und
bereits erfolgreich geschriebenen früheren Zeilen, Wiederholungsgrenze und
Katalog Count/Item/Enumeration/Name/IsLoaded geprüft. Anfänglich zu generische
SQL-Fehlernummer3146 im Test auf den tatsächlichen bestehenden3078-Pfad korrigiert.
Keine neue CTest-ID/Windows-Beobachtung. NATIVE_MACROS_RUNTIME.md beschreibt eigene
Format-/Budgetgrenzen; kein Access-Makroimport. Gruppen, Bedingungen/Ausdrücke,
RepeatExpression, Eventbindung, volle Aktionssprache und Oracle bleiben offen.
F088/F050/F057/F148 und456/384/72 unverändert. Direkt weiter programmieren.


Gespeicherte native Makronamen funktionieren jetzt auch als Ereignisbindungen:
vorhandene generische Steuerelementpfade (u.a. Click/DblClick/Change), Kontroll-
BeforeUpdate/AfterUpdate sowie Form-/Berichtsereignisse und Timer verwenden einen
geteilten Dispatcher für [Event Procedure], =Funktion() und benannte Makros.
OnActivate verschiebt Current bei jeder vorhandenen Bindung wie bisher bei VBA.
Makrofehler werden als echte Ereignisergebnisse/Statusdiagnose gespeichert;
fehlendes Kontroll-BeforeUpdate-Makro verhindert Save mit unveränderten SQL-Werten.
GUI/Layout gemeinsam2/2 (17,33s), App gebaut: tatsächliches Load/Current, Qt-Klick,
Doppelklick, Texttastatur/Change, Before/AfterUpdate mit SQL-Persistenz, tatsächliches
Timer-Timeoutsignal, Close sowie fehlender Click/BeforeUpdate-Name2485 nachgewiesen.
Timerprüfung prüft Signaldelivery, keinen neuen Wandzeit-/Access-Timing-Oracle.
Makronamen mit führendem =/[ sind zur Abgrenzung von Ausdrücken/Embedded-Makros
verweigert. Keine neue CTest-ID/Windows-Beobachtung. CancelEvent, Tastatur-
Makroargumente, Gruppen/Bedingungen/volle Ereignisgleichheit bleiben offen.
NATIVE_MACROS_RUNTIME.md; F088/F050/F057/F148 und456/384/72 unverändert. Direkt weiter.


Native Makros/DoCmd unterstützen jetzt CancelEvent ohne Argumente in unterstützten
abbrechbaren Bindungen (Open/Unload/NoData/Dirty/BeforeInsert/BeforeUpdate/DblClick).
Der verschachtelt restaurierte Ereigniszustand wird in VbaExecutionResult.cancelled
übernommen; Folgeaktionen laufen normal weiter. Außerhalb unterstützter Ereignisse
bleibt3251, zusätzliche Argumente450. GUI/Layout2/2 (17,08s) bestätigen echte
Kontroll-BeforeUpdate-Verweigerung ohne SQL-Save, Open-/Unload-Abbruch und keinen
über den Ereignisrahmen hinaus verbleibenden Cancel-Zustand. App gebaut; die neue
Designerzuordnung liegt separat noch als WIP vor.
Zusätzlich bestehender Original-Kassa-Lauf gezielt1/1 (114,46s) bestanden. Dabei
alte Rohtextannahmen für tatsächlich formatierte Original-Form-/Berichtsfelder und
Summen korrigiert: gemeinsame Prüfhilfe vergleicht reale skalare Werte mit SQL und
Anzeige separat über den bestehenden Formatter. SQLite-Ausdrucksmetadaten sind
kein voller Access-Varianttyp-Oracle. Eigenen anfänglich falschen Objektzugriff in
der Prüfhilfe behoben; keine Produktionsformatierung zur Erzwingung alter Erwartungen
abgeschaltet. Ungültige explizite WhereCondition meldet den tatsächlichen
Datenquellenfehler vor nachfolgenden Initialisierungsfolgefehlern; die bekannten
beiden fehlende-Verkaufs-ID-Fälle liefern wieder3075/Zeile1023 statt überlagertem3146.
Das ist Fehlerpriorität nach Konstruktion, keine Zusicherung ausbleibender früherer
Handlerwirkungen. Originalverkaufsworkflow bleibt offen. Uploads3/3 und Fixtures4/4
erneut hashgleich. Keine neue CTest-ID oder Windows-Beobachtung; kein frischer
Gesamtsuite-/Plattform-/CI-Nachweis. NATIVE_MACROS_RUNTIME.md; F088/F057/F114/F129/
F148 und456/384/72 unverändert. Direkt weiter.


Native Formulare können Makros jetzt über „Ereignisse zuordnen...“ an unterstützte
Form-/Feldereignisse binden. Automatisch/Makro/Funktion/[Event Procedure]/Keine
Bindung sind echte gespeicherte Alternativen; „Keine“ unterdrückt automatische
VBA-Erkennung. Unterdialog arbeitet auf Kopien, Cancel verwirft, OK validiert,
Eltern-Save persistiert. NativeStoredForm prüft/emittiert events/field_events/
timer_interval, unbekannte/doppelte Ereignisse/Felder und ungültige Werte werden
verweigert; keine Ausführung beim Save. Makronamenauflösung und Save-Konfliktprüfung
verwenden begrenzten Qt-Case-insensitive-Katalogscan (1M Namen/5s), danach exakten
Definitionslookup; Ä/ä und mehrdeutige Namen sind tatsächlich geprüft, keine
Jet-/Locale-Gleichheit behauptet. Eigenständiges VbaValue-Include korrigiert.
GUI/Layout zusammen2/2 (18,42s), App gebaut: echte verschachtelte Dialoge,
Cancel/Reopen/OK, gespeicherte Definition/Layout, Load-SQL-Wirkung, per Designer
gebundener BeforeUpdate-Abbruch mit unverändertem SQL, explizit deaktivierte
vorhandene Form_Load-Prozedur, unbekanntes Ereignis/Feld und negativer Timer.
Keine neue CTest-ID/Windows-Beobachtung. Originallauf1/1 (114,46s) ist der zuvor
geprüfte Stand, kein neuer Gesamtsuite-/Plattformnachweis. NATIVE_FORM_EVENT_DESIGNER.md;
F080/F088/F050/F057/F130/F148 und456/384/72 unverändert. Direkt weiter.


Native Makroaktionen unterstützen jetzt optionale Bedingungsausdrücke und
verschachtelte RunMacro-Aktionen. Der echte Aktionseditor speichert/verschiebt die
Bedingung mit der Aktion. Syntax/Typ/UTF16-/4096-Unit-Grenzen werden für die ganze
Definition vor ihren Aktionen geprüft; anschließend genau eine Auswertung pro
besuchter Aktion/Iteration im vorhandenen Eval-Kontext. False/Null überspringen.
Öffentliche Standardmodulfunktionen bleiben vom CodeBehind getrennt; anfangs im
falschen Kontext angelegte technische Prüffunktion in echten Standardmodul-Artefakt
verschoben, keine privaten/öffentlichen Zugriffsgrenzen aufgeweicht. Nested RunMacro
teilt16 aktive Aufrufe/10000 besuchte Aktionen, restauriert Budgets bei Fehlern;
StopMacro beendet nur den aktuellen Aufruf samt Wiederholungen, Eltern können
weiterlaufen. Vorherige SQL-Wirkungen werden bei späterem Laufzeitfehler nicht
pauschal zurückgerollt. RepeatExpression/Gruppen/volle Macro-Expr-/Event-Parität offen.
App gebaut; Layout1/1 (11,21s) im initialen Bündel; nach korrigiertem Modulkontext
nur betroffene GUI-Prüfung wiederholt1/1 (6,43s): echte Editorbedingung,
True/False/Null, exakt dreifache öffentliche Function-SQL-Wirkung, ungültige spätere
Bedingung ohne Prefix-SQL, Nested-Wiederholung, Child-Stop/Parent-Fortsetzung,
Rekursionsschutz/anschließende Recovery sowie echte bedingte Formularvalidierung
(blocked verhindert Save, allowed wird gespeichert). Keine neue CTest-ID/Windows-
Beobachtung, kein neuer Vollsuite-/Plattform-/CI-Nachweis.
Aktuelles .deb unter /workspace/work/native-conditional-macros/package/linaccess-
0.7.0-Linux.deb gebaut; vorhandene Paketprüfung bestätigt App-/MDB-Helfer-Build-IDs,
bytegleichen SQLite3.54.1-Backend, tatsächlichen Backend-Selbstcheck und erlaubten
Payload ohne Referenzdaten/Quellen. CPack-Preinstall-Build einmal mit Parallelität4
abgeschlossen, keine zusätzliche Gesamttestsuite ausgeführt. NATIVE_MACROS_RUNTIME.md;
F088/F041/F050/F057/F148 und456/384/72 unverändert. Direkt weiter programmieren.


Eigene Standardmodule sind jetzt echte native Projektobjekte: +Modul/Texteditor,
Ctrl+S ohne Schließen, Wiederöffnen/Bearbeiten, Syntaxinventar ohne Ausführungsclaim,
AllModules-/SysCmd-Metadaten und bestätigte Löschung nur eigener Module. Namen
werden nach Save gesperrt; Umbenennen/Classes/VBE/volle Sprache bleiben offen.
Version1/vba-Definitionen liegen getrennt in _la_objects, importierte VBA-Zeilen
werden nicht geändert. Gemeinsamer Sources-Leser versorgt native Aufrufe/Eval,
öffentliche Variablen und bestehende SQL-Registry, erhält Archiv-Latest-/Shadow-
Prüfungen und Quellenrevision/ROLLBACK-Kohärenz und verweigert Kollisionen statt
Auswahl. Bounds256 Module/16MiB Quellen/1M Katalogzeilen/5s Scan, eigene Quelle1MiB.
Löschen prüft Projektdentität/-generation nach Modalinteraktion und gespeicherte
Revision. Fehlende direkte Includes und gemischte auto-Typdeklaration korrigiert.
DAO1/1 (2,55s), Layout1/1 (11,36s) im Bündel bestanden. Nach tatsächlicher
Qt-Fokussetzung für Ctrl+S und normalem LinAccess-Abfragepfad zur initialen UDF-
Registrierung nur betroffene GUI-Prüfung erneut1/1 (6,23s), App gebaut: echte
Editor-/VBA-/SQL-Aufrufe, gemeinsamer Counter/Private-Helper, SQL-Private-Sperre,
vorbereitete Quellenänderung/Savepoint-Rücknahme, native/importierte Namens- und
TEMP-Schattenfehler, Löschung/alte Prepared-Verweigerung sowie unveränderte
vorhandene Archivbytes. Keine neue CTest-ID/Windows-Beobachtung, kein neuer
Gesamtsuite-/Plattformnachweis. Paketprüfung gehört zum vorherigen64abd69-Stand,
kein frisch paketierter Modulstand behauptet. NATIVE_STANDARD_MODULES_RUNTIME.md;
F042/F050/F068/F118/F130/F155 und456/384/72 unverändert. Direkt weiter.


Application.Run für unqualifizierte öffentliche Standardmodul-Functions/Subs
implementiert: Procedure/Arg1..Arg30 mit bestehender benannter Bindung,
typisierte Rückgabe bzw. Empty, optionale Parameter und geteilter Modulzustand.
Private/fehlende Ziele2517, Projektqualifikation3251, Arity450; Fehlernummer,
Source und Help bleiben erhalten. Eigener Run-Grenzwert32 liefert28 vor der
allgemeinen Interpreter-Stapelgrenze. Ausführungstiefe schützt laufenden Kontext.
Kein DoCmd.Run, keine neue COM-Automation oder ByRef-Aufrufer-Rückschreibung.
App gebaut; bestehendes mainwindow_documents1/1 (6,48s) bestanden. Nach falscher
neuer Arity-Testannahme450 statt449 und konkret gefundenem Zusammenspiel der
Rekursionsgrenzen nur betroffene GUI-Prüfung wiederholt. Keine neue CTest-ID,
Gesamtsuite/Windows-Beobachtung/Paketprüfung behauptet. NATIVE_APPLICATION_RUN.md;
F042/F050/F118/F148 und456/384/72 unverändert. Direkt weiter mit Makro-RunCode.


Native Makroaktion RunCode implementiert, keine DoCmd.RunCode-Methode:
ein einzelner unqualifizierter Function-Aufruf mit Klammern, Rückgabewert ignoriert;
Quellsyntax/äußere Aufrufform vor erster Aktion geprüft. Geschachtelte Argumente,
doppelte Stringquotes, Datums-/Bracket-Literale im Scanner; ASCII-Prozedurnamen,
64Ki UTF16 Ausdruckslimit. Im gebundenen Formular-/Berichtsereignis CodeBehind
vor Standardmodulen, danach restaurierter Ereigniskontext; Standard-Private/
fehlende Funktionen2517, Subs/Namenskonflikte3251, originale Fehlerdaten plus
Makroaktionskontext. GUI/Layout zusammen2/2 (18,17s) ohne Korrekturlauf, App gebaut:
wirkliche Funktionen/Seiteneffekte, verschachtelte Makros, Private-CodeBehind,
Kontextrestauration, verschachtelte/quotierte Argumente, Sub-/Private-Sperren,
struktureller Abbruch ohne vorangestellten SQL-Effekt und Fehler7001/Source.
Aktuelles DEB gebaut; erste Paketprüfung hatte im aktivierten SDK dessen älteres
SQLite im LD_LIBRARY_PATH und scheiterte3251. Nur Paketprüfung mit wie bei den
gezielten Tests vorangestelltem gebündeltem SQLite wiederholt:5-Datei-Payload,
App-/Helper-BuildIDs, MDBTools1.0.1, bytegleiches SQLite3.54.1 samt tatsächlichem
JOIN/NULL-/Counter-Backendcheck bestanden. Keine Wiederholung der GUI-Prüfungen.
Originalupload-Hashes3/3 und Fixtures4/4 unverändert. Keine neue CTest-ID,
Gesamtsuite/Windows-Beobachtung. NATIVE_MACROS_RUNTIME.md; F088/F042/F050/F057/F148
und456/384/72 unverändert. Weiterprogrammieren, kein Gesamtabschluss.


Eigene gespeicherte Standardmodule können im realen Editor mit Prozedurauswahl,
Ausführen/F5 direkt laufen. Neuer NativeStandardModuleHost nutzt tatsächliches
CurrentDb/DBEngine, gemeinsamen Modul-/Staticzustand und öffentliche Modulaufrufe,
bestehende Qt-Vergleichs-/Format-/MsgBox-/InputBox-Pfade, keinen Formularattrappen-
kontext. Me/current-record und fehlende Application-Objekte bleiben abgegrenzt.
Projekt-/Generationsprüfung bei Modulentry/DB/Dialog; gespeicherte Definition
gegen gemeinsamen Katalog geprüft, ungespeicherter Text verweigert. Auswahl ist
transient und markiert Quelle nicht dirty; Edit deaktiviert Run bis Analyze/Save.
Neue gemeinsame MsgBox-Brücke korrigiert zunächst gefundenen Linkagefehler der
anonymen Hilfsfunktion; fehlenden QTextBrowser-Testinclude ergänzt. App gebaut,
bestehendes GUI1/1 (7,20s) bestanden: tatsächliches F5, DAO-Effekt und zweites Modul,
mit Formular geteilter Counter44, Rückgabewert, MsgBox-Ja6, Arity450, Me438,
GoSub-Preflight3251 ohne Präfix-SQL und Unsaved-Sperre. Keine neue CTest-ID oder
Gesamtsuite/Windows-Beobachtung; aktuelles DEB gehört vorherigem d36ca9f-Stand.
NATIVE_STANDARD_MODULES_RUNTIME.md; F042/F043/F050/F060/F118/F148 und456/384/72
unverändert. Direkt weiter mit formularfreiem Application.Run/Eval-Kontext.


Formularfreier Modul-Anwendungskontext erweitert: echte projekt-/generations-
gebundene Application für Run/Eval/DBEngine; Domänenfunktionen lesen tatsächliche
DB. Gemeinsamer Qt-freier Run-Helfer erhält in beiden Hosts Arity, öffentliche
Auflösung, typisierte Rückgaben/Fehlerdaten und32-Level-Grenze. Variant-Argumente
bewahren Objektverweise (DAO.Field); skalare Zielparameter behalten native
Coercion-Grenzen. Schwacher Projektcache hält keine Editor-/Host-Referenz und
prüft Projektgeneration/GUI-Thread; gehaltene Application bleibt nach Editorclose
verwendbar, Generationswechsel3420. Volles Application-Modell/identische Objekt-
identität zu Formularhosts/implizite Objekt-Skalarkoerzierung bleiben offen.
App gebaut; vorhandene GUI/Layout2/2 (18,74s) bestanden: dynamischer zweiter
Modulaufruf mit DAO-Schreiben, Eval/echtes DCount, DAO.Field-Objektargument,
rekursives Run28, gehaltene Application nach Editorclose, Identität über neuen
Host, Fehler7001/Source und Generationssperre. Testzugriff auf private Eval-Methode
vor dem Lauf auf öffentliche Run-Rückgabe eines Access.Application-Funktionstyps
korrigiert. Keine neue CTest-ID/Gesamtsuite/Windows-Beobachtung/Paketprüfung.
NATIVE_APPLICATION_RUN.md; F042/F043/F050/F118/F148 und456/384/72 unverändert.
Direkt weiter mit formularfreier nativer Makroausführung.


Gespeicherte native Makros können über Ansicht/Ausgewähltes Makro ausführen direkt
vom Datenbankfenster laufen. Reeller Modul-/DAO-Anwendungshost mit DoCmd-Proxy,
kein verstecktes Formular. Makro-Budget/Repeat/Bedingungen/Stop/Fehlerkontext in
gemeinsamen Runner extrahiert;16 aktive/10000 besuchte Aktionen gelten auch über
Hostwechsel, VbaExecutionScope hält Stop über den gesamten Aufruf. RunSQL hat
geprüfte Savepoint-/UseTransaction-Pfade; RunCode öffentliche Functions. Tatsächliche
Tabellen/Abfragen über bestehende Dokument-APIs; eigene gespeicherte Normalformulare
und Berichtsvorschauen über bestehende unabhängige showDocument-Lebensdauer.
Aktives reales Formular erhält RunCommand/Requery; sonst tatsächliches Datenblatt.
Projektgeneration/Verbindung/GUI-Thread und Modalgrenzen bleiben geprüft; Menülauf
schützt laufende Projektoperationen. Formular-/Berichtsfilter/weitere Modi und
Archiv-only-Dokumente sind in diesem Host noch nicht implementiert. ReportView0
wird3251 abgewiesen; Vorschau benötigt2. CancelEvent ohne Ereignis3251.
App gebaut, GUI/Layout zusammen2/2 (19,15s) bestanden: tatsächliche Menüstarter-
Kette mit SQL/Function/verschachteltem StopMacro/Bedingungen, Table/Query/Form/
Report/Requery/RunCommand, OnLoad-SQL und Form-RecordCount3, Fehler in Aktion2 mit
persistiertem Vorläufereffekt, Rekursion3251 mit folgendem unabhängigen Erfolg,
SELECT-RunSQL2342, nichtgebundenes CancelEvent und echtes abgebrochenes Form_Open.
Fehlenden QStatusBar-Testinclude vor Lauf korrigiert; keine neue CTest-ID.
Aktuelles DEB mit4 parallelen Preinstall-Jobs gebaut; Payload/Berechtigungen,
App-/Helper-BuildIDs, MDBTools1.0.1 und bytegleiches SQLite3.54.1 samt tatsächlichem
JOIN/NULL-/Counter-Backendcheck bestanden. Keine Gesamtsuite/Windows-Beobachtung.
NATIVE_MACROS_RUNTIME.md; F088/F042/F050/F057/F114/F148 und456/384/72 unverändert.
Direkt weiter, keine Gesamtabschlussmeldung.


Native AutoExec beim tatsächlichen Projektöffnen implementiert: nach Schema/
QueryViews/Explorer-Initialisierung einmal pro neuem Öffnen, CI-Katalogauflösung;
nur gespeicherte native _la_objects-Makros, keine Archiv-only-Access-Makros.
Gehaltener Shift-Zustand überspringt Startmakro. Menü/Startup nutzen gemeinsamen
Projektstarter mit Generations-/Verbindungs-/Modal-/Ausführungsschutz und Refresh
auch nach Teilfehlern; neue Generation leert schwachen Modul-Application-Cache.
AutoExec-Fehler bleiben sichtbar, Projekt offen, bereits abgeschlossene Effekte
werden nicht erfunden zurückgerollt; Shift erlaubt nächsten Start ohne Wiederholung.
AllowBypassKey=False/Startup-Formular/Trust-/Access-Makroimport bleiben offen.
App gebaut; nur bestehende GUI1/1 (7,30s) bestanden. Erste Shift-Prüfung hatte
synthetische Taste ohne Modifier-Zustand; korrigiert auf echte Qt-Shift-KeyPress
mit ShiftModifier im geöffneten Dateidialog, nur betroffene GUI erneut. Reale
Dateiprojekte: SQL+RunCode+Table beim ersten/erneuten Öffnen je einmal, Shift ohne
Effekte, unveränderte Archiv-only-Makrobytes ohne Ausführung, Fehler2517 mit
persistiertem Vorläufer und erfolgreichem Shift-Neustart, gehaltene Application
nach tatsächlichem Projektwechsel3420. Keine neue CTest-ID/Windows-Beobachtung/
Gesamtsuite; Paket gehört vorherigem db892cc-Stand. NATIVE_AUTOEXEC_RUNTIME.md;
F088/F042/F050/F121/F148 und456/384/72 unverändert. Direkt weiter.


RunMacro RepeatExpression implementiert, in API und nativem Aktionseditor:
Bedingung vor jedem Versuch; Anzahl allein begrenzt, Bedingung allein bisFalse,
beide bis erster Grenze, beide ausgelassen einmal. Bool-/Zahlkonstanten über
native String-/Expression-Konvertierung, Textsyntax bis4096 UTF16 vor Effekten.
Ausgelassene Count-Position im nativen JSON ausschließlich für RunMacro als null
(Missing); Editor-Save/Reopen erhält leere Anzahl und Argument3. Ohne Count max
1000 abgeschlossene Wiederholungen; weiterer wahrer Versuch3251, abschließendes
False normal. Gemeinsame16-/10000-Aktionsbudgets und StopMacro unverändert.
Layout1/1 (11,44s) bestand im Bündel, GUI fand verlorene Fehlerquelle im alten
VbaHost::eval-Pfad. Erklärung/Source/HelpFile/HelpContext dort korrekt weitergereicht,
gezielt bestehende Kern-/GUI2/2 (9,32s) bestanden, App gebaut: reale DAO-Predicate-
Effekte3/Prüfungen4, Count+Expr, False0/TrueCount, Fehler7002 samt Wiederholungs-
kontext und separater Kern-Source/Help42-Nachweis, ungültige spätere Expr ohne
Präfixeffekt, nichtendender verschachtelter Noop an1000er-Grenze, danach UI-
Erstellung/CtrlS/Wiederöffnen und erneuter erfolgreicher Ablauf. Alte explizite
False-RepeatExpression-Verweigerung auf tatsächlichen erfolgreichen Null-Durchlauf
umgestellt. Keine neue CTest-ID/Gesamtsuite/Windows-Beobachtung/Paketprüfung;
aktuelles DEB gehört db892cc. NATIVE_MACROS_RUNTIME.md; F088/F041/F051/F118/F148
und456/384/72 unverändert. Direkt weiter mit nativen Makrogruppen.


Native Makrogruppen implementiert: eigenes Formatversion2 mit geordneten groups,
1..64 eindeutig CI benannten Teilmakros, insgesamt max256 Aktionen/1MiB; Version1
bleibt kompatibel. Gruppe startet erstes Teilmakro, Gruppe.Teilmakro gezielt;
Namens-/Callvalidierung und Formular-/Feldereignisdesigner akzeptieren diese Form.
Alle Teilmakros strukturell geprüft, bevor ausgewählter Teil Effekte ausführt;
fehlerhafte inaktive Gruppe blockiert, doppelte Unicode-CI-Namen3251, fehlender
Teil2485. Bestehender Runner hält Wiederholung/Stop/übergreifende Budgets.
Realer Editor ergänzt Gruppenmodus, Name/Auswahl/Hinzufügen/Entfernen/Umordnen,
behält rohe auch ungültige Zell-Drafts beim Wechsel; Auswahl/Render markieren
Quelle nicht dirty. Mehrteilige Gruppe wird nicht still in Einzelmakro verwandelt.
Projektgeneration vor Save geprüft, aktive Zellbearbeitung vor CtrlS-Serialisierung
committed; neue Draft-Serialisierung bewahrt RepeatCount-Missing-Hole. Keine
automatische Verweisrefaktorierung oder Access-SaveAsText-/Binär-Makroimport.
App gebaut; bestehende GUI/Layout2/2 (20,10s) ohne Korrekturlauf bestanden: echte
Gruppenerstellung/CtrlS/Wechsel/Wiederöffnen/Umordnung mit geänderter Default-
Ausführung, Unicode-Teilaufruf, verschachtelter Parent/StopMacro/Repeat, reales
Form_Load, ungültige inaktive Gruppe und doppelte Labels ohne Effekte. Aktuelles
DEB gebaut/geprüft:5-Datei-Payload, Rechte, App-/Helper-BuildIDs, MDBTools1.0.1,
bytegleiches SQLite3.54.1 mit tatsächlichem JOIN/NULL-/Counter-Backendcheck.
Keine neue CTest-ID/Gesamtsuite/Windows-Beobachtung. NATIVE_MACROS_RUNTIME.md;
F088/F050/F057/F118/F148/F155 und456/384/72 unverändert. Direkt weiter mit
Speicher-/Löschschutz für eigene Makroobjekte.


Eigene Makroobjekte erhalten transaktionalen Speicher-/Löschschutz: neue Namen
kollidieren CI mit vorhandenen nativen/importierten Makronamen3010 statt Replace;
gespeicherte Objektnamen sind im Editor readonly. Save und Delete vergleichen
exakte alte Definition/Name unter Savepoint, staleRevision3251 statt Lost Update.
Main-Katalog physisch/ohne TEMP-Schatten geprüft, zusätzliche main/temp Trigger
verweigert; Größenlimit1MiB vor Materialisierung. Lesen/Ändern importierter
Quellarchive bleibt ausgeschlossen. Eigene Makros löschbar mit Bestätigung im
Editor/Explorer; auch unlesbare eigene Definitionen über Explorer entfernbar.
Projektgeneration nach Modalfrage und Definition nach Antwort erneut geprüft.
DoCmd.DeleteObject acMacro4 für eigene gespeicherte Objekte über bestehende
Projekt-API ergänzt. Gemeinsame laufende MacroFrames sperren Ersetzen/Löschen
auch für Gruppen und über Formular-/Modulhostgrenzen; kein globaler Jet-Lockclaim.
Layout1/1 (11,74s) im Bündel bestanden. GUI fand alte vorgelagerte Case-only-
Namensprüfung; entfernt zugunsten zentralem3010-Pfad. Nur betroffene GUI erneut
1/1 (7,41s), App gebaut: echte zweite Verbindung und Revision während bestätigter
Löschfrage, Native-/Archiv-Kollision, No/Yes, TEMP-/Trigger-Verweigerung, Löschen
ungültiger eigener Definition, aktive3251-Sperre in beiden Hosts mit folgenden
SQL-Effekten, anschließendes DeleteObject und unveränderte Archivbytes. Original-
Uploads3/3 sowie Fixtures4/4 erneut hashgleich. Keine neue CTest-ID/Gesamtsuite/
Windows-Beobachtung/Paketprüfung; DEB gehört9c80f00. NATIVE_MACRO_STORAGE.md;
F088/F050/F121/F148/F155 und456/384/72 unverändert. Direkt weiter mit allgemeinen
Makro-Projektbefehlen und getrenntem Projekt-Ausführungsschutz.


Native Makroaktionen Close/DeleteObject/Rename ergänzt über reale bestehende
Projekt-APIs: eigene geschlossene Tabellen/Abfragen umbenennen/löschen, eigene
Makroobjekte löschen, offene Datenblätter sowie reale Form-/Berichtsdokumente
schließen. Keine Save=acSaveNo-Verwechslung mit Datensatz-Undo; bestehende
validierte Close-Pfade committen Recorddrafts und geben SQL-Modelle frei.
Macro-Projektstarter erhält eigene Ausführungstiefe statt SQL-Tiefe; legitime
Objektbefehle bleiben nutzbar. Projekt-/Anwendungsschließen und neue UI-Makro-
Reentry bleiben bei aktiven Frames blockiert, auch im direkten Formularhost.
Gemeinsamer Preview-Dispatch hält NativeExecutionDepth bis Befehlsende, damit
Close/Requery/RunCommand keine laufenden Dokumentaufrufe vorzeitig entsorgen.
Wiederverwendete gespeicherte Fenster wiederholen keine historischen Initfehler;
neue fehlgeschlagene Initialisierungen bleiben echte Fehler mit neuerDoc-Bereinigung.
App gebaut; bestehende GUI/Layout2/2 (18,92s) ohne Korrekturlauf bestanden:
Menü-Makrokette mit realem Table/Query-Rename, Open/Close/Drop und Makrolöschung,
aktuellem Query-DefaultClose, Pending-Draft43 trotz acSaveNo, echtem Form_Unload/
Report_Close und anschließendem SQL. Während MsgBox ist tatsächliches Window.close
abgewiesen bei gleicher Projektgeneration/offener DB; danach Weiterlauf. Bestehende
Form mit altem Load7003 nur aktiviert, neue Init7003 abgewiesen/bereinigt. Grenzen
bleiben explizite Objekttypen/Modi, native Ownership, aktive Model-/Makro-Leases,
keine automatische Verweisrefaktorierung oder vollständige Access-Close-Parität.
Keine neue CTest-ID/Gesamtsuite/Windows-Beobachtung/Paketprüfung; DEB gehört9c80f00.
NATIVE_MACRO_PROJECT_COMMANDS.md; F088/F050/F057/F114/F121/F135/F140/F148/F155
und456/384/72 unverändert. Kontinuierlich weiter.


SelectObject als native Makroaktion und formularfreier DoCmd-Member ergänzt:
wirklich offene Tabellen/Abfragen über bestehende Aktivierung, echte Form-/Report-
Instanzen über lebensdauergeschützten Preview-Dispatch, eigene Makro-/Modulobjekte
und Ordner im Datenbankfenster. Gleiche4/5-InDatabase-Erweiterung im Formularhost;
kein erfundenes Modulfenster, kein Öffnen eines geschlossenen Datenobjekts.
Explorer behält tatsächliche logische Auswahl/Archiv-ID bei Refresh derselben
Projektgeneration; neues Projekt übernimmt keine alte Auswahl. Damit bleibt die
von einem Makro gesetzte Auswahl nach dessen abschließendem Refresh erhalten.
Gewachsene NativeStandardModuleHost/DoCmd/Application-Implementierung aus Header
in eigene Forms-C++-Einheit verschoben (öffentliche schlanke Deklaration); zwei
mechanische Qualifizierungs-/Duplikatfehler vor Testlauf korrigiert. Keine Änderung
an Kern-Qt-Freiheit oder Python/Java-App-Abhängigkeiten. App gebaut, bestehende
GUI/Layout2/2 (19,03s) bestanden: reale Table/Query/Form/Report-Aktivierung,
Macro/Module-/Ordnerauswahl, Refresh und tatsächlicher Menü-Makro-Endzustand,
Formularhost4/5, fehlende/geschlossene/unsupported Ziele. Vorige Modul-/Makro-/
Speicher-/Startprüfungen im gleichen Ziel weiter enthalten. Aktuelles DEB gebaut
und Payload/Rechte/App-/Helper-BuildIDs/MDBTools1.0.1/bytegleiches SQLite3.54.1
samt realem JOIN/NULL-/Countercheck verifiziert. Keine neue CTest-ID/Gesamtsuite/
Windows-Beobachtung. NATIVE_MACRO_PROJECT_COMMANDS.md; F088/F050/F135/F140/F148
und456/384/72 unverändert. Direkt weiter mit StopAllMacros/Echo-Wiederherstellung.


StopAllMacros als echte parameterlose Makroaktion implementiert, kein erfundener
DoCmd-Member. Gemeinsamer Stop-Zustand beendet den aktuellen verschachtelten
Aufrufbaum, Wiederholungen und bedingungsgesteuerte Fortsetzungen; nach äußerstem
Unwind sind Zustand/Budgets/Frames wieder frei. Bereits laufender nativer VBA-
Functioncode wird nicht global abgebrochen; folgende Makroaktionen bleiben aber
beendet. Neue unabhängige Rootaufrufe funktionieren. Echo-Aktion und Echo-/
Application.Echo im formularfreien Host nutzen bestehende reale UI-Suppression;
StopAllMacros stellt implementiertes Echo wieder her. Vollständiger SetWarnings-
Zustand/Access-Stop-/Reentry-Scheduling bleiben offen, keine Rücksetzung eines
nicht implementierten Warnungszustands behauptet. App gebaut, vorhandene GUI/
Layout2/2 (19,55s) ohne Korrekturlauf bestanden: tatsächliche MsgBox bei ausgeschal-
teten Updates, verschachtelter SQL/RunCode/Stop beider Hosts, native VBA-Fortsetzung,
fehlende nachfolgende Aktionen/Wiederholungen, unabhängiger Folgestart, Stop aus
RepeatExpression, False-Stop ohne vorzeitige Echo-Restauration, benanntes
Application.Echo, ungültiger Status ohne UI-Mutation, kein DoCmd.StopAllMacros.
Keine neue CTest-ID/Gesamtsuite/Windows-Beobachtung/Paketprüfung; DEB gehört cb8931b.
NATIVE_MACROS_RUNTIME.md; F088/F050/F051/F059/F135/F148 und456/384/72 unverändert.
Direkt weiter mit vollständigem Aktionsargument-Editor/weiterer Aktionsausführung.


## Fortschreibung 2026-11-09: vollständige native Aktionsargumente / MessageBox
Argumente… bearbeitet sämtliche deklarierten Aktionsargumente; zusätzliche Werte
bleiben bei Teilmakrowechseln und Zeilenbewegungen erhalten, Abbrechen verändert
keinen Entwurf. MessageBox ist eine echte vierargumentige Makroaktion in beiden
Hosts: kontextgebundene =Auswertung, Signalton, benannter Typ, Titel, escaped
@-Abschnitte; kein erfundenes DoCmd.MessageBox. Native JSON-Typnamen sind kein
Access-Binärimport. Quellen-/Ergebnisgrenzen und bekannte Typen werden geprüft.
App gebaut; vorhandene GUI/Layout2/2 bestanden (19,30s), einschließlich echtem
Argumentdialog, vierter Titel, Abbruch, Teilmakrowechsel/Speichern, realen Dialogen
in beiden Hosts und unbekanntem Typ. Build-Include für vollständigen VbaValue
korrigiert, kein Testfehler/Korrekturlauf. Originaluploads3/3 und Fixturehashes4/4
unverändert. Keine neue CTest-ID/Gesamtsuite/Windows-Beobachtung/Paketprüfung.
NATIVE_MACROS_RUNTIME.md; F088/F050/F051/F059/F135/F148,456/384/72 unverändert.
Weiter: GoToRecord aus tatsächlichen Standardmodulen/Makros mit vier Argumenten
an vorhandene Daten-/Formularnavigation anbinden; mehrzeilige Argumente erhalten.


## Fortschreibung 2026-11-09: GoToRecord aus Modul/Makro, mehrzeilige Argumente
Standardmodule Application.DoCmd.GoToRecord und native Makroaktion unterstützen
vier optionale/namentliche Argumente, tatsächliche offene Tabellen/Abfragen und
sichtbare Projektformulare. Bestehende Save-/Ereignis-/Reihenfolge-/Grenzlogik wird
wiederverwendet; keine erzeugten Ersatzdokumente. Native JSON-null ist hier eine
Auslassung, direktes VBA-Null bleibt ein Laufzeitwert. Detailargumente bearbeiten
Text mehrzeilig ohne SQL-Zeilenverlust; Bedingung bleibt separat.
App gebaut; Layout1/1 (11,40s), nach Korrektur nur des neuen Prüffall-Abfragemodus
auf Nur-Lesen GUI1/1 (7,61s) bestanden. Geprüft: tatsächliche Cursorpositionen,
benannter VBA-Modulaufruf, vierargumentiges Makro beider Hosts, optionale Löcher,
Tabellen-/Abfrage-/Formularwechsel, Offsetfehler ohne Cursoränderung, geschlossene
Form verweigert, mehrzeiliger Argumentdialog. Erste GUI verweigerte korrekt die
nicht unterstützte bearbeitbare sortierte Projektionsabfrage; keine Lockerung.
Keine neue CTest-ID/Gesamtsuite/Windows-Beobachtung/Paketprüfung. Modale Designer
und allgemeine versteckte Formularsuche bleiben explizite Grenzen.
NATIVE_MACROS_RUNTIME.md; F178/F088/F052 und456/384/72 unverändert.
Direkt weiter: vollständige FindRecord-Anbindung aus Standardmodulen/Makros.


## Fortschreibung 2026-11-09: FindRecord aus Standardmodul/Makro
Application.DoCmd.FindRecord unterstützt sieben/namentliche Argumente und echte
aktive Projektformulare, Tabellen und Abfragen. Gemeinsame Argumentprüfung mit
bestehendem Formularhost; typisierte Werte/Platzhalter/Formatgrenzen bleiben
unverändert. Native Makroaktion hat literaltextiges FindWhat und optionale Löcher;
keine Behauptung beliebiger Access-Makroausdrücke/Binärdekodierung.
Tatsächliche Tabellenöffnung wählt nun eine erste Zelle, damit FindRecord ohne
vorherigen manuellen Klick ein Suchfeld hat. Datensatzmarkierung selektiert die
aktuelle ganze Zeile ausdrücklich unabhängig vom initialen Qt-Auswahlanker und
bewahrt die aktuelle Spalte. Vorhandene GUI deckte diesen Folgefehler auf.
App gebaut; Layout1/1 (11,23s), nach gezielten Korrekturen GUI1/1 (7,91s) bestanden.
Der neue Suchprüffall setzt seine Sortierung ausdrücklich: die vorherige Annahme
aufsteigender Standardsortierung war falsch, der Zahlen-Treffer bereits korrekt.
Geprüft: sieben tatsächliche Editorargumente/Speicherung, Zahlen/Currency,
Platzhalter/Weitersuche, benanntes Modul, Nur-Lesen-Abfrage, Formular/Makro beider
Hosts, Null-/Mustervalidierung ohne Cursoränderung und vorhandene Zeilenauswahl.
Keine neue CTest-ID/Gesamtsuite/Windows-Beobachtung/Paketprüfung; vollständige
No-Match-/Zeit-/Untertyp-Parität bleibt offen. NATIVE_MACROS_RUNTIME.md;
F088/F052/F178/F363 und456/384/72 unverändert. Direkt weiter mit tatsächlichen
Fokus-/Cursorbefehlen und gemeinsamer UI-Ausführungslebensdauer im Modulhost.


## Fortschreibung 2026-11-09: Fokus/Cursor und gemeinsame GUI-Ausführungsrahmen
GoToControl und Hourglass stehen nativen Makros und dem formularfreien DoCmd-Host
mit namentlichen Argumenten zur Verfügung. Fokus geht an echte Controls eines
aktiven Projektformulars; ohne Formular2046, kein erfundener Me-Kontext.
Modul-Eintritt/-Austritt und beide Makrohosts verwenden denselben nativen Qt-
Ausführungszähler. Wartezeiger bleibt über verschachtelte Funktionen/RunCode bis
zu späteren Aktionen erhalten und wird bei äußerstem normalem/fehlerhaftem Ende
zurückgesetzt; fremde Cursor bleiben geschützt. Echo bleibt separat persistent.
Vorhandene GUI/Layout2/2 (20,12s) bestanden, kein Korrekturlauf: benannter Fokus mit
Textauswahl, tatsächlicher Cursor in verschachtelten und nachfolgenden Dialogen
beider Makrohosts, StopMacro-Ende, Fehler7004, direkte Steuerung/fremder CrossCursor,
fehlender aktiver Formularkontext. App nach finalem Formularhost-Nachlauf über
CPack neu gelinkt. Frisches DEB native-module-ui/package/linaccess-0.7.0-Linux.deb:
Payload/Rechte/Build-ID/bytegleicheSQLite und tatsächlicher Backendcheck mit
MDBTools1.0.1/SQLite3.54.1 JOIN/NULL/Zähler bestanden. Paket ohne MDB/VBA/Fixtures.
Originaluploads3/3 und Fixturehashes4/4 unverändert. Keine neue CTest-ID/Gesamtsuite/
Windows-Beobachtung. NATIVE_MACROS_RUNTIME.md/NATIVE_HOURGLASS.md;
F088/F052/F068/F135/F148 und456/384/72 unverändert. Direkt weiter mit gemeinsamem
CurrentProject/CurrentData aus echten Standardmodulen und Projektlebensdauer.


## Fortschreibung 2026-11-09: gemeinsame CurrentProject-/CurrentData-Objekte
Standardmodule können Application.CurrentProject/CurrentData und die unqualifi-
zierten VBA-Bezeichner verwenden. Formular- und Modulhost nutzen dieselbe echte
Projektkatalogimplementierung und Objektidentität am tatsächlichen Projektfenster;
kein kopierter Katalog/kein erfundenes Formular. Cache zusätzlich generations-
gebunden. Vorhandene Datei-/Readonly-/Deskriptor-/IsLoaded-Grenzen bleiben bestehen.
App gebaut; vorhandene GUI/Layout2/2 (19,49s) ohne Korrekturlauf bestanden:
wirklicher VBA-Katalogzugriff/Counts, gemeinsame Identität in beiden Hosts,
Fortbestand nach Formularschluss, live Modul-Löschung, Argument/Readonly-Fehler,
3420 für gehaltene Projekt-/Datenkatalogobjekte nach tatsächlichem Projektwechsel
im vorhandenen AutoExec-Ablauf. Keine neue CTest-ID/Gesamtsuite/Windows-Beobachtung/
Paketprüfung; zuletzt geprüftes DEB gehört69ac1b9. NATIVE_STANDARD_MODULES_RUNTIME.md;
F052/F042/F059/F135 und456/384/72 unverändert. Direkt weiter: echte Forms-/Reports-
Sammlungen und Screen aus dem formularfreien Host, ohne Platzhalterformular.


## Fortschreibung 2026-11-09: tatsächliche Forms/Reports/Screen aus Standardmodulen
Formularfreier Host unterstützt Forms/Reports und Application.Forms/Reports mit
realen Dokumentobjekten, Count/Name/Index/ForEach, einschließlich nach Öffnen
versteckter Formulare. Application.Screen/bare Screen liefern tatsächliches
ActiveForm/ActiveControl/MousePointer; vorhandene Cursor-/Fokusgrenzen bleiben.
Gemeinsame Implementierung statt Platzhalterformular. Neue Projektobjekte sind
via echtem CurrentProject-Lebenszyklus geschützt; alte formulargebundene Handles
behalten ihren bestehenden Caller-Lebenszyklus. Screen-Suche zusätzlich an DB-
Identität gebunden und nach skalarer Wertauswertung erneut geprüft.
App gebaut; vorhandene GUI/Layout2/2 (20,16s) ohne Korrekturlauf bestanden:
echte sichtbare/versteckte Formulare und Bericht, index-/namensbasierte identische
Dokumentreferenzen, native VBA-ForEach über beide Sammlungen, reale aktive Form/
Control, VBA-MousePointer mit Rücksetzung, live Schließen/alte Dokumentreferenz,
3420 für gehaltene Sammlung/Screen nach tatsächlichem Projektwechsel.
Keine neue CTest-ID/Gesamtsuite/Windows-Beobachtung/Paketprüfung; DEB69ac1b9.
NATIVE_STANDARD_MODULES_RUNTIME.md; F052/F042/F059/F135 und456/384/72 unverändert.
Direkt weiter: globale projektgebundene Sammlungsidentität und tatsächlich
geöffnete/versteckte Formulare als benannte Navigations-/Schließziele.


## Fortschreibung 2026-11-09: globale Projekt-Sammlungsidentität/versteckte Ziele
Native Forms/Reports/Screen teilen einen schwachen, verbindungs-/datei-/generations-
gebundenen Cache am echten Projektfenster, auch über kurzlebige Application-
Wrapper hinweg. Projektwechsel leert den Cache; gehaltene Altobjekte bleiben
3420-geschützt. Tatsächlich geöffnete lebende Dokumente inkl. versteckter/modaler
Instanzen werden nach nativer Öffnungsreihenfolge erfasst; Schließen/Abbruch wird
berücksichtigt. Benanntes GoToRecord und Close erreichen nun echte versteckte
geöffnete Formulare, ohne sie zur Navigation sichtbar zu machen.
App gebaut; vorhandene Layout1/1 und korrigierte GUI1/1 (7,96s) bestanden. Der neue
Prüffall wurde berichtigt: Konstruktor dispatcht bereits Open/Load vor show();
eine erfolgreich geladene noch nicht angezeigte Instanz zählt deshalb richtig
als geöffnet. Kein Produktionsfehler verdeckt. Geprüft: direkte native VBA-Is-
Identität ohne gehaltenen Application-Wrapper, reale unangezeigte/ausgeblendete
Instanzen, versteckte Navigation ohne Anzeigen, benanntes Close und ungültig
gewordene Dokumentreferenz. Keine neue CTest-ID/Gesamtsuite/Windows-Beobachtung/
Paketprüfung; DEB69ac1b9. NATIVE_STANDARD_MODULES_RUNTIME.md;
F052/F178/F042/F135 und456/384/72 unverändert. Direkt weiter mit echtem FileDialog
für Standardmodule und zugehöriger Projekt-/Dialogzustandslebensdauer.


## Fortschreibung 2026-11-09: echte FileDialog-Picker aus Standardmodulen
Application.FileDialog(DialogType) verwendet den vorhandenen tatsächlichen Datei-/
Ordnerpicker (3/4), Office.FileDialog-Typ, Filter/Show/SelectedItems. Keine
Ersatzform, kein injiziertes Ergebnis, Auswahl öffnet/führt keine Datei aus.
Projektfenster hält den Pickerzustand; tatsächlicher Projektwechsel leert den
Cache. Alte Picker bleiben3420-geschützt, neuer Projektaufruf erzeugt funktionie-
renden neuen Picker. Application besitzt jetzt zusätzlich eine eigene Datei-/
Generationslease, unabhängig von einer optionalen Prüfung des Aufrufers.
App gebaut; bestehende Layout1/1 (11,31s) und abschließende GUI1/1 (8,15s) bestanden:
reale temporäre Datei/Ordner, Abbruch/geleerte Auswahl, Filter, typisierter
namentlicher VBA-Aufruf, Open1 verweigert/Null94 sowie tatsächlicher Projektwechsel
mit frisch erzeugtem Application-/Pickerobjekt bei ausdrücklich leerer Zusatz-
prüfung des Hosts. Finale App/GUI auf demselben Quellstand neu gebaut.
Frisches DEB native-module-file-dialog/package/linaccess-0.7.0-Linux.deb:
Payload/Rechte/Build-ID/bytegleicheSQLite und tatsächliches MDBTools1.0.1/
SQLite3.54.1-Backend (JOIN/NULL/Zähler) bestanden; keine MDB/VBA/Fixtures enthalten.
Keine neue CTest-ID/Gesamtsuite/Windows-Beobachtung. NATIVE_FILE_DIALOG.md;
F052/F042/F059/F135 und456/384/72 unverändert. Direkt weiter mit SysCmd am echten
Projekt-Statusbalken/GetObjectState und weiteren vorhandenen Application-Diensten.


## Fortschreibung 2026-11-09: wirkliche Application-Dienste/gemeinsames SysCmd
Standardmodule unterstützen Application.CurrentDb (vorhandener echter DAO-Adapter)
und Application-Domänenfunktionen mit Expr/Domain/Criteria über die bestehenden
Implementierungen. SysCmd1..5 verwendet realen Projektstatus/Meter,10 den echten
CurrentProject-Objektzustand; bare und namentliche Application-Aufrufe unterstützt.
Beide Hosts teilen Argumentprüfung/einmalige Skalargetter-Auswertung mit erneuter
Kontextprüfung. Native Statusoberfläche erkennt jetzt Hauptfenster selbst und
verwendet dessen echten Statusbalken; korrigiert auch Echo-Statustextposition.
App gebaut; bestehende GUI/Layout2/2 (19,72s) ohne Korrekturlauf bestanden: echte
VBA-DAO-Lese-/Domänenaufrufe, benannte Parameter, Status/Meter/Entfernen, tatsäch-
licher Tabellen-Ladezustand, einmaliger Standardwertgetter beider Hosts und Echo
am Projektbalken. Keine neue CTest-ID/Gesamtsuite/Windows-Beobachtung/Paketprüfung;
DEBf4e9ff8. NATIVE_STANDARD_MODULES_RUNTIME.md; F052/F053/F135 und456/384/72 bleiben
unverändert. Direkt weiter mit vollständiger tatsächlicher OpenForm-/OpenReport-
Anbindung (Where/DataMode/OpenArgs/WindowMode) über gemeinsame Dokumentöffnung.


## Fortschreibung 2026-11-09: gemeinsame tatsächliche Dokumentöffnung
Standardmodul-DoCmd.OpenForm hat sieben, OpenReport sechs Argumente und nutzt nun
dieselbe tatsächliche Öffnung wie der Formularhost: gespeicherte/importierte
Objekte, WhereCondition, frische Datenmodi, OpenArgs vor Load, native Fenster-
modi/acDialog-Warten, unabhängige Projektzuordnung und verzögerte Freigabe. Keine
Platzhalterform. Importquellen main-qualifiziert; explizite Where-Quellfehler
behalten Vorrang und Initialisierungsfehler jetzt Source-/Help-Metadaten.
Normalformular0/Berichtsvorschau2 bleiben unterstützt; andere Ansichten/benannte
Filter und unerlaubte Änderung bestehender Instanzen weiterhin3251. Native Makro-
aktionen haben bis zum nächsten Editorumfang weiterhin nur Name/Ansicht.
App gebaut; vorhandene GUI/Layout/Original-Kassa3/3 (137,86s;9,24/12,39/116,22s)
bestanden. Neuer Prüffall-Cast korrigiert, importiertes Testlayout mit tatsächlicher
OnLoad-Bindung ergänzt und alte Report-Window-Verweigerung durch echtes Aus-/
Einblenden ersetzt. Geprüft: Filter/Readonly/OpenArgs/Load, Wiederverwendung ohne
Mutation bei verweigertem Modus, echtes Add-Insert, modales Warten, versteckter
Bericht, Fehler7006 inkl.Source/Help und tatsächlicher importierter Handler.
Original-F129/3075 bleibt bekannte Grenze. Keine neue CTest-ID/Gesamtsuite/Windows-
Beobachtung/Paketprüfung; DEBf4e9ff8. NATIVE_STANDARD_MODULES_RUNTIME.md;
F052/F178/F042/F135 und456/384/72 unverändert. Direkt weiter: optionale Makro-
argumente einschließlich ausdrücklich leerem OpenArgs verlustfrei bearbeiten und
vollständige OpenForm-/OpenReport-Makroargumente ausführen.


## Fortschreibung 2026-11-09: verlustfreie optionale Makroargumente
Argumente… hat Übergeben-Schalter und bewahrt ausdrücklich leeren Text getrennt
vom Standardwert. Präsenz über Laden/Teilmakrowechsel/Aktionsreihenfolge/Speichern
und Zusatzargumente erhalten; Übersichtsänderungen inferieren neu, Pflichtwerte
bleiben erforderlich. Vollständige native OpenForm7-/OpenReport6-Makroargumente.
Optionales JSON-null bezeichnet einheitlich Missing, kein Datenbank-/VBA-Null;
Pflicht-null bleibt ungültig. Alte Formularbefehle nutzen ihre bewährten Defaults,
MessageBox verarbeitet ausgelassene Werte explizit. Kein erfundenes DoCmd-API.
Vorhandene GUI/Layout2/2 (20,59s) ohne Korrekturlauf bestanden: echte Editor-
Schalter/Move/Gruppe/Save/Reopen, leeres OpenArgs bleibt String und wird in beiden
Hosts tatsächlich von Null unterschieden; reale alte Befehle mit Default-Löchern
und MessageBox-Icon-Default. Anschließend nur irreführende optionale Übersichts-
spaltenbeschriftungen entfernt, keine Logikänderung. Frische App/DEB native-macro-
presence/package/linaccess-0.7.0-Linux.deb mit Payload/Rechten/Build-ID/bytegleicher
SQLite und echtem MDBTools1.0.1/SQLite3.54.1-Backendcheck (JOIN/NULL/Zähler) geprüft.
Uploads3/3 und Fixtures4/4 hashgleich. Keine neue CTest-ID/Gesamtsuite/Windows-
Beobachtung. Originaltest zuletzt f86d8eb:3/3 inkl.Original116,22s; F129/3075 offen.
NATIVE_MACROS_RUNTIME.md; F088/F052/F178/F135 und456/384/72 unverändert.
Direkt weiter mit echter SetValue-Makrozuweisung über vorhandene native Formular-
/Steuerelementobjekte und bestehende Ausdrucksauswertung; dokumentierte Grenzen.


## Fortschreibung 2026-11-09: tatsächliche SetValue-Makrozuweisung
SetValue(Item,Expression) greift auf echte native Controls/ausdrückliche Formular-/
Berichts-/Control-Eigenschaften und vorhandene Eval-/Setterwege zu, keine erzeugte
VBA-Zuweisung/keine Ersatzform/kein erfundenes DoCmd.SetValue. Bare Controls brauchen
echten Aufrufer/aktives Dokument, explizite Forms-Ziele funktionieren ohne Me.
Syntax/4096-Grenzen vorab, dokumentierte führende=-Doppelauswertung ausdrücklich
3251, Ziel-/Kontextprüfung vor/nach Auswertung. Vorhandene Readonly-/Puffer-/Save-
und Validierungsregeln gelten; Feld-ohne-Control/Funktionsziel-/Vollreport-/Ereignis-
parität bleibt offen. App gebaut; Layout1/1 (11,34s), abschließende GUI1/1 (8,49s)
bestanden. Gezielter Prüffall nach Einblenden wählt tatsächliche Form ausdrücklich;
vorher438 war mangels aktivem Dokument korrekt. Reale Zuweisung/Speicherung beider
Hosts, explizites verstecktes Ziel/Visible, Null, Fehler11/2448 ohne Wertänderung,
Schließen während RHS2450, ungültiges Rohmakro vor SQL-Präfix, DoCmd.SetValue438.
Keine neue CTest-ID/Gesamtsuite/Windows-Beobachtung/Paketprüfung; DEB22a9284.
NATIVE_MACROS_RUNTIME.md; F088/F052/F030/F135 und456/384/72 unverändert.
Direkt weiter mit tatsächlichem RepaintObject für geöffnete native Dokumente.


## Fortschreibung 2026-11-09: tatsächliches RepaintObject
RepaintObject(ObjectType,ObjectName) in Standardmodulen/Formularen und nativen
Makros zeichnet echte geöffnete Tabellen/Abfragen/Modul-/Makroentwürfe sowie
Formulare/Berichte neu. Form.Repaint verwendet vorhandene berechnete Controls,
Berichte rendern ihren vorhandenen Datensatzbestand erneut; keine Requery oder
implizite Speicherung. Projekt-/Generation-/Lebensdauer- und Kontextprüfungen
bleiben wirksam. Tabellen-Cache, laufende Zelleneingabe und Entwurfsänderungen
bleiben erhalten. GUI1/1 (9,43s), Layout1/1 (11,34s) bestanden; App gebaut.
Der letzte gezielte Korrekturlauf betraf ausschließlich den Prüffall: Speichern
per Schaltfläche schließt den Editor bestimmungsgemäß; Strg+S hält ihn offen.
GUI-Erweiterung in kleinere Quelldatei desselben bestehenden CTest-Ziels getrennt,
keine neue CTest-ID/Gesamtsuite/Windows-Beobachtung/Paketprüfung; DEB22a9284.
Vollständige Access-Neuzeichnungs-/Ereignis-/Berichtparität bleibt offen;
F088/F052/F050/F135 und456/384/72 unverändert.


## Fortschreibung 2026-11-09: SetValue-Ausdruck mit führendem Gleichheitszeichen
Expression mit führendem= wertet erst den Rest einmal aus und danach dessen
Skalarwert als Ausdruckstext einmal. Reale Referenzen/Funktionen beider Hosts
verwenden vorhandene Eval-/Setterwege; keine bloße Entfernung des=, keine
unbegrenzte Rekursion. Null94/Errornummer, Syntax-/4096-Textgrenzen, erneutes=3251
und Kontext-/Lebensdauerprüfungen vor Zuweisung. Dynamisches Item bleibt offen.
App gebaut, vorhandene GUI/Layout2/2 (20,35s) ohne Korrekturlauf bestanden:
Zahl, referenzierender Text, tatsächlicher öffentlicher Funktionsaufruf exakt
einmal je Host (SQL-Audit), sechs Fehlerfälle ohne Wert-/Dirty-Änderung und reale
Speicherung. Ungültiger Rohmakro-Präfixtest verwendet nun tatsächlich ungültige
Syntax=1+; vollständiger Vorabcheck vor SQL bleibt belegt. Keine neue CTest-ID/
Gesamtsuite/Windows-Beobachtung/Paketprüfung; bisheriges DEB22a9284.
NATIVE_MACROS_RUNTIME.md; F088/F052/F118/F164 und456/384/72 unverändert.
Direkt weiter mit echten Formularfeldern ohne gebundenes Steuerelement.


## Fortschreibung 2026-11-09: echte Formularfelder ohne Steuerelement
Native FormField-Referenzen lesen/ändern tatsächliche RecordSource-Spalten ohne
Ersatz-Controls; Name/Value/Parent und Lebensdauer geprüft, Controls haben Vorrang,
bereits gebundene Feldnamen verwenden tatsächliches Control. Eigener typisierter
Datensatzpuffer ist in Save/Undo/ControlUndo/Neuanlage/Navigation und berechnete
Controls integriert. Parameterisierte gemeinsame Speicherung mit anderen Controls,
CHECK-Fehler behalten Entwurf, AutoWert/Readonly/geschlossene Form verweigern.
VBA und SetValue beider Hosts verwenden denselben Weg; keine DAO-/Berichts-/
Join-/Vollereignisparität behauptet. FORM_RECORD_FIELDS_RUNTIME.md.
App gebaut; abschließende GUI1/1 (8,50s), Layout1/1 (11,32s), Original1/1 (124,97s)
bestanden. Korrekturlauf änderte ausschließlich den ControlUndo-Prüffall auf
echte Tastatureingabe statt programmatischer Zuweisung ohne Eingabepuffer.
Original-F129/3075 bleibt erwartete Grenze. Neues DEB native-form-record-fields/
package/linaccess-0.7.0-Linux.deb mit tatsächlichem MDBTools1.0.1/SQLite3.54.1,
JOIN/NULL/Zähler, Payload/Rechten/Build-ID/fehlenden Fixtures geprüft. Uploads3/3
und Fixtures4/4 hashgleich. Keine neue CTest-ID/Gesamtsuite/Windows-Beobachtung.
F056/F050/F088/F052 und456/384/72 unverändert.
Direkt weiter mit echten ApplyFilter-/ShowAllRecords-Dokumentfiltern.


## Fortschreibung 2026-11-09: tatsächliche Dokumentfilter
Gemeinsamer nativer Filterweg für echte Datenblatt-Dialoge und DoCmd/Makro
ApplyFilter/ShowAllRecords beider Hosts. Jet-Prädikat ohne zusätzliche Klauseln/
Subqueries; VBA32768/Makro256 UTF-16-Grenzen, Null94, benannte Filter3251. Tabellen
behalten Entwürfe (Filterwechsel2501), bearbeitbare Abfragen behalten Basis-WHERE.
Readonly-Ergebnisse verwenden tatsächliche Quelle/Parameter und werden erst nach
Erfolg ausgetauscht; Requery behält Filter. Formulare nutzen vorhandenen atomaren
Filter-/Validierungsweg; Berichte nur tatsächliches gebundenes Open-Ereignis,
Kontext nach Verschachtelung wiederhergestellt. DATASHEET_FILTER_RUNTIME.md.
App gebaut; abschließende GUI1/1 (8,62s), Layout1/1 (11,77s) bestanden. Prüffall
sortiert technische Tabellen ausdrücklich statt Qt-Standardreihenfolge zu raten.
Wiederherstellung berücksichtigt Qt: nach fehlgeschlagener Selektion ist Query
inaktiv und setFilter allein selektiert nicht erneut; explizite Auswahl ohne
doppelte erfolgreiche SQL-Ausführung. Reale UI/Cancel/SQL/Dirty/LIKE/Datum/
Parameter-/Abfrage-/Makro-/Form-/Report-Open-Prüfungen, kein neues CTest-Ziel.
Keine Gesamtsuite/Windows-Beobachtung/neue Paketprüfung; DEB073c20a, Original
zuletzt073c20a1/1 (124,97s), erwartetes F129/3075 unverändert offen.
F080/F056/F052/F088/F087 und456/384/72 unverändert.


## Fortschreibung 2026-11-09: echte Fensterbefehle und kürzere Änderungswege
Standardmodule/Makros unterstützen Maximize/Minimize/Restore ohne Ersatzform/Me.
Tatsächliche aktive Form/Report delegiert bestehende native Befehle, Tabellen/
Abfragen/Explorer adressieren echtes Projektfenster. Spezialisierter VBA-
Maximize-Statementpfad überschreibt tatsächliche Host-API; Minimize/Restore
laufen über DoCmd, dafür existieren keine virtuellen Hostmethoden. Modulfenster
3251, andere modale Entwürfe2046; keine implizite Speicherung/Entwurfsverwerfung.
Makroparser/Katalog/Loader einmalig als native_macro.cpp in bestehendem Daten-
Staticlib statt Headerdefinitionen; kleinere GUI-Erweiterungsdatei erhält festen
Einstiegspunkt. Keine neue DSO/GUI-Abhängigkeit des Datenlayers/CTest-ID.
App gebaut; abschließende GUI1/1 (8,69s), Layout1/1 (11,29s) bestanden. Gezielte
Prüffallkorrektur berücksichtigt Qt-Kombination minimiert+zuvor maximiert statt
fälschlich exklusive Statusflags zu verlangen. Reale Fensterzustandsereignisse,
Statement/Makros, unveränderte SQL-/Cache-/Formwerte, extra450, tatsächlicher
Moduleditor mit erhaltenem Entwurf. NATIVE_WINDOW_STATE_RUNTIME.md.
Volle Access-MDI-/Popup-/Geometrie-/Windows-Ereignisparität bleibt offen. Keine
Gesamtsuite/Windows-Beobachtung/neue Paketprüfung; DEB073c20a, Originalzuletzt
073c20a1/1 (124,97s), erwartetes F129/3075 offen. F052/F088/F135 und456/384/72
unverändert. Direkt weiter mit zusammenhängenden Makro-Bedingungsblöcken.


## Fortschreibung 2026-11-09: zusammenhängende Makro-Bedingungsblöcke
Condition=... übernimmt einmal ausgewertetes Ergebnis des unmittelbar vorigen
Bedingungsblocks. Leere Zeile beendet Block, neue Bedingung startet neuen;
Wiederholungen/Unteraufrufe haben eigenen Zustand. Verwaiste Fortsetzungen und
Gruppenanfänge werden im vollständigen Vorabcheck3251 vor SQL-Präfix abgelehnt.
Bestehende Aktions-/Tiefenbudgets zählen übersprungene Einträge weiter. Tatsächlicher
Makroeditor speichert Fortsetzungstext, Standard-/Formhost führen echte SQL-Audits
aus. Runner/Frame-/Stop-Helfer nach native_macro_runner.cpp im bestehenden
Daten-Staticlib ausgelagert, UI bleibt Callback, keine neue GUI-Abhängigkeit/DSO.
App gebaut; bestehende GUI/Layout2/2 (20,25s; 8,89/11,36s) ohne Korrekturlauf
bestanden: echter Editor, Funktion exakt einmal pro Block, True/False-Repeats,
Child isoliert Parent, defekte andere Gruppe/Präfix ohne Ausführung, Folgelauf
sauber. MACRO_CONDITION_BLOCKS_RUNTIME.md; keine neue CTest-ID/Windows-Referenz/
Gesamtsuite. Frisches DEB native-macro-continuation/package/linaccess-0.7.0-Linux.deb
mit Payload/Rechten/Build-ID/echtem MDBTools1.0.1 und SQLite3.54.1-JOIN/NULL/Zählern
geprüft; umfasst auch Filter/Fensterbefehle. Uploads3/3 und Fixtures4/4 hashgleich.
Originalzuletzt073c20a1/1 (124,97s), erwartetes F129/3075 bleibt offen.
Vollständige Access-Makro-/Ereignisparität bleibt offen. F088/F052/F135 und
456/384/72 unverändert. Direkt weiter mit SQL-Zugriff auf tatsächliche skalare
Formular-/Berichts-/Steuerelementeigenschaften ohne Methodenaufrufe.


## Fortschreibung 2026-11-09: echte skalare Objekteigenschaften in SQL
Jet-Tokenadapter unterstützt Forms!Dokument.Eigenschaft und Control.Eigenschaft,
auch Reports/lokalisierte Namen/Bracket-/Punktnamen. Bestehende Werte/.Value
bleiben auf alter3er-API, neue4er-Property-API ist driver-/verbindungsgebunden.
Explizite Getter-Allowlist verwendet tatsächliche Form-/Report-/Control-/
Datensatzfeldobjekte; Undo/Requery/Objektgetter/Methoden werden nicht ausgeführt
(438), tiefere/parametrisierte Pfade bleiben offen. Bisherige skalare SQL-
Konvertierung wiederverwendet, keine neue Typ-/Currency-/Decimal-Vollbehauptung.
Gemeinsamer Lookup prüft GUI-Thread, Connection/Pfad/Kind/echte Lebensdauer,
mehrdeutige Namen5 und nach Getter erneut tatsächliche Dokumentgültigkeit.
SQL_OBJECT_PROPERTIES_RUNTIME.md. App und Ziele gebaut; gebündelt4/4 (115,92s)
ohne Korrekturlauf bestanden: Layout11,68s, GUI9,01s, SQL-Objekte0,02s, Original
115,91s. Reale Caption/Tag/Visible/Dirty/OldValue/ListCount, native versteckte
Feldentwürfe, Live-View nach Änderung, fehlender Kontext, API-Reopen, echte
Verbindungsisolation/Mehrdeutigkeit, Getterverweigerung ohne Undo, Form-/
Reportfehler2450/2451. Original-F129/3075 bleibt erwartete Grenze.
Keine neue CTest-ID/Gesamtsuite/Windows-Beobachtung/Paketprüfung; DEB72ed466.
F067/F068/F118/F164/F050/F052 und456/384/72 unverändert.


## 2026-11-09: nativer VBA-Debugger, F060

Implementiert: tatsächliche gespeicherte Standardmodul-Ausführung mit F9/F5/F8,
StepInto/Over/Out, lebenden Aufrufrahmen/ByRef-Locals, paginierter Array- und
vollständiger Stringinspektion ohne Objektgetter, unbehandelten Fehlern auch
bei normalem F5, Resume-Next-Überspringen und explizitem Abbruch. Keine erneute
Ausführung bereits erfolgter SQL-Aktionen zum Fehlerhalt. Quelle/Projektwechsel
sind während des Halts geschützt; beendete Rahmen werden nicht weiter gelesen.
Details und Grenzen: docs/NATIVE_VBA_DEBUGGER_RUNTIME.md.

Nachweise in bestehenden Zielen: Kern1/1 (2,18s), Layout1/1 (11,80s), originales
Kassa-Ziel1/1 (118,52s), abschließender GUI-Ablauf1/1 (9,06s) sowie separat frisch
GUI-freier Kern1/1 (1,89s). Abschließendes DEB/Build-ID/MDBTools1.0.1/SQLite3.54.1
und vier Fixture-/drei Uploadhashes bestanden. Kein neuer Vollsuite-, Windows-,
Distributionsmatrix- oder GitHub-CI-Nachweis. Original-F129/3075 bleibt offen.
F060 bestätigt diesen beschriebenen nativen Teilumfang; vollständige IDE-,
Makrodebugger-, Sprach-, COM- und Ereignisparität bleibt in den übrigen offenen
Anforderungen. Fester Katalog456:385 bestätigte Teilumfänge/71 offene
Vollanforderungen, ausschließlich theoretisch und ungewichtet. Kein Gesamtabschluss.


## 2026-11-09: projektbezogene Einstufung, F090

Extras-Bericht mit deutschem Statusbaum und explizitem JSON-Export für den
wirklichen Import-/Objekt-/Tabellenkatalog. Getrennte statische Syntaxannahme,
Kontextgrenzen und tatsächlich beobachtete VBA-/Makro-/SELECT-Aufrufe samt
Fehlern und Abbruch. Verschachtelte Aufrufe, Quellrevisionen, nicht gestartete
Prozeduren, ungültige Definitionen und Rollback-Beobachtungen nachgewiesen.
Keine automatische Geschäftsprozedurausführung, keine VBA-/Argument-/
Datensatzwerte im Export. Projektwechsel leert die maximal1024 Sitzungsbelege;
gleichlautende Quellen belegen ausdrücklich keine konkrete Objektinstanz.
Andere SQL-/Tabellenmodellpfade und komplettes Dokumentrendering gelten als
nicht beobachtet, niemals als erfolgreich. Scope: docs/PROJECT_CAPABILITY_REPORT.md.

Gebündelter Lauf: Kern1/1 (2,04s), Layout1/1 (12,09s), Original1/1 (120,17s);
GUI-Testreihenfolge nach absichtlich defektem Modul korrigiert. Abschließende
gezielte GUI/Layout-Nachprüfung2/2 (13,93s; GUI13,93s/Layout12,54s), frisch
GUI-freier Kern1/1 (2,02s). Tatsächliche Makrorevision und gültiger Projektkontext
für Beobachtungen geschützt. Abschließendes DEB/Build-ID/MDBTools1.0.1/
SQLite3.54.1 und vier Fixture-/drei Uploadhashes bestanden. Keine neue CTest-ID,
keine aktuelle Vollsuite/Distributionsmatrix/GitHub-CI behaupten. Original-F129/
3075 bleibt offene Ausführungsgrenze. F090 bestätigt den beschriebenen
Klassifizierungsumfang, keine vollständige Ausführungs- oder Access-Parität.
Fester Katalog456:386 bestätigte Teilumfänge/70 offene Vollanforderungen,
ausschließlich theoretisch und ungewichtet. Kein Gesamtabschluss.


## 2026-11-09: Boolean-Bedingungen und If-Fehlerfortsetzung

Gemeinsame native Let-Boolean-Konvertierung für If/ElseIf, Do-/While-Köpfe und
Fußbedingungen, IIf sowie Makrobedingungen/Wiederholung: skalare Standardwerte
einmal, Null als False, True/False-Strings und case-sensitive #TRUE#/#FALSE#.
CBool(Null) bleibt94; Fehlerwerte/Arrays/Nothing werden nicht als True versteckt.
Belegte Tokenformen mit angrenzenden Klammern, EndIf/End If/End With mit normalen
Tokenabständen und leere einzeilige Then-/Else-Teile. EndIf vor Doppelpunkt ist
kein Label; leeres Case wird vor Effekten verweigert. Fehlgeschlagene If/ElseIf-
Köpfe behalten ausstehende Rahmen für Resume, Resume Next und Debugger-Skip;
keine irrtümlichen Then-/Else-Effekte und keine beschädigten äußeren If-Rahmen.
Scope/Sprachspezifikationsbeleg/Grenzen: docs/VBA_BOOLEAN_CONDITIONS.md.

Gezielte Nachweise nach vollständiger Implementierung/Korrektur der technischen
Prüfszenarien: Kern1/1 (2,22s), GUI1/1 (13,00s), abschließend aktuelles originales
Kassa-Ziel1/1 (113,95s), frisch GUI-freier Kern1/1 (2,00s). Tatsächliche Makro-
SQL-Effekte, Null-Standardwerte/Getterzahl, alle Bedingungspositionen, Resume und
Debugger-Skip geprüft. Finales DEB/Build-ID/MDBTools1.0.1/SQLite3.54.1 sowie
vier Fixture-/drei Uploadhashes bestanden. Kein neuer CTest, kein aktueller
Gesamtsuite-/Windows-/Distributionsmatrix-Nachweis. F129/3075 bleibt erwartet.
F045/F041/F051/F088 und456/386/70 bleiben offen/unverändert; kein Gesamtabschluss.


## 2026-11-09: fehlerhafte Schleifen-/With-Einstiege

Native Resume-Next-/Debugger-Skip-Fortsetzung überspringt fehlgeschlagene
While/Do-/For-/For-Each-Einstiege bis zur tatsächlichen Blockgrenze. Kein
Rumpflauf ohne gültigen Schleifenrahmen; äußere Schleifen/If bleiben erhalten.
With hält die äußere Aktivierung getrennt vom ausstehenden Objektkopf, damit
sein fehlgeschlagener Rumpf inaktiv bleibt und End With den äußeren Zustand
wiederherstellt. Wiederholendes Resume versucht den wirklichen Kopf erneut.
Scope und Grenzen in docs/VBA_BOOLEAN_CONDITIONS.md.

Ein gebündelter bestehender Kern-/GUI-Lauf2/2 (13,20s; Kern2,13s/GUI13,20s)
ohne Korrekturlauf bestanden. Aktuelle App und DEB/Build-ID/MDBTools1.0.1/
SQLite3.54.1 geprüft. Letzter Original- und GUI-freier Kernnachweis bleibt der
unveränderliche vorherige dae6681 (Original113,95s/Kern2,00s), kein neuer
Original-/Qt-frei-/Gesamtsuite-/Windows-/Distributionslauf behaupten.
Keine neue CTest-ID oder Katalog-ID; F045/F051 und456/386/70 bleiben offen/
unverändert. Kein Gesamtabschluss; weitere Arbeit läuft fort.


## 2026-11-09: Datensatzcursor und übergeordnete Requery-Berechnung

Schreibgeschützte Abfrageformulare wählen nach leer→gefüllt den ersten echten
Datensatz. Nur zulässige direkte Tabellenformulare behalten eine neue Leerzeile.
Geänderte tatsächliche Kind-Snapshots aktualisieren berechnete Elternwerte, ohne
Elterneingaben zu speichern; unveränderte Snapshots bleiben zwischengespeichert.
Details und Grenzen: docs/READONLY_FORM_REQUERY_RUNTIME.md.

Gebündelt Layout/GUI2/2 (14,90s), originales Kassa-Ziel1/1 (118,10s), finales
DEB/Build-ID/MDBTools1.0.1/SQLite3.54.1 und vier Fixture-/drei Uploadhashes
bestanden. Originale Keyboard-Grenze jetzt3146/1023; echter getaggter Klick
reicht bis2485 (fehlende exportierte Makroquelle). Beide MDB-Kataloge enthalten
vier Makros, bisheriges Brückenarchiv null. Kein neuer Gesamtsuite-/Windowslauf,
keine vollständige F129-Erfüllung. Katalog456/386/70 bleibt unverändert.


## 2026-11-09: natives Form.RecordsetClone

Eigene DAO-Snapshotcursor über tatsächliche Tabellen-/Abfrageformularergebnisse,
mit unabhängigem Suchen/Navigieren, Bookmark-Rücknavigation, leerem Ergebnis,
Requery-Familienwechsel und weiterlebenden gehaltenen Clones. Eigenschaft ohne
Argumente/schreibgeschützt; keine editierbaren Abfragen behauptet. Bestehende
Layout-/GUI-Prüfung gebündelt2/2 (12,40/14,29s) bestanden, kein neuer Testtarget
oder Windows-/Gesamtsuite-/Installerlauf. Details FORM_RECORDSET_CLONE_RUNTIME.md;
F054/F160 und Katalog456/386/70 bleiben offen/unverändert.


## 2026-11-09: tatsächliche Dateisammlungen und Unterordner

Native Scripting.Files/Folders unterstützen Item(Key), Standardzugriff und
indizierte Folder.Files/SubFolders-Eigenschaften über echte direkte Linux-Kinder.
SubFolders.Add(Name) erzeugt einen wirklichen Unterordner; Duplikate/Null/fehlende
Einträge/ungeeignete Pfade bleiben explizite Fehler. Native Namens-/Snapshotgrenzen
siehe NATIVE_FILE_SYSTEM_OBJECT.md; keine Windows-Gesamtparität behauptet.

Bestehender Kernlauf nach Implementierung gebündelt: Weiterleitung indizierter
Elterneigenschaften korrigiert, neue technische Unterordner vor dem vorhandenen
Symlink-Schutzfall entfernt; derselbe abschließende Kern1/1 (2,17s) bestanden.
App frisch gebaut. Kein neuer CTest, kein neuer vollständiger Desktop-/Windows-/
Qt-frei-/Installernachweis. F091 und Katalog456/386/70 bleiben offen/unverändert.


## 2026-11-09: direkter Formular-Recordset-Feldzugriff

Me.Recordset.Fields unterstützt tatsächliche Felder per Name/Ordinal, Count,
Item, Standard-/Bang-Zugriff und typisiertes For Each. Gehaltene Felder folgen
der aktuellen gespeicherten Zeile; Requery ersetzt ihre Identität (3420),
Formularschluss sperrt weitere Zugriffe (2450). Tatsächliche BLOB-Werte, FieldSize
und GetChunk; keine impliziten Speicherungen oder Feldschreibrechte. Nicht
beweisbare eingefrorene Metadaten bleiben Fehler; DataUpdatable=False beschreibt
den tatsächlich schreibgeschützten Adapter. Details FORM_RECORDSET_FIELDS_RUNTIME.md.

Gebündelt GUI1/1 (14,05s) bestanden. Layoutlauf fand fehlende DataUpdatable-
Fähigkeitsauskunft; korrigiert und nur dieses fehlgeschlagene Ziel erneut:
Layout1/1 (12,76s) bestanden. App gebaut, keine neue CTest-ID, kein neuer
Windows-/Gesamtsuite-/Qt-frei-/Installerlauf. F054/F160 und456/386/70 bleiben
offen/unverändert. Letzter Kern d2488e0:2,17s; letzte DEB-Prüfung9eae43f.


## 2026-11-09: echter Rückgängig-Verlauf im Formulardesigner

Natives Undo/Redo stellt Feldreihenfolge/-auswahl, Beschriftungen, Twip-Geometrie,
Datenquelle, VBA samt importierten Originalbytes, Form-/Feld-Ereignisse und Timer
im geöffneten Entwurf wieder her. Erfolgreiches Speichern setzt den tatsächlichen
Vergleichsstand; Rückkehr dorthin löscht den Änderungsstatus. Kein Undo-Speichern
von Datenbankobjekten; Verlauf bleibt begrenzt im Speicher, kein VBA-Historyexport.
Textfelder behalten ihr eigenes Text-Undo. Grenzen FORM_DESIGNER_HISTORY.md.

Bestehender GUI-Lauf1/1 zunächst13,67s; unverändert bestätigte Ereignisdialoge
als Nichtänderung bereinigt, abschließend dasselbe Ziel1/1 (13,52s) bestanden.
Neue aktuelle DEB-/Build-ID-/MDBTools1.0.1-/SQLite3.54.1-Prüfung bestanden, alle
vier Fixture- und drei Uploadhashes unverändert. Keine neue CTest-ID, kein neuer
Windows-/Gesamtsuite-/Qt-frei-Nachweis. Letzter Kern d2488e0:2,17s; letzter
Layout2a3a781:12,76s; letzter Original9eae43f:118,10s. F080/F148 und
Katalog456/386/70 bleiben offen/unverändert. Allgemeine Formularerstellung folgt.


## 2026-11-09: freie Formularerstellung und eigene Steuerelemente

Native leere/ungebundene Formulare und Formulare ohne ausgewählte gebundene Felder
sind speicherbar. Steuerelementdialog erzeugt bis128 tatsächliche TextBox/Label/
CommandButton/CheckBox/ComboBox/ListBox mit geprüften Namen, Twip-Geometrie,
Bindungen, Listenherkunft und Klickereignis. Vorhandene VBA-Ereignisse binden
automatisch; eigener Entwurf einschließlich Controls nimmt am Undo/Redo teil.
Ungebundene Ansicht zeigt keinen erfundenen Datensatz oder Quellenfehler an.
Umfang/Grenzen FREE_FORM_CONTROLS.md, bestehende F080/F081/F083/F148/F155.

Gebündelter bestehender Lauf: Layout1/1 (12,02s) bestanden. GUI fand einen
Prüfklick außerhalb der tatsächlichen Checkbox-Markierung; auf echten Qt-
Indikator korrigiert und nur dieses Ziel erneut: GUI1/1 (13,08s) bestanden.
Tatsächlicher Designer, Duplikatfehler/Korrektur, Save/Reload, Eingabe/Berechnung,
Label/Checkbox/Combo/List und unveränderte technische VBA mit unabhängig perSQL
bestätigtem Klick-Effekt geprüft. App gebaut; kein neuer CTest oder Katalogpunkt,
kein neuer Original-/Qt-frei-/Windows-/Gesamtsuite-/DEB-Lauf. Letzte geprüfte
DEB a837563; Katalog456/386/70 bleibt unverändert, keine Gesamtparitätsbehauptung.


## 2026-11-09: literale Formular-Feldbindungen

ControlSource dekodiert ]] in geklammerten Feldnamen konsistent für Laden,
Speichern, Feldreferenz auf vorhandene Alias-Controls und OldValue. Native
Generatoren klammern/escapen Feldnamen mit führendem =, äußeren Klammern oder
äußeren Leerzeichen. Keine Umbenennung von Spalten oder Originalquellen.
Native tatsächliche Lese-/Schreib-/Alias-/OldValue-/SQL-Nachweise im bestehenden
Layout-/GUI-Ziel abschließend2/2 (12,39/14,16s), aktuelle DEB-/Build-ID-/
MDBTools1.0.1-/SQLite3.54.1-Prüfung bestanden. Details
LITERAL_CONTROL_SOURCE_BINDINGS.md. Kein neuer CTest oder Katalogpunkt; keine
aktuelle Gesamtsuite/Windows/Original-Workflow-Ausführung behauptet. F155/F056/
F160 und456/386/70 bleiben offen/unverändert. Mehrspaltige native Listen folgen.


## 2026-11-09: mehrspaltige eigene Kombinations-/Listenfelder

Steuerelementdialog speichert/geprüft ColumnCount, BoundColumn und ColumnWidths,
inklusive getrennter Anzeige-/Wertespalte, gebundenem Zeilenindex und versteckten
Spalten. Gemeinsamer tatsächlicher Twip/cm/mm/in/pt-Parser für Designer und
Runtime; keine abweichend erfundene Breiteninterpretation. Bestehende Definitionen
behalten ihre Standardwerte. Native Save/Reload-/ValueList-/SQL-Modelle und echte
Mausauswahl im bestehenden GUI/Layoutlauf2/2 (13,83/12,51s) bestanden.
Details FREE_FORM_CONTROLS.md. Kein neuer CTest/Windows-/Original-/Gesamtsuite-/
DEB-Lauf; letzte DEB15131f0. F080/F083/F162 und456/386/70 bleiben offen/unverändert.


## 2026-11-09: eigene Standardwerte, TextBox-Format und Passwortmaske

Steuerelementdialog speichert DefaultValue/Format/InputMask mit bestehenden
Typ-/Längen-/Kontextgrenzen. Designer und tatsächliche Initialisierung verwenden
denselben nativen Literalparser; keine berechneten Standardwerte im Designer
ausführen. Format bleibt Darstellung bei unverändert typisiertem Value, Password
ist native Echo-Policy. Echte Create/Save/Reload-/Default-/Wert-vs-Anzeige-/SQL-
Nachweise gebündelt im bestehenden GUI/Layout2/2 (14,98/12,78s) bestanden.
Details FREE_FORM_CONTROLS.md. Kein neuer CTest, kein Kern-/Windows-/Original-/
Gesamtsuite-/DEB-Lauf; letzte DEB15131f0. F015/F071/F080/F081/F115 und456/386/70
bleiben offen/unverändert. Allgemeine Steuerelementeigenschaften folgen.


## 2026-11-09: eigene Zustände und Tabreihenfolge

Auswahlpanel für Visible/Enabled/Locked/TabStop/TabIndex im echten Controls-
Designer, mit Typ-/Bereichsprüfung und bestehenden nativen Laufzeitregeln.
Labels ohne Tab-Eigenschaften, Locked für TextBox/ListBox. Tatsächlich
speichern/neuladen, ausgeblendetes deaktiviertes Label, gesperrtes maskiertes
Textfeld, Tab-Ausschluss und echte Tastatur-Zielwahl im bestehenden Layout/GUI
gebündelt2/2 (12,28/14,29s) bestanden. Details FREE_FORM_CONTROLS.md.
Keine neue CTest-ID oder aktuelle Gesamtsuite/Windows/Original/DEB-Prüfung;
letzte DEB15131f0. F080/F081/F104/F148 und456/386/70 bleiben offen/unverändert.


## 2026-11-09: echte Schrift-/Farbauswahl eigener Controls

Qt-Auswahldialoge speichern geprüfte FontName/Size/Weight/Italic/Underline und
RGB-ForeColor; TextBox/Label zusätzlich BackColor/BackStyle transparent/opaque.
Standarddarstellung löscht Overrides. Persistenter Zieleintrag plus Auswahl-/
Fokuswiederherstellung verhindert nach Modalrückkehr die Änderung eines anderen
Controls. Keine Windows-Font-/Pixelparität behaupten. FREE_FORM_CONTROLS.md.

Tatsächliche Picker/Mausklicks, Save/Undo/Redo/Reload, native Paletten/Transparenz
und aufgelöster QFontInfo-Schnitt geprüft. GUI-Testsetup auf wirklich installierten
Fett/Kursiv-Schnitt korrigiert; Modal-Auswahlwechsel behoben. Abschließend nur
fehlgeschlagenes GUI-Ziel1/1 (14,70s) wiederholt/bestanden; Layout1/1 (13,09s)
bereits bestanden. Kein neuer CTest, keine Gesamtsuite/Original/Windows/Kern-
Ausführung; aktuelle DEB-Prüfung folgt.456/386/70 unverändert; F080/F081/F148 offen.


## 2026-11-09: Ausrichtung, Tag und lange Eigenschaftstexte

Native Controls speichern 2048-Zeichen-Tag sowie TextBox-/Label-Ausrichtung
Allgemein/Links/Zentriert/Rechts. Quellen-/Typ-/Bereichsgrenzen bleiben explizit.
Echter Tabellen-Delegate übernimmt den bestehenden65536-Eigenschaftsrahmen statt
32.767 stiller Eingabekürzung; kurze Teilpropertygrenzen gelten weiter.
GUI-Editor bearbeitet33.000 Zeichen+Anhang, Save/Reload prüft33.001, native
Ausrichtung/Tag ebenfalls. Offscreen-Dialog für echte Eingabe explizit aktiviert;
abschließend nur GUI1/1 (14,11s) wiederholt/bestanden; Layout1/1 (12,81s) bereits
bestanden. FREE_FORM_CONTROLS.md samt tatsächlicher Microsoft-Tag-Dokumentation.
Kein neuer CTest/Windows/Original/Kern/Gesamtsuite/DEB-Nachweis; letzteDEB d8ecf41.
F050/F080/F081/F148/F155 und456/386/70 unverändert. SQL-Kompatibilität folgt.


## 2026-11-09: Typisierte Kreuztabellen-Minima und -Maxima

Feste TRANSFORM-Min/Max-Überschriften verwenden echte physische Tabellen und
vorhandene native Currency-/Single-/Ganzzahl-/Boolean-/Datumsprüfungen. Pivottypen
bleiben erhalten, auch bei leeren Ergebnissen; fehlende Kategorien liefern Null.
Parameter, Aliasse, gespeicherte QueryDefs und dbQCrosstab16 sind eingebunden.
Text/Locale, unbestimmte Integerbreiten, Decimal, gespeicherte Quellansichten und
temporäre Überschattung werden nicht als unterstützt ausgegeben. Schemaänderungen
invalidieren typisierte Abfrageansichten. Details: COUNT_CROSSTAB_RUNTIME.md.
Abschließend bestehendes SQL-Bündel3/3 bestanden (0,12s); Anwendung neu gebaut.
Kein neuer CTest/Windows-/Original-/Gesamtsuiten-/Paketnachweis. F067/F068 bleiben
offen; Katalog456/386 bestätigte Teilumfänge/70 offene Vollanforderungen unverändert.


## 2026-11-09: Statistische SQL-Aggregate

Var und StDev sind in echte SQL-Ausführung eingebunden, mit bestehendem nativen
Berichts-/Domänen-Akkumulator, Double-Metadaten, n-1, Null-Auslassung und Null bei
weniger als zwei Werten. Gruppierung/DISTINCT/Parameter und große gemeinsame
Offsets sind geprüft; Text/BLOB13 und Überlauf6 bleiben explizit.
SQL_STATISTICAL_AGGREGATES.md beschreibt Umfang und Microsoft-Dokumentationsbasis.
Abschließend bestehendes SQL-Bündel4/4 bestanden (0,13s), Anwendung neu gebaut;
keine neuen CTests oder Windows-/Original-/Gesamtsuiten-Nachweise. Noch kein neues
Paket; folgt aus dem aktuellen Stand. F067/F068/F071 und456/386/70 unverändert.


## 2026-11-09: Berechnete Domänenausdrücke

DLookup/DFirst/DLast/DMin/DMax/DAvg/DVar/DStDev unterstützen begrenzte skalare
SQL-Ausdrücke, statt ausschließlich Feldnamen. Echte Projektion/Typen/Null, Currency-
und Datumsidentität, Zahlen-/Textausdrücke und statistische Ergebnisse sind im
bestehenden nativen VBA-Host geprüft. Tokenprüfung sperrt Zusatzklauseln, Subqueries,
verschachtelte Aggregate/Window-Aufrufe, mehrere Projektionen und fremde Funktionen.
NATIVE_DLOOKUP.md hält Grenzen fest; keine Text-/Locale-Extrema oder beliebige
Projektfunktionsbindung. Abfrageansichten invalidieren nun auch bei temporären
Tabellen, einschließlich tatsächlichem Entfernen/Wiederherstellen der zuvor
validierten Min/Max-Ansicht bei Überschattung. Bestehendes Bündel5/5 (12,21s)
bestanden, Anwendung gebaut; kein neuer CTest/Windows-/Original-/Gesamtsuitenlauf.
Letzte geprüfte DEB ist eaaeb73, inklusive Build-ID/MDBTools1.0.1/SQLite3.54.1;
alle vier Fixture-/drei Uploadhashes unverändert. Diese neue Domänenetappe ist
noch nicht paketiert. F050/F067/F068/F118/F164 und456/386/70 unverändert.


## 2026-11-09: Native bedingte VBA-Kompilierung

Qt-freier nativer Preprocessor für #Const/#If/#ElseIf/#Else/#EndIf/#End If samt
begrenzter Fortsetzungen. Auch ausgeschlossene Konstanten/Bedingungen werden
entsprechend MS-VBAL3.4/5.6.16.2 ausgewertet; aktive Modulfelder/Prozeduren,
OptionBase/Compare, lokale/öffentliche/SQL-Aufrufe und native Ereignisbindungen
verwenden eingeschlossenen Code. Quellenbytes/Hashes und Debuggeridentitäten
bleiben ursprünglich; echte Haltepunkte prüfen Originalzeilen. Linux-Hostflags
und genaue Grenzen in VBA_CONDITIONAL_COMPILATION.md; keine Windows-API-Fiktion.
Abhängigkeits-/Nesting-/Textbudgets sperren Zyklen und exponentielle Expansion.
Ein fehlender Array-Header/eine Syntaxklammer wurden beim Build korrigiert. Im
gebündelten Prüflauf bestanden7/8; Kernelprüfung fand falsches Zahlenfeld beim
Boolean-Lesen. Exakt auf .boolean korrigiert; danach gezielt Kernel/DAO2/2
(2,67s) bestanden, echter SQL-UDF-True-Zweig/Typ ebenfalls. GUI12,88s und15,34s
waren bereits bestanden und wurden für diesen Einzelfix nicht wiederholt.
Separater Qt-freier Build/Kernlauf ebenfalls bestanden; ldd ohne Qt. Keine neuen
CTest-IDs/Original-/Windows-/frische Gesamtsuitenbehauptung. Anwendung neu gebaut;
letzte geprüfte DEB weiterhin eaaeb73. F040/F042/F043/F045/F051/F068 bleiben offen;
Katalog456/386 bestätigte Teilumfänge/70 offene Vollanforderungen unverändert.


## 2026-11-09: Stop und Debug.Assert im echten Debugger

Native Interpreter-Anweisungen halten mit eigenen Stop-/Assertion-Gründen an
Originalzeilen und lebenden Variablen. Fortsetzen/Schritt/Beenden, eager vollständige
Boolean-Auswertung, einmaliger Standardgetter und Err-Erhalt unter ResumeNext
sind im bestehenden Umfang eingebunden. Ohne Debuggersitzung wird benötigter Halt
3251 gemeldet; kein stiller Erfolg. Echter Modul-Editor/Debugger prüft Status,
Variable17 und tatsächliche SQL-Fortsetzung18 mit anschließender Test-Rücknahme.
Bestehendes Kernel-/GUI-Bündel2/2 (14,49s) bestanden; Anwendung gebaut. Keine neuen
CTest-IDs oder weiteren Original-/Windows-/Gesamtsuiten-/Qt-frei-Nachweise für diese
Stop-Erweiterung; Details in NATIVE_VBA_DEBUGGER_RUNTIME.md.

GitHub-CLI-Token wurde nach931ea64 ungültig. Vorhandener GitHub-Connector besitzt
weiterhin bestätigten Schreibzugriff auf das private Repository. Die vorige
bedingte-Kompilierung-Etappe wurde darüber als1a599f6 veröffentlicht und per Ref
verifiziert: gleicher vollständiger Git-Dateibaum37583763b86188457803ca22944c4911e2c9c9af
wie lokaler6f95ca0, lediglich Veröffentlichungsmetadaten/neue Commit-ID. Ungepushtes
Originalobjekt bleibt als lokaler Sicherungsref erhalten; Arbeitsdateien unverändert.
Kein Token wurde gelesen/ersetzt/gespeichert. Wiederverwendbarer kontrollierter
Veröffentlichungshelfer außerhalb Git: /workspace/work/github-connector/publish.mjs
(über functions.exec/tools ausführen; Blob-/Tree-SHA, private Sichtbarkeit und
Fast-forward-Lease werden geprüft, lokale identische Remote-Commitobjekte übernommen).
F040/F045/F051/F060-Status und Katalog456/386/70 unverändert. Letzte geprüfteDEB eaaeb73;
frisches Paket folgt aus dem aktuellen Quellstand.


## 2026-11-09: Feste skalare VBA-Strings

Lokale/statische und modulweite String*n-Variablen mit1..65526 UTF-16-Einheiten,
bekannten Const-Längen und späteren Modulkonstanten sind implementiert. Initial
NUL-Zeichen, bei Zuweisung Abschneiden/Leerzeichenauffüllung; echte String8-Werte,
Metadatenbreite und statischer Zustand bleiben erhalten. LSet/RSet und Mid sind
eingebunden. Die Kernprüfung fand fehlende Mid$-Zuweisungssyntax; dieser konkrete
Parserpfad wurde ergänzt, inklusive unverwechselbarem intrinsischem Mid$-Modus.
Null94 verändert den vorigen Wert nicht; isolierte Surrogatabschneidung bleibt3251.
Details/Grenzen einschließlich noch fehlender Arrays/UDTs/kompletter ByRef-Parität:
VBA_FIXED_STRINGS.md. GUI14,77s bereits bestanden; nach Mid$-Fix ausschließlich
Kern1/1 (2,31s) wiederholt/bestanden, Anwendung neu gebaut. Keine neuen CTests oder
neuen Windows-/Original-/Gesamtsuiten-/Qt-frei-Belege für diese String-Erweiterung.
Vorige Stop-Etappe ist als72b936d über Connector veröffentlicht, gleicher geprüfter
Baumccd6e1f54fcbed0d1b72a5d4ff4cb6416d07c2be wie lokaler633fb98. Frische72b936d-DEB
unter /workspace/work/native-debug-halts/package/linaccess-0.7.0-Linux.deb besteht
Build-ID/Inhalt/Berechtigungen/MDBTools1.0.1/SQLite3.54.1-Prüfung. Diese neuen
String-Quellen sind noch nicht paketiert. F040/F041/F042/F043/F122 und456/386/70
unverändert. Keine Gesamtparität/Abschlussbehauptung.


## 2026-11-09: Feste Stringelemente in VBA-Arrays

Feste/dynamische/modulweite/statische String*n-Arrays unterstützen Initial-NUL,
Element-Konvertierung, Trunkierung/Leerzeichen, ReDim(Preserve), Erase und begrenzte
Variant-Kopien. Canonical-Elementmetadaten behalten die Breite unabhängig von
öffentlichem String()/8200. Fehlende Null-Zuweisung und Breitenänderung lassen
vorige Werte/Dimensionen unverändert. Neue Elemente/Erase-Festarrays werden NUL,
dynamisches Erase behält den deklarierten Typ.16MiB-Payloadgrenze prüft vor
Reservierung count*n*3 (UTF-8-Obergrenze pro UTF-16-Einheit); bestehende Bounds/
Iterator-/ByRef- und Millionen-Element-Grenzen bleiben aktiv. VBA_FIXED_STRINGS.md.
Skalare String-Prüfung aus großer core_tests.cpp in kleine vba_fixed_string_tests.cpp
innerhalb desselben CTest-Ziels verschoben; keine neuen Testsuite-IDs.
Abschließend Kernel/GUI2/2 (15,24s), Qt-freier Kern1/1 (2,58s) bestanden,
Anwendung gebaut. Kein neuer Original-/Windows-/Gesamtsuiten-/Paketnachweis.
Vorige skalare Etappe als e9ce132 veröffentlicht/verifiziert (gleicher getesteter
Baumce34e6522e2363f234c2ca20c23cca9da0f57395 wie lokaler d053a80). Letzte geprüfte
DEB weiterhin72b936d, frisches Paket folgt. F040/F041/F042/F043/F122 bleiben offen;
456/386 bestätigte Teilumfänge/70 offene Vollanforderungen unverändert.


## 2026-11-09: Visueller nativer Abfrageentwurf

SQL-Editor hat einen echten Qt-Entwurf für reale Tabellen/Felder: Ausgabeauswahl,
Alias/Feldreihenfolge, Sortierung, Kriterien und DISTINCT erzeugen tatsächliches
SELECT. Vorschau bereitet Metadaten ohne Daten-/VBA-Auswertung vor; erst Ausführen
liest die echten Zeilen. Abbruch ändert SQL nicht. Gespeicherter Entwurf ist an
exakten SQL-SHA gebunden; manuelle SQL-Änderung entfernt alte visuelle Metadaten.
Projekt-/Generations-Eingangsprüfungen und Ressourcen-/Namen-/Klauselgrenzen;
fortgeschrittenes SQL bleibt im SQL-Editor. Kein Kassa-spezifischer Designer.
Scopegrenzen in VISUAL_QUERY_DESIGNER.md, F067/F068/F080/F148 bleiben offen.

Im bestehenden SQL-/GUI-Bündel bestand SQL; GUI fand einen realen QJsonValueRef-
Kopierfehler beim Verschieben. Wertkopie korrigiert, danach nur betroffene GUI1/1
(14,73s) wiederholt/bestanden. Echte verschachtelte Dialoge, unabhängiges SQL,
Speichern/Wiederherstellen, Abbruch und Callback-Zähler beweisen tatsächlichen
Pfad ohne Metadaten-Seiteneffekte. Anwendung gebaut; keine neue CTest-ID/Windows-/
Original-/Gesamtsuitenbehauptung. Fester Katalog456/386/70 unverändert.
Vorige Stringarray-Etappe als1dafedd veröffentlicht/verifiziert (gleich geprüfter
Baum023f57fab171cac5324e8853bb2bb7c026c38c01 wie lokaler145786b). Neue1dafedd-DEB
unter /workspace/work/fixed-vba-string-arrays/package/linaccess-0.7.0-Linux.deb besteht
Build-ID/Inhalt/Berechtigungen/MDBTools1.0.1/SQLite3.54.1. Aktueller visueller Entwurf
noch nicht paketiert. Weiterentwicklung läuft; kein Gesamtabschluss.


## 2026-11-09: ODER-Kriterien und typisierte Entwurfsparameter

Visueller Abfrageentwurf unterstützt bis8 UND/ODER-Gruppen und64 deklarierte
Parameter (Text/Long/Integer/Byte/Single/Double/Currency/DateTime/YesNo). Echte
PARAMETERS-Deklaration, bestehende Eingabedialoge und SQL-/DAO-Bindung; Antworten
werden nicht gespeichert. Metadatenvorschau bleibt ohne Zeilenauswertung. Gruppen/
Parameter öffnen wieder korrekt, Feldverschiebung behält alle Kriterien. Alte
Ein-Gruppen-Entwürfe bleiben lesbar; genaue Grenzen in VISUAL_QUERY_DESIGNER.md.
Bestehendes GUI-/SQL-Bündel2/2 (14,51s) bestanden, inklusive tatsächlicher zweier
Parametereingaben, unabhängigem SQL und späterer manueller Metadateninvalidierung.
Keine neuen CTests oder Kern-/Original-/Windows-/Gesamtsuiten-/Paketnachweise.
Voriger visueller Entwurf als33cda14 veröffentlicht/verifiziert; gleicher geprüfter
Baum6e85a5c233a9ab51f14542cb36e703905789ad12 wie lokaler8e5458f. Letzte geprüfteDEB
weiterhin1dafedd. F067/F068/F080/F148 und456/386/70 unverändert; Entwicklung läuft.


Berechnete Felder im visuellen Abfrageentwurf: echte skalare Ausdrücke mit Alias,
Ausgabe, Sortierung und UND/ODER-Kriterien; Hinzufügen/Entfernen betrifft nur den
Entwurf. Gespeicherte Ausdrücke öffnen wieder korrekt. Begrenzte native skalare
Funktionen; Aggregate, Unterabfragen, unbekannte Callbacks und externe/lebende
Werte werden in diesem Entwurf ausdrücklich begrenzt. Keine Zeilenauswertung in
Vorschau/Übernahme. Details: docs/VISUAL_QUERY_DESIGNER.md. Anwendung gebaut;
bestehendes GUI-/SQL-Bündel2/2 bestanden (15,09s), inklusive wirklicher berechneter
Ergebnisse gegen unabhängiges SQL, Sortierung und Wiederöffnung. Keine neuen
CTest-IDs oder Windows-/Original-/Gesamtsuiten-/Paketnachweise. Parameter/Oder-
Etappe als b1b5858 veröffentlicht; letzte geprüfteDEB weiterhin1dafedd.
Katalog456/386/70 unverändert; keine vollständige Abfrageentwurfsparität behauptet.


Gruppierung im visuellen Abfrageentwurf: pro Feld Group By/Where/Count/Sum/Avg/
Min/Max/Var/StDev, echte GROUP BY-/WHERE-/HAVING-Klauseln mit getrennten ODER-
Gruppen; Where-Felder nur ausgeblendet/unsortiert. Entwurfsmetadaten behalten
Funktionen und Modus. Nachträglich hinzugekommene Quellspalten verändern keine
bestehende Gruppierung. Details: docs/VISUAL_QUERY_DESIGNER.md. Bestehendes GUI-/
SQL-Bündel2/2 (14,62s) bestanden; nach Korrektur der Schemaerweiterung nur betroffene
GUI-Prüfung1/1 (14,36s) wiederholt. Anwendung gebaut; keine neuen CTest-IDs oder
Windows-/Original-/Gesamtsuiten-/Paketnachweise. Berechnete-Felder-Etappe als
cee05d8 veröffentlicht/verifiziert; letzte geprüfteDEB weiterhin1dafedd.
F067/F068/F080/F148 sowie fester Katalog456/386/70 bleiben unverändert offen.


Visueller Datenquellenentwurf im nativen Berichtseditor: derselbe tatsächliche
SQL-Entwurf einschließlich berechneter Felder/Gruppierung; Übernahme als ein
Text-Undo, Abbrechen unverändert, Metadaten nur bei passendem SQL-Hash. Speichern/
Wiederöffnen und Invalidierung bei manueller SQL-Änderung nachgewiesen. Keine
Parameterdeklarationen in diesem Berichtsweg, da Berichtsparameterbindung noch
fehlt. Projektidentität/Generation wird bei Entwurf/Definition/Speichern/Ausgabe
geprüft. Bestehendes GUI-/SQL-Bündel2/2 (15,31s), Anwendung gebaut. Tatsächlicher
Callbackzähler: Entwurf und Speichern ohne Auswertung, explizite Ausführung liest
Zeilen. Keine neuen CTests oder Windows-/Original-/Gesamtsuiten-/Paketnachweise.
Details: docs/VISUAL_QUERY_DESIGNER.md. Gruppierungs-Etappe als31c50d2 veröffentlicht;
letzte geprüfteDEB weiterhin1dafedd. F080/F085/F087/F148 und456/386/70 unverändert.


Gespeicherte Auswahlabfragen als visuelle Datenquelle: tatsächliche native Query-
Views werden vor Entwurfsöffnung installiert und zusammen mit Tabellen angeboten;
Metadatenvorschau bleibt ohne Zeilenauswertung. Tatsächliche verschachtelte Abfrage
auf gespeichertem Entwurf inklusive berechnetem Feld/Kriterium/Sortierung und
Wiederöffnung geprüft; Callbackzähler trennt Vorschau von Ausführung. Native
Abhängigkeits-/Typ-/Parametergrenzen bleiben ausdrücklich erhalten. Anwendung
gebaut, bestehendes GUI-/SQL-Bündel2/2 (15,84s) bestanden. Ein Qt-Enum-Baufehler
wurde vor dem Prüflauf korrigiert. Keine neuen CTests oder Windows-/Original-/
Gesamtsuiten-/Paketnachweise. Berichtsentwurf als4e20658 veröffentlicht; letzte
geprüfteDEB weiterhin1dafedd. Details VISUAL_QUERY_DESIGNER.md;456/386/70 unverändert.


Visuelle Verknüpfungen: bis8 geordnete INNER-/LEFT-Equijoins auf tatsächlichen
Tabellen/gespeicherten Auswahlabfragen, pro neuer Quelle vorhandenes Feld = neues
Feld. Vollqualifizierte Feldidentität und persistierte Quellzuordnung; echte
Nullwerte bei fehlendem LEFT-Treffer. Entfernen betrifft nur Entwurfsfelder;
Selbstverknüpfungen/Aliase/zusammengesetzte Vergleichsbedingungen bleiben im
SQL-Editor mit bestehenden Grenzen. Strikte Metadatenprüfung; fehlgeschlagenes
Einlesen hinterlässt keine halbfertigen Feldzeilen. Bestehendes GUI-/SQL-Bündel2/2
(14,92s), nach Metadatenguards nur GUI1/1 (14,90s) wiederholt; Anwendung gebaut.
Tatsächliche zweistufige LEFT-Verknüpfung gegen unabhängiges SQL verglichen,
Wiederöffnung/Entfernen/INNER-Wechsel und kaputte Metadaten geprüft. Keine neuen
CTests oder Windows-/Original-/Gesamtsuiten-/Paketnachweise. Details in
VISUAL_QUERY_DESIGNER.md. Gespeicherte-Abfragequellen-Etappe28e3302 veröffentlicht;
letzte geprüfteDEB weiterhin1dafedd. F067/F068/F080/F148 und456/386/70 unverändert.


Visueller Formulardatenquellenentwurf: derselbe echte SQL-Entwurf erzeugt SELECT-
Quellen mit vorbereiteten Feldmetadaten; keine Parameterbindung und weiterhin
schreibgeschützte Abfrageformulare. Passende SQL-Hash-/Entwurfsmetadaten werden
mitgespeichert und im bestehenden Formular-Undo/Redo erfasst/wiederhergestellt.
Manuelle Quelländerung entfernt veraltete Metadaten. Projektidentitäts-/Generations-
prüfung bei Entwurf/Speichern/Verlauf. Bestehendes GUI-/SQL-Bündel2/2 bestanden
(15,47s), Anwendung gebaut; tatsächliches gespeichertes Formular mit echten
Ergebniswerten/schreibgeschütztem Zustand geöffnet. Callback zählt keine Auswertung
bei Entwurf/Speichern/Wiederöffnung, erst tatsächliche Formularausführung liest.
Keine neuen CTests oder Windows-/Original-/Gesamtsuitennachweise.

Verknüpfungs-Etappe alsfe01e1d veröffentlicht/verifiziert. Geprüftes DEB dieses
Standes unter /workspace/work/visual-query-joins/package/linaccess-0.7.0-Linux.deb:
Build-ID passt, MDBTools1.0.1/SQLite3.54.1, Payload/Berechtigungen/keine Fixtures
bestanden. Paket enthält noch nicht die anschließende Formularintegration. Vier
Fixture- und drei Upload-Originalhashes erneut unverändert geprüft. Details in
VISUAL_QUERY_DESIGNER.md/NATIVE_QUERY_FORM_DESIGN.md. F050/F080/F148 und fester
Katalog456/386/70 unverändert. Nächste unabhängige Untersuchung: ursprüngliche
3146/1023-Grenze beim Öffnen des Originalformulars, keine erneute Makro-Rückfrage.


Korrektur der bisherigen Original-3146/1023-Deutung: private Diagnose zeigte
Document initialization failed: Native macro catalog unavailable bei source-error0.
Die Import-Testkopie besitzt keinen optionalen main._la_objects-Katalog. Native
Makrosuche behandelt dessen Abwesenheit jetzt wie einen leeren Katalog (2485),
ohne Tabellen anzulegen oder temporäre Schatten zu lesen; vorhandene defekte
Kataloge bleiben3146. Beide tatsächlichen Tastaturhandler erreichen nun den
fehlenden Makroquelltext2485 in Zeile1023. Das ist keine offene Datensatzquellen-
SQL-Operation. Getaggter Click erreicht separat ebenfalls fehlenden Makroquelltext.
Bestehendes GUI-/Original-Bündel2/2 bestanden (125,46s; GUI15,85s); Anwendung gebaut.
Isolierte Prüfungen für fehlenden/temporären/defekten/leeren Hauptkatalog sind in
bestehendem Ziel enthalten. Keine neuen CTests, Windows- oder positive Original-
Verkaufs-/Makrolaufnachweise. Die diagnostischen Instrumentierungen/privaten Details
liegen ausschließlich außerhalb Git; sie sind keine zusätzlichen erfolgreichen
Prüfziele. Native Makroquelle fehlt weiterhin im Bridge-Export; vorhandene
Codex-Call-Rückfrage dazu blieb unbeantwortet, kein wiederholter Anruf nötig.
Formularintegration als3b18277 veröffentlicht; letzte geprüfteDEB fe01e1d vor dieser
Integration. Details NATIVE_MACROS_RUNTIME.md/READONLY_FORM_REQUERY_RUNTIME.md.
F129-Text korrigiert, fester Katalog456/386/70 unverändert. Weiter mit echten
Quellenaliasen und Selbstverknüpfungen im visuellen Entwurf.


Quellenaliase/Selbstverknüpfungen im visuellen Entwurf: wiederholte echte Tabellen/
Query-Views mit eindeutigem explizitem Alias, getrennte physische Herkunft und
SQL-Feldidentität, Metadaten immer von der tatsächlichen Quelle. Linker Selektor
bietet vorhandene Aliase und zugehörige Felder, Tooltip zeigt Herkunft. Summen-
dialog qualifiziert doppelte Feldnamen. Reale Eltern-Kind-LEFT-Verknüpfung gegen
unabhängiges SQL verglichen, Nullfälle/Alias mit Leerzeichen/Wiederöffnung und
Ablehnung mehrdeutiger aliasloser Wiederverwendung geprüft; Metadaten führen
keine Zeilenberechnung aus. Anwendung gebaut, bestehendes GUI-/SQL-Bündel2/2
(16,45s) bestanden ohne weitere Wiederholung. Keine neuen CTests oder neue
Original-/Windows-/Gesamtsuiten-/Paketnachweise. Details VISUAL_QUERY_DESIGNER.md;
F080-Beschreibung aktualisiert, Katalog456/386/70 unverändert. Makrokatalog-Fix
als657d6b3 veröffentlicht/verifiziert; letzte geprüfteDEB weiterhinfe01e1d.


Visuelle Verknüpfungen unterstützen nun INNER/LEFT/RIGHT mit1..16 UND-verknüpften
Feldvergleichen (=, <>, <, <=, >, >=) pro neuer Quelle, auch von vorhandenen
Aliasen. Zusammengesetzte Schlüssel werden tatsächlich als ON-Vergleiche erzeugt;
RIGHT erhält die Zeilen der neuen Quelle mit echten Nullwerten links. Vorgemerkte
Vergleiche plus vollständiges aktuelles Paar bilden die Bedingung; unvollständige/
unhinzugefügte Entwürfe müssen ergänzt oder explizit verworfen werden. Visuelle
Metadaten schreibenVersion2, Version1/alte Einfachvergleiche bleiben lesbar.
Anwendung gebaut; bestehendes GUI-/SQL-Bündel2/2 (15,54s) ohne Nachlauf bestanden.
Tatsächlicher nativer executeAccessQueryPlan-Output für zusammengesetzten RIGHT
JOIN gegen umgekehrten unabhängigen LEFT JOIN inklusive Null verglichen; <=-INNER
nach Wiederöffnung ebenfalls unabhängig geprüft. Keine neuen CTests oder neue
Original-/Windows-/Gesamtsuiten-/Paketnachweise. Details VISUAL_QUERY_DESIGNER.md.
Alias-Etappe9953a83 veröffentlicht; letzte geprüfteDEB weiterhinfe01e1d; jetzt
aktuelles Paket nach dieser zusammenhängenden Entwurfserweiterung prüfen.
F067/F068/F080/F148 und fester Katalog456/386/70 unverändert.


Visuelle Aktionsabfragen implementiert: Auswahl/Tabellenerstellung/Anfügen sowie
Aktualisieren/Löschen für einzelne physische Tabellen ohne Joins/Summen/DISTINCT/
Sortierung. Expliziter freier neuer Tabellenname bzw. reale Zielspaltenzuordnung;
Updatefelder mit begrenzten skalaren SQL-Ausdrücken und deklarierten Parametern,
WHERE aus tatsächlichen Kriterien. Vorschau/Übernahme erzeugen nur SQL und
Metadaten; ausschließlich Ausführen verwendet den bestehenden nativen DAO-
QueryDef-/Bindungs-/Aktionspfad. Reale Tabellen/Zeilen/Anfüge-Standardwerte sowie
Parameteraktualisierung und kriteriumsgebundene Löschung nachgewiesen. Speichern/
Wiederöffnung behalten Modus/Ziel/Zuordnung/Ausdrücke; Formular-/Berichtsquellen
bleiben Auswahl-only. Entwurfsformat schreibtVersion3, liestVersion1/2/3.
Initiales bestehendes GUI-/SQL-Bündel2/2 (16,23s) für Make/Append bestanden. Nach
Update/Delete nur betroffenes GUI-Ziel: Qt-Iterator-Baufehler und DELETE-*-Vorschau
korrigiert; endgültig1/1 (15,79s), Anwendung gebaut. Der zunächst fälschlich dem
Update zugeordnete Fehler lag bei der DELETE-Vorbereitung. Keine neuen CTests
oder neue Kern-/Original-/Windows-/Gesamtsuiten-/Paketnachweise für diese Stufe.
Details VISUAL_QUERY_DESIGNER.md; native Laufzeit-/Typ-/Metadatengrenzen bleiben.

Zusammengesetzte Joins als44b1f70 veröffentlicht/verifiziert. Geprüftes DEB dieses
Standes /workspace/work/visual-query-composite/package/linaccess-0.7.0-Linux.deb:
Build-ID, MDBTools1.0.1/SQLite3.54.1, Payload/Berechtigungen/keine Fixtures bestanden;
dieses Paket enthält noch nicht die anschließenden Aktionsentwürfe. F067/F068/
F080/F148 und fester Katalog456/386/70 unverändert. Weiter mit nativer TOP-Auswahl
im visuellen Entwurf; kein Gesamtabschluss, kein Abschlussanruf.


Native TOP-Steuerung im visuellen Auswahlentwurf: positive32-Bit-Zahl oder Alle
Datensätze(0), tatsächlicher nativer RANK-/LIMIT-Pfad einschließlich Grenzwert-
Gleichständen und ohne sichtbare Rangspalte. Gespeicherter Wert öffnet wieder,
Bruch-/negative/überlaufende Metadaten werden abgelehnt. Rohansichten __la_top_
werden aus Quellen-/Join-/Ziellisten ausgeschlossen; öffentliche gespeicherte
Abfrage bleibt verfügbar. TOP in diesem visuellen Pfad nur Auswahl; Aktionsmodus
darf den Wert nicht stillschweigend ignorieren. Native PERCENT-/DISTINCT-/Budget-
Grenzen bleiben. EntwurfsformatVersion4 schreibt, Version1/2/3/4 liest.
Anwendung gebaut; nur bestehendes betroffenes GUI1/1 (16,14s) bestanden. Tatsächliche
TOP2-Ergebnisse mit3 Grenzwert-Zeilen gegen unabhängige nth-Wert-SQL verglichen,
Ausgabe/Metadaten/Begrenzungslöschung/Callback-Nichtauswertung geprüft. Keine neuen
CTests oder zusätzliche Kern-/Original-/Windows-/Gesamtsuiten-/Paketnachweise.
Aktionsstufe als1481810 veröffentlicht/verifiziert; letzte geprüfteDEB44b1f70
vor Aktions-/TOP-Erweiterung. Details VISUAL_QUERY_DESIGNER.md. F067/F068/F080/
F148 und fester Katalog456/386/70 unverändert; Arbeit fortsetzen.


Exakte SQL-Currency-Parameterarithmetik: tabellenfreie, vollständig bewiesene
Currency/Money-Parameter-Ausdrücke mit +,-,*, Vorzeichen und Klammern verwenden
signed64-Ticks bis zu beiden Currency-Grenzen. Berechnete Ausgaben benötigen
eindeutige explizite Aliase; Parser64 Ebenen/4096 Knoten. Native integer-only-
Callbacks verwenden currencyBinary einschließlich Zwischenwert-/Überlaufprüfung
und ties-to-even, Null bleibt Null. DAO-Type5/Size8/Value sowie Anzeige/Clipboard
behalten exakte logische Werte. Bindung regeneriert den bewiesenen SQL-Text und
verweigert auch geänderte Quellausdrücke mit gleicher Ausgabenzahl. Bestehende
REAL-Feldgrenze, Rohzahlen/Mischtypen/FROM/WHERE/Aktionsspeicherung/Filter/Find
bleiben ausdrücklich begrenzt. Keine physische Currency-Speichererweiterung.
Anwendung gebaut; bestehendes Parameter-/Arithmetik-/Ergebnis-/DAO-Bündel4/4
(2,84s) ohne Nachlauf bestanden. Grenzwerte/Rundung/Null/Überlauf/Präzedenz,
DAO-EOF-Metadaten und manipulierter Quellbeweis tatsächlich geprüft. Keine neuen
CTests oder neue Original-/Windows-/Gesamtsuiten-/Paketnachweise dieser Stufe.
Details CURRENCY_PARAMETERS_RUNTIME.md/CURRENCY_ARITHMETIC_SQL_RUNTIME.md.
TOP-Entwurf alsc068ccd veröffentlicht/verifiziert; letzte geprüfteDEB44b1f70
vor Aktions-/TOP-/exakter Parameterarithmetik. Jetzt aktuelles Paket prüfen;
F014/F067/F068 und fester Katalog456/386/70 unverändert, Entwicklung fortsetzen.


Native Debug.Print/Debug.? ist implementiert: echte VM-Auswertung, Komma-/Semikolon-
Listen, Spc/Tab, einmalige Links-nach-rechts-Auswertung und erhaltene Teilausgabe
bei späterem Fehler. Gemeinsamer flüchtiger thread-lokaler Puffer für tatsächliches
Ansichtsmenü-Dock und Debugger-Pane, auch während Stop/Continue; explizites Löschen,
Scrollposition erhalten. Grenzen1MiB/8192 Zeilen, Listen64KiB/256 Elemente/64 Ebenen,
UTF-8-sichere Aufbewahrung, UTF-16-Spalten. Kein Dateidruck/Immediate-Evaluator oder
vollständiger Windows-/Locale-Nachweis. Details NATIVE_VBA_DEBUGGER_RUNTIME.md.
Anwendung gebaut; bestehendes Kern-/GUI-Bündel2/2 (17,43s) bestanden, Kern ohneQt
verlinkt. Keine neuen CTests, kein neuer Original-/Windows-/Gesamtsuitennachweis.
Letzte veröffentlichte Vorstufe0e4f084; dazu DEB unter
/workspace/work/exact-currency-expressions/package/linaccess-0.7.0-Linux.deb mit
scripts/check-deb.sh einschließlich BuildID/SQLite3.54.1/MDBTools1.0.1 geprüft.
Dieses Paket liegt vor Debug.Print. F060 und Katalog456/386/70 unverändert.
Entwicklung ohne Zwischenabschluss fortsetzen.


Native Property Get/Let/Set mit getrennten Zugriffsmethoden ist implementiert:
unveränderte Modulquelle/Zustandsidentität, echte gleichnamige Rückgabebindung,
End/Exit Property, unqualifizierte Modulzugriffe und öffentliche Form-/Report-
Objektzugriffe einschließlich indizierter ByVal-Parameter. Setterwert immerByVal,
Objektidentität/Nothing bleiben erhalten; Erkennung führt keine Getter aus.
Signatur-/Sichtbarkeits-/Eigentümerprüfungen und tatsächliche Fehlerweitergabe;
Modulauswahl/Beobachtungen unterscheiden Get/Let/Set. Grenzen ausdrücklich in
NATIVE_VBA_PROPERTIES.md: keine vollständigen Klassen/ByRef-Indizes/ParamArray/
Default-Properties/Windows-Parität. Anwendung gebaut; bestehendes Kern-/GUI-
Bündel2/2 (16,12s) nach Korrektur der tatsächlich gefundenen Bindungsfehler bestanden.
Keine neuen CTests. Originalinventar32 VBA-Artefakte/0 Property-Deklarationen;
keine zusätzliche Originalausführung behauptet. Vorstufe Debug.Print als469ad64
veröffentlicht und Remote-Baum/privat verifiziert. Letzte geprüfteDEB0e4f084 liegt
vor Debug.Print/Properties. F042/F043/F050 und Katalog456/386/70 unverändert;
anschließend statische Prozedurlebensdauer umsetzen, ohne Zwischenabschluss.


Statische Sub-/Function-/Property-Prozeduren besitzen jetzt tatsächliche lokale
Modullebensdauer: Dim-Skalare/Objekte/feste und dynamische Arrays/feste Strings,
implizite Variant-Bindungen ohne Option Explicit, Alias-Synchronisierung bei
ByRef/Rekursion/Fehlern. Parameter, Rückgabe, Err und Konstanten bleiben proAufruf.
Initial-/Trailing-Static unterstützt, doppelte Angabe vor Effekten abgelehnt.
Def-Typ-Direktiven vor Effekten und implizite Typzeichen ausdrücklich3251 statt
falschem Variant-Erfolg; vollständige Compile-/Klassen-/Reset-Parität bleibt offen.
Details STATIC_LOCAL_RUNTIME.md/NATIVE_VBA_PROPERTIES.md. Anwendung gebaut;
bestehender Kern3,02s und GUI1/1 16,68s nach Korrektur des ungültigen null-Fenster-
besitzers im technischen Host bestanden. Spätere Typgrenzen nur im betroffenen
Kern1/1 (2,95s) geprüft; kein zusätzlicher GUI-/Original-/Windows-/Gesamtsuitenlauf.
Originalinventar32 VBA-Artefakte/0 statische Prozedurkandidaten. Keine neuen CTests.
Vorstufe Properties39e3cb2 auf privatem Branch veröffentlicht/Remote-Baum verifiziert.
F042/F043 und Katalog456/386/70 unverändert; jetzt Paket aktualisieren und danach
parameterlose Prozedur-/Property-Deklarationen ohne Klammern implementieren.
Kein Zwischenabschluss. Letzte geprüfteDEB0e4f084 liegt vor Debug.Print/Properties/
statischen Prozeduren; Original-/Fixture-Hashes7/7 zuvor unverändert nachgewiesen.


Parameterlose Sub-/Function-/Property-Get-Deklarationen ohne Klammern sind jetzt
im echten Interpreter und Modul-Inventar unterstützt, einschließlich typisierter/
Variant-Getter und Initial-/Trailing-Static. Norm MS-VBAL5.3.1; ursprüngliche
Quellidentität bleibt erhalten. Parameter ohne erforderliche Klammern, typisierte
Sub-/Setter-Rückgaben und Setter ohne Wertparameter verweigern vor Effekten.
Anwendung gebaut; erstes bestehendes Kern-/GUI-Bündel2/2 (15,90s) bestanden;
keine neuen CTests, Originalinventar32/0 Headerkandidaten, kein neuer Original-/
Windows-/Gesamtsuitennachweis. Details NATIVE_VBA_PROPERTIES.md.
Vorstufe statische Prozeduren46297bf veröffentlicht, Remote-Baum/privat verifiziert.
Dazu neueDEB /workspace/work/native-language-package/package/linaccess-0.7.0-Linux.deb
mit scripts/check-deb.sh geprüft: BuildID identisch, MDBTools1.0.1/SQLite3.54.1,
Payload/öffentliche Rechte/keine Fixtures. Paket liegt vor klammerlosen Deklarationen.
F040/F042/F043 und fester Katalog456/386/70 unverändert. Anschließend implizite
lokale Variablen auch in gewöhnlichen Prozeduren ergänzen (frisch proAufruf,
OptionExplicit und Typgrenzen beachten), gezielt prüfen und ohne Abschluss weiter.


Implizite lokale Variablen jetzt auch in gewöhnlichen nativen Prozeduren: echte
frische Empty-Variant-Bindungen proAufruf nach bestehender Namensauflösung;
statische Prozeduren behalten Modulspeicher. Unterstützte externe öffentliche
Funktionen werden aufgelöst, Sub-/Prozedurnamen nicht durch erfundene Variablen
ersetzt. Option Explicit verhindert Anlage auch beim Schreiben; native Laufzeit-
prüfung statt vollständiger Access-Compile-Diagnostik. Def-Typ-Direktiven vor
Effekten und implizite Typzeichen verweigert;4096 implizite Namen proRahmen.
Details VBA_IMPLICIT_LOCALS.md. Anwendung gebaut; GUI16,32s bestanden, nur zuvor
betroffene Kern-/Original-Ziele nach technischen Fixturekorrekturen erneut geprüft:
2/2, Kern3,11s und unveränderter real_kassa_dao_event124,12s. Keine neuen CTests,
kein GUI-Nachlauf/Windows-/Gesamtsuitennachweis. Originalinventar32/19OptionExplicit/
13ohne/0Def; unveränderter Ereignislauf bleibt an2485/1023 fehlender Makroquelle,
kein vollständiger Verkauf behauptet. Vorstufe klammerlose Deklarationen b05c240
veröffentlicht/privater Remote-Baum verifiziert; letzte geprüfteDEB46297bf vor
klammerlosen Deklarationen/impliziten normalen Variablen. F037/F042/F043 und
Katalog456/386/70 unverändert. Anschließend objektwertige Getter/default-Member-
Weiterleitung nach MS-VBAL5.4.3.8/5.6.13 ergänzen; ohne Zwischenabschluss weiter.


Objektwertige Getter unterstützen jetzt native Default-Member-Weiterleitung:
parameterlose Property-/samelokale Function-Ergebnisse mit positionalen Indizes,
Get-only Objektproperties für native Dictionary-Let/Set und Default-Wertzuweisung
an tatsächliche Textfelder, auch über indizierten Getter. Echte Setter haben Vorrang;
Getter/Argumente/RHS nicht doppelt ausgeführt, Objektidentität erhalten, native
Fehler13/91/Getterfehler weitergegeben. Eigentümerschließen im Getter verhindert
Default-Zuweisung mit2450. Grenzen/default-Queues/Arrays/COM/ByRef-Fassade bleiben
explizit, Norm MS-VBAL5.4.3.8/5.6.13, Details NATIVE_VBA_PROPERTIES.md.
Anwendung gebaut; erstes Laufzeit-Kern-/GUI-Bündel2/2 (17,19s) bestanden. Keine
neuen CTests. Originalinventar32/0 Properties/0 parameterlose Objektfunktionen;
kein neuer Original-/Windows-/Gesamtsuitennachweis. Implizite normale Variablen
als3c61be9 privat veröffentlicht/Remote-Baum verifiziert; Original-/Fixture-Hashes
7/7 unverändert geprüft. Letzte geprüfteDEB46297bf liegt vor klammerlosen
Deklarationen/normalen impliziten Variablen/Default-Weiterleitung. Katalog456/386/70
und F042/F043/F050 unverändert. Anschließend ByRef-Property-Indizes für direkte
samelokale Get/Let/Set-Aufrufe mit echter Rückbindung ergänzen; externe Objekt-
fassade weiterhin explizit begrenzen. Ohne Zwischenabschluss fortsetzen.


ByRef-Indizes bei direkten samelokalen Property-Get/Let/Set-Aufrufen sind echt
rückgebunden: vorhandener Prozedurargument-Binder statt abgetrennten Wertkopien,
einschließlich Array-Elementen, temporären Ausdrücken und Fehlerseitenwirkungen.
Setterwert bleibtByVal, Objektidentität/Typ-/Array-Allokationsguards bleiben.
Auch Default-Wertzuweisung über indizierten Objektgetter nutzt diese Rückbindung;
realer Textfeldnachweis. Externe Objektfassade bleibt bei ByRef-Indizes explizit3251,
keine erfundene Rückbindung. Details NATIVE_VBA_PROPERTIES.md. Anwendung gebaut;
erstes bestehendes Kern-/GUI-Bündel2/2 (17,35s) bestanden, keine neuen CTests oder
Original-/Windows-/Gesamtsuitennachweise. Vorstufe Default-Weiterleitung e588edc
privat veröffentlicht/Remote-Baum verifiziert. Letzte geprüfteDEB46297bf vor
klammerlosen Deklarationen/normaler impliziter Bindung/Default-/ByRef-Erweiterung.
F042/F043/F050 und Katalog456/386/70 unverändert. Nächste Stufe: optionale und
ParamArray-Indizes in Properties einschließlich besonderem letzten ByVal-Wert-
parameter korrekt binden; danach Paket aktualisieren. Ohne Zwischenabschluss weiter.


Optionale und ParamArray-Property-Indizes sind implementiert: letzter erforderlicher
Setterwert getrennt von Indexliste, fehlende Positionen vor dem Wert ergänzen,
API-Argumentvektorlänge und ByRef-Indexrückbindung erhalten. Unterschiedliche
Get-/Let-Defaults nach MS-VBAL5.3.1.5/5.3.1.7; Cell(,8)/Cell() und benannte optionale
Getter einschließlich echter Formularfassade. ParamArray enthält nur Indizes,
RHS stets actual-last/ByVal, auch bei leerem Array oder zusätzlichen Argumenten;
Objektidentität erhalten. Variadische Signaturen verhindern erfundene benannte
ParamArray-Bindung; Mutationen/Weiterleitung/Holes explizit begrenzt. Ausgelassene
Arrayindizes werden vor Bodyeffekten verweigert, nicht0 geraten. Details
NATIVE_VBA_PROPERTIES.md. Anwendung gebaut; erstes bestehendes Kern-/GUI-Bündel
2/2 (16,60s) bestanden. Keine neuen CTests, kein zusätzlicher Original-/Windows-/
Gesamtsuitennachweis. ByRef-Vorstufe bc52b22 privat veröffentlicht/Baum verifiziert.
Letzte geprüfteDEB46297bf liegt vor mehreren Sprachstufen; jetzt Paket aktualisieren.
Katalog456/386/70 und F042/F043/F050 unverändert; anschließend typisierte Array-
Funktions-/Getterrückgaben und echte Leseindizierung ergänzen. Ohne Abschluss weiter.


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


Undo/Redo vorgemerkter Mehrfachlöschungen: TABLE_DELETE_HISTORY.md beschreibt die
bestehende Einzelaktionshistorie für saubere gespeicherte Schlüsselzeilen.
Strg+Z/Strg+Y ändern nur den Cache; genaue nicht zusammenhängende Löschmarkierungen
und Rücknahme/Wiederholung nach transaktionalem Fremdschlüsselfehler geprüft.
App gebaut; bestehendes GUI-/Währungsfeld-Bündel2/2 in19,17s (18,68/0,49).
Keine neuen Tests/IDs, keine Aussage zu Windows-Ereignissen oder committed SQL
Undo.456/386/70 unverändert. Vorstufe ba256ac privat remote/Baum geprüft.
Letztes Original118,41s/DAO2,75s/7Hashes liegt vor diesem Modellumfang; geprüftes
DEB9affc5a bleibt älter. Ohne Zwischenabschluss mit sauberer gespeicherter
Einzeldatensatzlöschung und Invalidierung der Aktionshistorie weiterarbeiten.


Saubere gespeicherte Einzeldatensatzlöschungen nutzen ebenfalls die geprüfte
Cache-Aktionshistorie aus TABLE_DELETE_HISTORY.md. Zelländerungen invalidieren
alte Wiederholungen; doppelte/ungültige Zeilen und schreibgeschützte Quellen
bleiben unverändert. Dirty-/schlüssellose Einzelpfade behalten bestehendes
Verhalten. App gebaut; bestehendes GUI-/Währungsfeld-Bündel2/2 in19,88s
(19,39/0,49), keine neuen CTests;456/386/70 unverändert. Vorstufe22a29fa privat
remote/Baum verifiziert. Weiter ohne Zwischenabschluss: native VarP/StDevP
in Berichts-/SQL-Aggregaten und DVarP/DStDevP in echten Domainaufrufen ergänzen.


Native Populationsstatistik: POPULATION_STATISTICS.md beschreibt VarP/StDevP in
SQL und Berichtssteuerelementen samt PDF/Requery, DVarP/DStDevP für echte Domains
inklusive benannter/Application-Aufrufe und die realen visuellen QBE-Funktionsmodi.
Null/leer/Einzelwert, große Offsets, Gruppen/DISTINCT/Parameter und Fehler geprüft.
App gebaut; erstes Bündel Kern5,32s/Bericht12,97s bestanden. Neue SQL-Testannahme
StDevP-Überlauf korrigiert (StDevP kann1e308 liefern, VarP überläuft): nur betroffenes
Abfragemodell erneut0,13s. Designer-Fixture erwartet zusätzlich verborgene Where-
Quellfelder; nur GUI erneut18,41s. Keine neuen CTests/Vollsuiten-/Windowsbehauptung;
456/386/70 unverändert. Sieben Hashes unverändert; Vorstufe c119f5e privat remote
und Baum geprüft. Anschließend aktuelles Debianpaket bauen/gegen echte Binärdatei
prüfen und unabhängige Implementierung fortsetzen; kein Zwischenabschluss.


Native Folder.Name-Zuweisung in NATIVE_FILE_SYSTEM_OBJECT.md: tatsächlicher
Ordner wird mit vorhandenem atomarem No-Replace-Helfer im bisherigen Elternpfad
umbenannt; gehaltene Objektidentität/Alias und Pfad erst nach Erfolg aktualisiert.
Unicode, Dateiinhalt, zweimaliges Umbenennen und unveränderte Pfade bei Ziel58,
Traversal3251 und Null94 geprüft. App gebaut; bestehender Kern im ersten Lauf1/1
in5,20s; keine weiteren Tests/IDs. F091 und456/386/70 bleiben offen/unverändert.
Vorstufe c8bd82b privat remote/Baum geprüft. Dessen aktualisiertes DEB liegt in
/workspace/work/population-statistics/package/linaccess-0.7.0-Linux.deb: tatsächliche
Build-ID, MDBTools1.0.1, SQLite3.54.1/nestedJOIN/NULL/counter, Rechte/Desktop/keine
Fixtures geprüft; danach DAO2,81s+Original120,75s,2/2 in123,57s. Bekannte Makrogrenze
bleibt. Sieben Hashes vor Paketlauf unverändert. Paket liegt vor Folder.Name.
Ohne Zwischenabschluss weitere allgemeine native Objekt-/Entwurfsanforderungen
implementieren; keinen Abschlussanruf für diese Etappe auslösen.


Native Ordnerverlagerung: Folder.Move und FileSystemObject.MoveFolder benutzen
den vorhandenen atomaren No-Replace-Helfer für einzelne tatsächliche Ordner und
exakte Zielpfade; Name-/Alias-Pfad erst nach Erfolg. Benannte Source/Destination,
Unicode, tatsächliche Unterdatei, frische Collections und unveränderte Pfade nach
58/3251/94 geprüft. App gebaut; bestehender Kern erster Lauf1/1 in5,26s. Keine
neuen Tests/IDs; F091/456/386/70 unverändert. Vorstufe1d8f606 privat remote/Baum
verifiziert. Geprüftes DEBc8bd82b unter /workspace/work/population-statistics/package
liegt vor Name/Move; letztes Original120,75s/DAO2,81s ebenfalls davor. Weiter ohne
Zwischenabschluss: feste numerische TRANSFORM-Kreuztabellen um Avg/Var/StDev/
VarP/StDevP mit nachgewiesener Double-Ausgabetypisierung erweitern.


Feste statistische Kreuztabellen: STATISTICAL_CROSSTABS.md dokumentiert tatsächliche
TRANSFORM Avg/Var/StDev/VarP/StDevP über nachgewiesene numerische physische Felder,
inklusive Aliase, Parameter, Null/fehlende Spalten und stabile Double-Metadaten.
Zuerst gemeinsames Schema-/Abfragemodell-Bündel; Modell0,12s bestand. Zwei konkrete
Implementierungslücken geschlossen: Feldkonverter REAL/DOUBLE und tatsächliche
DAO-Zuordnung nachgewiesener DOUBLE→Typ7. Danach nur betroffenes Schema+DAO2/2 in
2,98s (0,13/2,85); App gebaut. Vorhandene Count/Min/Max-Währungs-/Datums-/Ganzzahl-
Fälle bleiben im gleichen Schemafall geprüft. Keine neuen Tests/Katalogpunkte;
F068/F122 und456/386/70 unverändert. Vorstufe8cbf0af privat remote/Baum geprüft.
Geprüftes DEBc8bd82b und Original120,75s liegen vor diesen späteren Änderungen.
Ohne Zwischenabschluss weitere allgemeine SQL-/Objektmodellanforderungen umsetzen.


Feste Kreuztabellensummen: FIXED_CROSSTAB_SUMS.md beschreibt TRANSFORM Sum mit
nachgewiesenen Double/REAL- oder nativen Currency-Quellen. Geld verwendet den
vorhandenen exakten Currency-Akkumulator, keine rohe SQLite-Gleitkommasumme.
Tatsächliche Double-/Currency-Werte, Dezimalkompensation, Null und DAO-Typ7/5 im
bestehenden Schemafall geprüft: erster Lauf1/1 in0,13s; App gebaut. Integer/Single/
Decimal-Promotion und größere physische MONEY-Bereiche bleiben ausdrücklich offen.
Keine neuen Tests/IDs;456/386/70 unverändert. Vorstufe103af93 privat remote/Baum
geprüft. Geprüftes DEBc8bd82b bleibt älter. Ohne Zwischenabschluss weiter an echten
festen Pivot-Auswahlen First/Last arbeiten, ohne nichtpassende Null-Zeilen als
falsche erste/letzte Pivotwerte einzubeziehen.


Feste First/Last-Pivots: FIXED_CROSSTAB_SELECTORS.md beschreibt tatsächliche
Filterung passender Pivotzeilen vor der Auswahl, inklusive echter passender
Nullwerte. INTEGER2/Currency/Datum und reale DAO-Typen/Variantwerte gegen separate
native Zeilenscans geprüft; keine Behauptung über identische Windows-Besuchsfolge.
App gebaut; vorhandener Schemafall erster Lauf1/1 in0,17s. Nach zusammenhängenden
SQL-Erweiterungen originales Abfrageinventar einmal geprüft1/1 in47,83s:100
klassifiziert,78 ausgeführt (73 gewöhnlich/5 mit technischen Parametern ohne
Access-Ergebnisoracle),22 Grenzen (3 Quellspalten/9 Abhängigkeiten/10 Jet-Bindung).
Alle sieben Upload-/Fixturehashes unverändert. Vorstufe49a60c7 privat remote/Baum
geprüft. Keine neuen Tests/IDs; F068/456/386/70 unverändert. Geprüftes DEBc8bd82b
liegt vor späteren Ordner-/Kreuztabellenänderungen. Ohne Zwischenabschluss
weitere allgemeine DAO-/SQL-Metadatenlücken implementieren.


Nachgewiesene Boolean-DAO-Ergebnisse: DAO_BOOLEAN_RESULTS.md dokumentiert die
fehlende Abbildung tatsächlicher BOOL/BOOLEAN-Deklarationen auf dbBoolean1 und
Size1. Echte QueryDef.Fields ohne Modulausführung sowie Recordset-True/False-
Varianten, vorhandene Currency/Date/Long- und Identitätsgrenzen im gleichen DAO-
Ziel geprüft. App gebaut; Test-Compilerfehler (fehlender Namensraum) präzise
korrigiert, erster tatsächlicher DAO-Lauf1/1 in2,78s. Keine neuen Tests/IDs und
keine Boolean-Inferenz aus beliebigen Integern. F121/F122/456/386/70 unverändert.
Vorstufeff933a1 privat remote/Baum geprüft. Letztes Originalinventar100/78/22 und
sieben Hashes liegen vor diesem kleinen Metadatenumfang; geprüftes DEBc8bd82b
bleibt älter. Ohne Zwischenabschluss allgemeine native SQL-Funktionen fortsetzen.


Native SQL-Zahlenkonvertierung: SQL_NUMERIC_CONVERSIONS.md dokumentiert tatsächliche
CBool/CByte/CInt/CLng/CSng/CDbl-Callbacks über vorhandene VBA-Konvertierung sowie
Boolean/Byte/Integer/Long/Single/Double-Abfrage-/DAO-Metadaten. Rundung, Grenzen,
Verschachtelung, Parameter, WHERE und präzise Fehler im gemeinsamen bestehenden
Bündel geprüft. DAO erster Lauf2,92s bestand; SQLite-BLOB-Literal war im neuen
Test keine Access-SQL-Grammatik: durch tatsächliches Binärfeld ersetzt, nur
Abfragemodell erneut1/1 in0,12s. App gebaut. Keine neuen Tests/Katalogpunkte;
Datums-/Objekt-/Decimal-Transport, CStr und vollständige Regionalsemantik bleiben
offen.456/386/70 unverändert. Vorstufebc7cbf9 privat remote/Baum geprüft. Ohne
Zwischenabschluss nachgewiesene Count-Ergebnistypen in DAO/SQL ergänzen und danach
aktuelles Debianpaket bauen; geprüftes DEBc8bd82b liegt noch vor späteren Umfängen.


Count-Ergebnistypen: COUNT_RESULT_TYPES.md dokumentiert nachgewiesene Long/
INTEGER4-Metadaten tatsächlicher Count-Ausdrücke, QueryDef.Fields und echter DAO-
Felder (Typ4/Size4/Long-Variante). Gesamt-/Null-/Feld-/Gruppen-/Parameter-/leere
Counts und metadatenreine Projektzustandserhaltung im ersten bestehenden Bündel
2/2 in3,12s (DAO2,99/Abfragemodell0,13) geprüft; App gebaut. Keine neuen Tests/IDs,
keine Ableitung aus beliebigen Integerwerten; F068/F121/F122/456/386/70 unverändert.
Vorstufe3fa14b9 privat remote/Baum geprüft. Anschließend ohne Zwischenabschluss
aktuelles Debianpaket bauen und prüfen, GUI/Originalpfad nach zusammenhängenden
SQL-/DAO-Änderungen einmal integrieren; geprüftes DEBc8bd82b bleibt bis dahin älter.


Direkte deklarierte Parameterprojektionen: PARAMETER_RESULT_TYPES.md dokumentiert
Byte/Short/Long-Integer/Single/Double/Boolean-YesNo-Metadaten einschließlich
normalisierter Aliase, kompletter Klammern, echter DAO-Typen/Größen/Varianten und
Null. Nur direkte Parameter annotiert; Byte-Negation erhält keinen falschen Byte-
Typ. App gebaut; erstes bestehendes DAO-/Abfragemodell-Bündel2/2 in3,02s
(2,89/0,13), keine neuen Tests/IDs;456/386/70 unverändert. Vorstufecc19000 privat
remote/Baum geprüft. Dessen aktualisiertes DEB /workspace/work/count-result-types/
package/linaccess-0.7.0-Linux.deb besteht echte Build-ID/MDBTools1.0.1/SQLite3.54.1/
nestedJOIN/NULL/counter/Rechte/Desktop/keine-Fixtures-Prüfung; danach GUI19,32s und
Original119,69s,2/2 in139,02s, alle sieben Hashes unverändert. Bekannte Makrogrenze
bleibt. Paket und Integration liegen vor dieser Parameterannotation. Ohne
Zwischenabschluss native SQL-Potenzoperatoren mit echter VBA-Präzedenz fortsetzen.


Nativer SQL-Potenzoperator: SQL_POWER_OPERATOR.md dokumentiert tatsächliche
^→VBA-Auswertung mit strukturellen Atomen, Linksassoziativität/Vorzeichen nach
MS-VBAL5.6.9.1, Klammern/Funktionen/Parametern/Feldern/WHERE, Double-/DAO-Typen und
Null. Quotierte/literale ^ bleiben intakt, Originalüberschriften bleiben sichtbar.
App gebaut; erstes bestehendes DAO-/Abfragemodell-Bündel2/2 in3,17s (3,02/0,15).
Vorhandene Zero-negativ-/Unterlaufgrenze3251 beibehalten; kein Windows-/voller SQL-
AST-/Datum-/Decimal-Anspruch. Keine neuen Tests/Katalogpunkte;456/386/70 unverändert.
Vorstufeca0ebc0 privat remote/Baum geprüft. Geprüftes DEBcc19000 unter
/workspace/work/count-result-types/package/linaccess-0.7.0-Linux.deb und GUI19,32/
Original119,69/7Hashes liegen vor Parameterannotation/Potenz. Ohne Zwischenabschluss
weiter echte typisierte numerische SQL-Arithmetik und Divisionsfehler umsetzen.


Typisierte numerische SQL-Arithmetik: TYPED_SQL_ARITHMETIC.md beschreibt echte
Double-Operationen, gebrochene Ganzzahldivision, nachgewiesene Feld-/Parametertypen,
Null, WHERE/HAVING und reale SQLite-Ursprünge auch bei Sternprojektionen. Read-only
Treiber-Cache prüft Schema/Katalog/Definition/Speicher; keine Werte/VBA-Module.
0/0 und numerischer Überlauf melden normgemäß6, sonstige Division durch0 meldet11;
SQL-Fehler11 jetzt korrekt weitergereicht. Core5,55s/GUI18,91s bestanden vor letzten
SQL-Verfeinerungen; zuletzt DAO3,00s und Abfragemodell0,19s. Import-Cachefixture
verwendet tatsächlichen REAL-Import mit Double-Metadaten samt normaler Logtabelle;
DAO erstellt physisch DOUBLE. Metadatenänderung ohne Schemaversion, TEMP-Schatten,
reale Untertypen und unveränderte Überschriften geprüft. Normale indizierte
ID-Prädikate bleiben unverpackt, tatsächlicher Scan-Zähler0 bestätigt. Keine
neuen Tests/IDs; F041/F067/F068/F071/F121/F122/456/386/70 unverändert. Vorstufe2fdd9dc
privat remote/Baum geprüft. Geprüftes DEBcc19000 liegt vor späteren Umfängen.
Danach ohne Zwischenabschluss aktuellen Paket-/Original-/Inventarlauf ausführen
und weitere native Zahlentyp-/Operatorschritte fortsetzen. Kein Abschlussanruf.


Tatsächliche Report-RunningSum-Ausgabe: REPORT_RUNNING_SUM.md dokumentiert
Over-All-Detailpräfixe aus tatsächlichen geordneten numerischen Feldern und =1,
Currency-/Decimal-Summenpfad, Formatierung, berechnete Abhängigkeiten, Mehrseiten-
PDF und wiederholte Ausgabe ohne Doppelzählung. Steuerelement-/Datensatzwerte
werden wiederhergestellt; Null/Gruppen/andere Ausdrücke/Running-CanGrow bleiben
explizit gesperrt. Vorhandene Druck-/Ereignisgrenzen gelten. App/Layout gebaut,
ein bestehender gebündelter access_2003_layouts-Lauf13,23s bestanden, echte PDF-
Texte verglichen und atomarer Zielschutz geprüft. Keine neuen Test-/Katalog-IDs;
456/386/70 unverändert. Vorstufe2db95826 privat remote/Baum geprüft. Geprüftes
DEB8fd1d50 liegt vor nativen Operatoren und RunningSum. Ohne Zwischenabschluss
native Berichtsentwurfsbedienung für laufende Feldsummen fortsetzen.


Native RunningSum-Berichtsbedienung: optionales running_sum_field im echten
Berichtsentwurf, vorbereitete tatsächliche Ergebnisspalten/-typen ohne Zeilen/
Benutzerfunktionsausführung, fehlend3265/nichtnumerisch3251. Speicherung/Wieder-
öffnung, laufendes Detailfeld ohne CanGrow und unveränderte andere Textfelder,
VBA-/Quell-/Seitenmetadaten-Zusammenführung geprüft. Präfixumfang auf1MioZellen
begrenzt. App/GUI/Layout gebaut; ein bestehendes Bündel2/2 in32,05s (19,09/12,96)
bestanden, PDF-Zeilen mit tatsächlichen Ordinalen1/2/3 und Currency-Präfixen geprüft.
Keine neuen Test-/Katalog-IDs;456/386/70 unverändert. Vorstufe3dd9019c privat remote/
Baum geprüft. Ohne Zwischenabschluss jetzt aktuellen DEB-/Originalpfadnachweis
für beide Operator- und RunningSum-Umfänge bündeln, danach weiter implementieren.


Gebundene native Arithmetik: BOUND_NUMERIC_ARITHMETIC.md dokumentiert tatsächliche
Deklarationskategorien getrennt von VarType, lokale/Argument-/Alias-/Funktions-
und Arrayelementtypen, Literale/unterstützte explizite Konvertierungen und CVar.
Festes Integer/Long/Single meldet Überlauf6, nachgewiesene Variant-Werte erweitern
sich Long/Double; Unary-Minima und Date-Grenzen unterscheiden feste/Variant-Typen.
Erster gezielter Lauf fand echten ParamArray-Bindungsfehler99*1000: Element ist
Variant, darf keinen alten RHS-Static-Tag übernehmen. Tatsächliche Arraylesegrenze
korrigiert. Allgemeine Host-/Member-/Cache-/rohe Lokalwerte ohne Deklarationsresolver
bleiben ungebunden; interne Hilfskonvertierungen erzeugen keine Quell-Deklaration.
CVar erhält tatsächliche Daten-/Objektidentität, liest kein ScalarDefault hinzu.
Variant-Byte-Überlauf, vollständige Intrinsics/Member/Logik/Integerdivision-
Provenienz/Misch-Currency/AST/IEEE/Zusatzsuffixe/Windows bleiben offen.
App/Core/DAO/Abfragemodell/GUI gebaut. DAO2,93s/Modell0,19s bestanden zunächst;
nach Array-/Hostkorrektur Core5,67s/GUI19,70s gemeinsam2/2 in25,37s; letzte rohe
Lokalgrenze nur Core5,47s erneut bestanden. Testhelper-Final-Ableitung kompiliert
zunächst nicht, vor bestandenem Lauf behoben. Keine neuen CTest-/Katalog-IDs;
456/386/70 unverändert. Vorstufe7d78d04c privat remote/Baum geprüft. Letztes DEB/
Originalinventar46,35/Originalpfad120,51/7Hashes liegt auf5df3cba2 vor Literal- und
Deklarationsänderungen. Ohne Zwischenabschluss native numerische Intrinsic-
Rückgabebindungen und weitere tatsächliche Anforderungen fortsetzen.


Numerische Intrinsic-Rückgabebindungen: NUMERIC_INTRINSIC_BINDINGS.md ergänzt
tatsächliche deklarierte Variant-Ergebnisse für Abs/Int/Fix/Round/Sgn und
Choose/Switch/IIf/Nz; ausgewählte feste Integer-Argumente machen diese Funktionen
nicht selbst statisch. Sqr/Exp/Log/Sin/Cos/Tan/Atn sind nachgewiesenes Double.
Unveränderte eigentliche Zahlen-/Fehler-/Präzisions-/Eager-Pfade, generische Host-
Aggregate ohne fälschlich übernommene Cache-Deklaration, Access-Eval unverändert.
App/Core/DAO/Abfragemodell gebaut; erstes bestehendes Bündel3/3 in8,36s
(5,18/2,98/0,20) bestanden. Keine neuen CTest-/Katalog-IDs;456/386/70 unverändert,
keine Voll-Varianten-/AST-/Windows-Parität. Vorstufef38f6970 privat remote/Baum
geprüft. Letztes geprüftes DEB5df3cba2 liegt vor Literal-/Deklarationsumfängen.
Ohne Zwischenabschluss aktuellen Paket-/Originalnachweis für diese Umfänge
bündeln und danach weitere tatsächliche Anforderungen implementieren.


Native typisierte Logik: NATIVE_TYPED_LOGIC.md dokumentiert Nicht-Null-And/Or/
Xor/Eqv/Imp mit Byte-/Integer-/Boolean-/Long-Ergebnissen und tatsächlicher8Bit-
Eqv/Imp-Maske. Not/Binäroperationen tragen nachgewiesene feste/Variant-Bindungen
in Folgearithmetik; vorhandene Nullpfade und Access-Eval-Sonderpfad unverändert.
MS-VBAL-Allgemeintabelle Or/Null ist kein vollständiges Access-Orakel: tatsächliche
Referenz Null Or False ergibtNull. Unbelegte numerische Imp/Null-Fälle bleiben
bewusst offen. Metadatenlesepfad vermeidet Typstringkopie ohne Auswahländerung.
App/Core/DAO/GUI gebaut; erstes bestehendes Bündel3/3 in27,29s (5,34/18,72/3,22)
bestanden. Tatsächliche native Byte-VBA-Funktion über SQL/DAO liefertEqv252 mit
Type2/Size1/Byte-Value. Keine neuen CTest-/Katalog-IDs;456/386/70 unverändert.
Vorstufe4cd0f848 privat remote/Baum geprüft und aktuellstes geprüftes DEB unter
/workspace/work/numeric-intrinsic-bindings/package/linaccess-0.7.0-Linux.deb:
Build-ID/SQLite3.54.1/MDBTools1.0.1/Inhalt/Rechte/Desktop/keine-Fixtures bestanden,
danachGUI19,97/Inventar46,77/Original118,66,3/3 in185,40s;100/78/22 und7Hashes
unverändert. Makroquellengrenze bleibt. Paket liegt vor Logikumfang. Ohne
Zwischenabschluss native Ganzzahldivision/Modulo und weitere Anforderungen
fortlaufend implementieren, kein Abschlussanruf vor tatsächlichem Gesamtabschluss.


Native Ganzzahldivision/Modulo: INTEGRAL_DIVISION_TYPES.md dokumentiert tatsächliche
Byte-/Integer-/Long-Ergebnisse für unterstützte integrale Operandentypen und
Deklarations-/Variant-Provenienz für Folgearithmetik. Feste signed-Minima bleiben
Überlauf6, Variant-Integer/Long erweitern Long/Double. Bisherige geprüfte Long-
Eingabekonvertierung, int64-Zwischenrechnung für32Bit-Werte, Abschneiden Richtung0,
vorzeichenbehaftetesMod, Null und Nullteiler11 bleiben. Gemischte weitere Typen,
Suffixe/64Bit/Windows nicht pauschal behauptet. App/Core/DAO gebaut; erstes
bestehendes Bündel2/2 in8,72s (5,55/3,17) bestanden. Tatsächliche native Byte-
Funktion prüft Quotienten-VarType und liefertMod1 über SQL/DAO Type2/Size1/Byte.
Keine neuen CTest-/Katalog-IDs;456/386/70 unverändert. Vorstufe9226c05e privat remote/
Baum geprüft. Letztes geprüftes DEB4cd0f848 liegt vor Logik-/Divisionsumfang.
Ohne Zwischenabschluss genaue gemischte Currency-/Integral-Arithmetik und danach
weitere offene Anforderungen fortlaufend umsetzen. Kein Abschlussanruf.


Gemischte native Currency-/Zahlenarithmetik: CURRENCY_NUMERIC_ARITHMETIC.md
beschreibt +,- mit ganzzahligen/Single/Double/Boolean/Empty-Operanden und * mit
Integral/Boolean/Empty in beiden Reihenfolgen über vorhandene exakte64Bit-
Skalierung/int128-Kernel. Keine Double-Näherung für vollständige Currencywerte;
Range6 und ScalarDefault-einmal belegt. Microsoft-Plus-/Minus-Dokumentation
bestätigt PräzisionsrangCurrency überDouble/Single für+,-. Double-Produkte bleiben
Double; vollständige Single-Produktpromotion/Strings/Locale bleiben unbewiesen.
Decimal-/Date-/Null-/Access-Eval-Grenzen bleiben; keine erweiterte physische
MONEY-/rawSQL-/MDB-Parität. App/Core/DAO gebaut, finales bestehendes Bündel2/2
in8,64s (5,37/3,27) bestanden; tatsächliche native Funktion liefertCurrency2.0003
überSQL/DAO Type5/Size8/20003Einheiten. Keine neuen CTest-/Katalog-IDs;456/386/70
unverändert. Vorstufeb80c9521 privat remote/Baum geprüft. Letztes DEB4cd0f848 liegt
vor späteren Umfängen. Ohne Zwischenabschluss nach öffentlich dokumentierter
Byte-Variant-Erweiterung weitere Sprach-/Datenbankanforderungen fortsetzen.


Byte-Variant-Erweiterung: BYTE_VARIANT_WIDENING.md stützt fehlende Byte→Integer→
Long-Stufe auf tatsächliche öffentliche Microsoft-Plus-/Minus-/Multiplikations-
Dokumentation. Nur nachgewiesene Variant-Deklaration plus tatsächlicher Byte-
Subtyp, feste Byte-Überläufe bleiben6; ungebundene/Access-Eval-/physische SQL-
Grenzen unverändert. App/Core/DAO gebaut; erstes bestehendes Bündel2/2 in8,68s
(5,50/3,18) bestanden. Echter nativer Integer-VBA-Rückgabewert überSQL/DAO
Type3/Size2/Integer256 geprüft. Keine neuen CTest-/Katalog-IDs;456/386/70 unverändert.
Vorstufe04db5776 privat remote/Baum geprüft. Letztes geprüftes DEB4cd0f848 liegt vor
späteren Umfängen. Ohne Zwischenabschluss typisierte native Zahlenliterale mit
Suffices und weitere offene Anforderungen fortlaufend implementieren.


Native Zahlenliteral-Typzeichen: NATIVE_NUMERIC_SUFFIXES.md dokumentiert%Integer,
&Long,!Single,#Double,@Currency im expliziten VBA-Lexermodus. Currency direkt
Dezimaltext→exaktesFestkomma, Single direktfrom_chars→float ohneDouble-Rundung;
32Bit-Ranges, Exponenten, Präzedenz, Dates und bisherige angrenzende Konkatenation
bleiben. Metadaten-Klassentypescanner verwenden denselben nativen Lexermodus.
Keine Identifier-Typzeichen-/signedMinimum-Faltung-/Vollcompiler-/SQL-/Eval-
Behauptung. App/Core/DAO gebaut; erstes bestehendes Bündel2/2 in8,61s
(5,38/3,23) bestanden. Tatsächliche VBA-Funktion verwendet alle5Zeichen samt
OptionalLong=2& und liefertüberSQL/DAO Type5/Size8/exakt20002Currencyeinheiten.
Keine neuen CTest-/Katalog-IDs;456/386/70 unverändert. Vorstufe56f0e012 privat
remote/Baum geprüft. Letztes geprüftes DEB4cd0f848 liegt vor späteren Umfängen.
Ohne Zwischenabschluss weitere tatsächliche DAO-/Datenbankfunktionen fortsetzen.


DAO Recordset.CopyQueryDef: DAO_COPYQUERYDEF.md beschreibt echte unabhängige,
unbenannte QueryDef-Kopien aus parameterfreien QueryDef-Cursorn mit beim Öffnen
bewahrtem Access-SQL. Spätere Quelländerung/Close und Cursor-Clone sind belegt;
parameterisierte Quellen sowie Cursor ohne belegte QueryDef-Herkunft bleiben3251.
App/DAO gebaut, einmaliges bestehendes dao_execution-Bündel1/1 in3,23s bestanden,
einschließlich tatsächlichem typisiertem VBA-Aufruf und unabhängigen SQL-Zeilen.
Keine neuen Test-/Katalog-IDs,456/386/70; F054/F055 bleiben offen. Vorstufe
df22cf9b privat remote/Baum geprüft. Kein Gesamtabschluss/Abschlussanruf.


DAO Filter/Sort/abgeleitete Snapshots: DAO_DERIVED_SNAPSHOTS.md dokumentiert echte
SQL-Filterung/Sortierung bestehender querybasierter Cursor über explizites
OpenRecordset(dbOpenSnapshot). Ausgangscursor unverändert, neue Ergebnisse und
Klon-Eigenschaften unabhängig; Parameter und Currency-Metadaten bleiben erhalten.
Editierbare/implizite Modi, Tabellen-/Frozenadapter und volle Oracleidentität offen.
App gebaut; bestehendes DAO-Bündel1/1 in3,37s bestanden nach Korrektur eines
Prüfcode-Feldnamens beim Übersetzen. Keine neuen CTest-/Katalog-IDs;456/386/70.
Vorstufe ad42eadace27a518fbbaf1f7a9819facc1e383c1 privat remote/Baum geprüft:
neues DEB geprüft, bestehende Integration3/3 in188,36s (GUI19,85/Inventar47,59/
Original120,91), sieben geschützte Hashes unverändert. Inventar100/78/22;
Original-Makrogrenze bleibt. Paket liegt vor abgeleiteten Snapshots. Kein
Gesamtabschluss/Abschlussanruf; weitere allgemeine Anforderungen fortsetzen.


Native Ordner-Metadaten: NATIVE_FILE_SYSTEM_OBJECT.md beschreibt Folder.Size
über tatsächliche verschachtelte/versteckte Dateien mit exakter Double-Ganzzahl-
Grenze, Besuchs-/Tiefenlimits und ausdrücklichen Link-/Spezialdatei-Grenzen sowie
Folder.DateLastModified über vorhandene lokale Date-Konvertierung. Kein Windows-
Variant-/atomarer Snapshot-/FSO-Vollparitätsanspruch. App/Core gebaut, einmaliges
bestehendes Core-Bündel1/1 in5,56s bestanden, reale kleine und sparse>2GiB-Dateien,
Zeitstempel, Link-/Löschgrenzen und typisierter VBA-Aufruf.456/386/70 unverändert,
keine neuen Test-/Katalog-IDs. Vorstufe45a6f86a privat remote/Baum geprüft.
Aktuelles geprüftes DEB bleibt ad42eada vor abgeleiteten Snapshots/Ordner-Metadaten.
Ohne Zwischenabschluss weitere allgemeine offene Anforderungen fortsetzen.


Native DateCreated/DateLastAccessed für File und Folder: vorhandene
NATIVE_FILE_SYSTEM_OBJECT.md dokumentiert tatsächliche Linux-statx-Zeitstempel,
Verfügbarkeits-/Objektartprüfung und3251 bei fehlender Erstellzeit; keine
Änderungs-/Inode-/Jetzt-Ersatzwerte. Vorhandene lokale Date-/Range-Konvertierung
nimmt zusätzlich geprüfte Unixsekunden an. App/Core gebaut; einmaliges bestehendes
Core-Bündel1/1 in5,45s bestanden: echte getrennte Zugriffs-/Änderungszeiten,
typisiertes VBA, tatsächliche/ausdrücklich fehlende Erstellzeit und Löschgrenzen.
Keine neuen Test-/Katalog-IDs;456/386/70 unverändert. Vorstufe1da6ad45 privat
remote/Baum geprüft. Geprüftes DEB weiterhin ad42eada vor späteren Umfängen.
Ohne Zwischenabschluss weitere allgemeine offene Anforderungen umsetzen.


Quelllokale qualifizierte Enum-As-Typen: VBA_ENUM_TYPE_BINDINGS.md dokumentiert
Modulname.Enumname nur aus tatsächlichem Attribute VB_Name derselben Quelle,
inklusive Const-/Feld-/Parameter-/Rückgabe-/Array-/Inline-Bindungen alsLong.
Keine erfundene Modulidentität oder projektweite Sichtbarkeit; fremde Qualifier
bleiben abgewiesen. Quellreine Parsercache-Identität bleibt unverändert. App/Core/
GUI gebaut; einmaliges bestehendes Bündel2/2 in25,29s (5,51/19,78), unveränderter
präziser20s-GUI-Watchdog. Echte SQL-Funktion mit qualifiziertem Enum-Resultat
registriert Long und liefert1234. Keine neuen Test-/Katalog-IDs;456/386/70.
Vorstufef4e44eb2 privat remote/Baum geprüft; DEB weiterhinad42eada vor späteren
Umfängen. Projektweite Enum-Bindung benötigt sichtbarkeitsabhängiges Typumfeld
und dessen Cache-/Modulzustandsidentität; keine Klassenheuristik einbauen.
Ohne Zwischenabschluss weitere echte offene Anforderungen fortsetzen.


Native MonthName/WeekdayName: CALENDAR_FUNCTIONS_RUNTIME.md dokumentiert echte
lokalisierte String-Namen über aktuellen QLocale-/Standalone-Kalenderadapter,
WeekdayName-Systemstart0 bzw. explizite VBA-Sonntag1..Samstag7, Abkürzungen,
benannte/ausgelassene/qualifizierte Aufrufe und kontrollierte Range-/Null-Grenzen.
Formular-, Klassen-, Standard-/Importeditor-, direkter SQL- und quellbasierter
SQL-Modulhost integriert; kein zusätzlicher SQL-UI-/DAO-Zugriff. SQL-Funktionen
wegen Sprachabhängigkeit nicht als deterministisch registriert. WindowsNLS/
Hijri-/vollständige Metadaten-/Oracle-Parität bleibt offen. App/Core/GUI/Format
gebaut; einmaliges bestehendes Bündel3/3 in24,21s (5,44/18,72/0,05),20s-Watchdog
unverändert. Reale deutsche/englische SQL-Namen und echte Formular-/Modul-/SQL-
Modulfunktionen geprüft. Keine neuen CTest-/Katalog-IDs;456/386/70. Vorstufe
d339d4b2 privat remote/Baum geprüft. Letztes geprüftes DEB weiterhinad42eada vor
den späteren Umfängen; anschließend gebündelt Paket/Originalbelege erneuern und
weitere offene Anforderungen ohne Zwischenabschluss umsetzen.


DAO SourceField/SourceTable: DAO_FIELD_ORIGINS.md dokumentiert echte physische
main-Schema-Herkunft aus vorbereiteten Statements für Recordset/QueryDef/TableDef,
Aliasse/Views/Joins/Snapshots, schemareine QueryDef-Metadaten ohne Parameterausführung
und erweiterte Herkunftsidentität. Keine erfundene Label-/UNION-/TEMP-/berechnete
Herkunft. Tabellen-Recordset.SourceTable ist ausdrücklich unbelegt3251, Index- und
Relationskontexte bleiben geschützt. App/DAO/Query/Currency gebaut; einmaliges
bestehendes Bündel5/5 in3,90s (0,09/0,06/0,16/3,41/0,18) bestanden. Kein neuer
CTest-/Katalogpunkt;456/386/70 unverändert. Vorstufe4ebc3f6c privat remote/Baum
geprüft; neues DEB/Paketprüfung und Originalbündel2/2 in167,38s (47,31/120,07),
sieben geschützte Hashes unverändert. Inventar100/78/22, fehlende Makros weiterhin
Grenze. Kein erneuter bereits bestandener GUI-Lauf; DEB liegt vor Feldherkunft.
Nächster belegter DAO-Fehler: AddNew/Update soll einen zuvor gültigen aktuellen
Datensatz erhalten; neuer Datensatz über LastModified. Bisher springt update()
pauschal zur neuen Identität. Erhaltenen aktuellen Datensatz und LastModified
getrennt behandeln; leere/BOF/EOF-/entfernte Vorposition nicht ungeprüft als
Windows-konform behaupten. Danach weitere offene Anforderungen ohne Abschluss.


DAO AddNew/Update aktueller Datensatz: DAO_ADDNEW_CURRENT.md dokumentiert Erhalt
eines zuvor gültigen weiterhin vorhandenen aktuellen Datensatzes und davon
unabhängiges LastModified des eingefügten Datensatzes. Tabellen/Dynaset, Klone,
Bookmarkidentität, Cancel/Fehler/Retry und echtes VBA belegt. Erstes DAO-Bündel
legte QSQLITE-null-lastInsertId bei gültiger physischer Zeilenkennung0 offen;
gezielt durch tatsächliche SQLite-Verbindungskennung ohne Double/zusätzlichesSQL
behoben. Nur betroffenen DAO-Lauf wiederholt:1/1 in3,27s bestanden, App gebaut.
Leere/BOF/EOF/verschwundene Voridentität, Navigation mit Entwurf, dynamische
Anfügereihenfolge und Windows-Vollorakel bleiben offen. Keine neuen CTest- oder
Katalog-IDs;456/386/70. Vorstufec77bdf6a privat remote/Baum geprüft. Geprüftes DEB
weiterhin4ebc3f6c vor Feldherkunft/AddNew. Ohne Zwischenabschluss weitere echte
DAO-Zustands-/Sprach-/Datenbankanforderungen implementieren; keine AddNew-ADO-
FieldList/Values-Signatur in DAO erfinden (DAO-AddNew ist parameterlos).


DAO Entwürfe bei tatsächlicher Navigation: DAO_PENDING_NAVIGATION.md dokumentiert
Verwerfen von Edit-/AddNew-Wert-/Chunkpuffern erst nach erfolgreicher Zielprüfung
für tatsächliche Positionswechsel in Move/Positions-/Bookmark-/Find-/Seek-Pfaden.
Keine SQL-Speicherung, Klonentwürfe getrennt; Reentranz-/Versionsprüfungen erhalten.
Ungültige Ziele/Suchkriterien erhalten Entwürfe. No-op-Navigation bleibt3219;
Clone/Index/Requery/GetRows und vollständige Windows-Fehlerpriorität weiter offen.
Formular-UI-Speicher-/Cancelereignisse unverändert. App/DAO gebaut; einmaliges
bestehendes DAO-Bündel1/1 in3,67s bestanden, echte Tabellen/Dynaset Edit-/AddNew-
Bewegungsmatrix und unabhängiger SQL-Nichtschreibbeleg. Keine neuen CTest- oder
Katalog-IDs;456/386/70. Vorstufec0dc87a9 privat remote/Baum geprüft. Letztes
geprüftes DEB4ebc3f6c liegt vor Feldherkunft/DAO-Zustandsänderungen. Kein neuer
GUI-/Original-/Paketbeleg für diese Etappe. Ohne Zwischenabschluss weitere
allgemeine offene Anforderungen fortlaufend implementieren.


Native Dynaset-Reihenfolge: DAO_DYNASET_ORDER.md beschreibt linearen Merge nach
exakter physischer Identität bei nichtinvalidierendem Refresh nativer Tabellen-
Dynasets: vorhandene Zeilenordnung behalten, neue Zeilen anhängen, anschließend
vorhandene Bookmark-/Klon-/Cursorwiederherstellung. Nachfolgende Updates setzen
Reihenfolge nicht auf SQLite-Physicalorder zurück. Tabellen-Cursor/Indexpfade und
explizites invalidierendes Requery bleiben getrennt; keine neue SQL-Ausführung.
App/DAO gebaut, einmaliges bestehendes DAO-Bündel1/1 in3,24s bestanden: reale
Ordnung mit neuem Schlüssel0 (Tabelle0/1/2/3 vs.Dynaset1/2/3/0), Klone, Updates,
weitere Inserts und unveränderte Positions-/LastModified-/Entwurfsmatrizen.
Keine neuen CTest-/Katalog-IDs;456/386/70. Vollständige externe Live-Dynaset-/
Delete-/Orakelparität bleibt offen. Vorstufe32d4a49d privat remote/Baum geprüft.
Geprüftes DEB4ebc3f6c liegt vor Feldherkunft/DAO-Zustands-/Ordnungsänderungen;
anschließend gesammelte Änderungen gebündelt paketieren/Original prüfen und
weitere Anforderungen ohne Zwischenabschluss fortsetzen.


Physische DAO-Quellbindung: DAO_PHYSICAL_SOURCE_SCOPE.md dokumentiert tatsächliche
main-Qualifikation der Tabellen-Lese-/INSERT-/UPDATE-/DELETE-Pfade und explizite
Prüfung auf TEMP-Namensüberlagerung vor Laden/Reload und Mutationsbeginn. Verhindert
Mischen echter main-Metadaten mit anderen TEMP-Zeilen/-Zieltabellen. Verweigerung3251
vor SQL erhält Entwurf; nach Entfernen der Kollision wird beabsichtigtes main-
Update möglich. Cached main-Herkunft bleibt tatsächlich zu cached main-Werten.
App/DAO gebaut; einmaliges bestehendes DAO-Bündel1/1 in3,42s (gesamt3,43) bestanden,
unabhängige SQL-Belege für beide unveränderten Tabellen und TEMP-only-Grenze.
Keine neuen CTest-/Katalog-IDs;456/386/70. Vorstufea4bbf4a4 privat remote/Baum geprüft:
neues DEB/Paketprüfung und gebündelter Originaltest1/1 in118,03s, sieben geschützte
Hashes unverändert, Makrogrenze bleibt. Bereits bestandene GUI-/Inventarläufe nicht
wiederholt; tatsächlicher letzter Inventarbeleg4ebc3f6c100/78/22. Dieses DEB liegt
vor physischer Quellbindung. Danach weitere allgemeine offene Anforderungen ohne
Zwischenabschluss fortsetzen; keine umfassende roheSQL-/Provider-/Crossconnection-
Schattierungs-/Generationsparität behaupten.


DAO Restartable/Origin-Properties: DAO_RESTARTABLE_PROPERTIES.md dokumentiert
Boolean-Capability für unterstützte Query/Dynaset/Snapshot-Cursor, dokumentiertes
False für Tabellen-Cursor sowie False für eigenständig nicht-requeryfähige Frozen-
Adapter. Stale Query nach Rollback kannFlag lesen und anschließend echtRequery
aufrufen; native bisherige Tabellen-Requery-Erweiterung bleibt getrennt. Native
SourceField/SourceTable über echte schreibgeschützte Properties-Objekte/dbText10,
Quelle/Getterschutz unverändert. App/DAO gebaut; einmaliges DAO-Bündel1/1 in3,37s
bestanden inklusive echtem VBA und tatsächlicher Rollback-Wiederabfrage. Keine
neuen CTest-/Katalog-IDs;456/386/70. Vorstufe718ca355 privat remote/Baum geprüft;
letztes DEBa4bbf4a4 vor Quellbindung/Metadaten. Moderne veröffentlichte DAO-PIA
15.0.4420.1018 nur als Vertrag inspiziert: GetRows.NumRows istoptional OHNE
Metadaten-Standardwert, AddNew parameterlos. Keine Default1-Behauptung aus
ADO/Optionalität ableiten; keine PIA inGit oder Original-Ausführung.
Ohne Zwischenabschluss weitere echte DAO-/Datenbankanforderungen implementieren.


DAO gelöschte Cursorposition: DAO_DELETED_POSITION.md dokumentiert einen echten
logischen Löschmarker mit Nachbaridentitäten statt Sprung zum ersten Datensatz.
Betroffene Klone, ungültige Bookmarks, Next/Previous/Move/FindNext und BOF-/EOF-
Zustände nach letzter Löschung sowie invalidierendes Requery werden berücksichtigt.
Pending Edit nach Klonlöschung prüft Current vor Identitätsindex/SQL und kann
keine Nachbarzeile überschreiben; Reload-Fehlerrestore umfasst neue Zustände.
Native3021 bleibt; vollständige3167-/Eigenschafts-/Provider-/Concurrency-Identität
unbelegt. App/DAO gebaut; einmaliges bestehendes DAO-Bündel1/1 in3,40s bestanden,
reale Tabellen/Dynaset-/Klon-/Lösch-/Entwurfsfälle und echtes VBA Delete/MoveNext
löscht alle3Zeilen, unabhängigSQLCOUNT0. Keine neuen CTest-/Katalog-IDs;456/386/70.
Vorstufe063402b9 privat remote/Baum geprüft. Letztes geprüftes DEBa4bbf4a4 liegt
vor Quellbindung/Restartable/Löschzustand; anschließend diese gesammelten Änderungen
gebündelt paketieren/Original prüfen und weitere offene Anforderungen fortsetzen.
Spezifische BOF-/EOF-/Move-Dokumentation verwenden; moderne generierte Delete-
Seiten enthalten widersprüchliches unsupported, keine Blindübernahme als Orakel.


DAO schreibbare Sternabfragen: DAO_EDITABLE_STAR_QUERIES.md dokumentiert echte
Mode2-Edit/Update/Delete auf streng geprüften schlüsselgebundenen einzelnen main-
Tabellen mit vollständigem */t.*. WHERE wird beim Laden und Schreiben angewandt;
ORDER BY nutzt physische Felder. Klone, Herkunft und abgeleitete Snapshots bleiben
nutzbar. Ungefiltertes AddNew unterstützt; gefiltertes AddNew explizit3251 vor
Entwurf/Defaults/Reservierung. Projektionen/Parameter/komplexe/keylose Quellen und
Mode4 werden nicht schreibbar erklärt. App/DAO/QueryModel/Currency gebaut; einmaliges
bestehendes Bündel5/5 in3,88s bestanden, DAO3,34s; echte SQL-/VBA-/Klon-/Stale-
Predicate-/Retry-/Lösch-/Insert-Belege. Keine neuen CTest-/Katalog-IDs;456/386/70.
Vorstufec67cbd9e privat remote/Baum geprüft, neues DEB geprüft und Original1/1
in120,41s bestanden, sieben geschützte Hashes unverändert; Makrogrenze bleibt.
Dieses Paket liegt vor Sternabfragen. Inventar unverändert letzter tatsächlicher
Lauf4ebc3f6c100/78/22. Keine Windows-Orakel-/vollständige Dynaset-Paritätsbehauptung.
Weitere allgemeine Anforderungen ohne Zwischenabschluss umsetzen.


DAO Recordset.Name: DAO_RECORDSET_NAME.md dokumentiert getrennte echte Öffnungs-
identität für Tabelle/gespeicherte Abfrage/QueryDef und erste256Zeichen des
ursprünglichen Access-SQL vor Übersetzung/physischem Schreibziel. Klone behalten
Identität; schreibgeschützt383, Argument450, geschlossen3420. Frozen/transformierte
Adapter ohne bewiesene Namenssemantik explizit3251. App/DAO gebaut; einmaliges
bestehendes DAO-Bündel1/1 in3,41s bestanden. Keine neuen CTest-/Katalog-IDs;
456/386/70. Vorstufe8de1ead privat remote/Baum geprüft; letztes geprüftes Paket/
Original weiterhinc67cbd9e vor Abfrageschreiben/Namen. Ohne Zwischenabschluss
weitere allgemeine Anforderungen implementieren; keine Windows-Orakelbehauptung.


DAO physische Abfrageprojektionen: DAO_EDITABLE_QUERY_PROJECTIONS.md dokumentiert
schreibbare eindeutige physische Spaltenlisten/Teilmengen/Umordnung ohne Alias.
Sichtbare Reihenfolge/Herkunft/Typen bleiben echt, unsichtbarer Schlüssel wird
getrennt für Mutation geführt. Nur ausgewählte Defaults gelangen in Entwürfe;
ausgelassene Text-/Date-Defaults bleiben echter DB-Ausführung. Ungefiltertes
AddNew unterstützt; fehlende Generatorprojektion vor Entwurf/Reservierung3251.
App/DAO/QueryModel/Currency gebaut; einmaliges Bündel5/5 in3,98s bestanden.
Review ergänzte Generatorgrenze, nur betroffenes DAO erneut1/1 in3,43s.
Echte SQL-/VBA-/Klon-/Metadaten-/Subset-/Umordnungsbelege. Keine neuen CTest-/
Katalog-IDs;456/386/70. Vorstufe93a8a9f privat remote/Baum geprüft; letztes geprüftes
DEB/Original weiterhinc67cbd9e. Alias-/Join-/Provider-/Orakelparität offen.
Ohne Zwischenabschluss weitere allgemeine Anforderungen implementieren.


Formular-Quell-/Insertidentität: FORM_PHYSICAL_INSERT_IDENTITY.md dokumentiert
tatsächliche main-Qualifikation beim Laden/INSERT/UPDATE/DELETE und TEMP-/Main-
Quellprüfung nach Benutzerereignissen sowie vor Schreibausführung. INSERT nutzt
wie UPDATE echte einzelne RETURNING-rowid einschließlich0; DataEntry/Selektion/
Snapshot/Folgeänderung behalten Identität. Abgewiesene Entwürfe bleiben erhalten.
App/FormCurrent/Layout/Delete/Currency-Insert/Update gebaut. Fixture-Compilezugriff
auf private Methoden durch echte öffentliche VBA-Routen ersetzt, davor keine
Prüfung ausgeführt; einmaliges bestehendes Bündel5/5 in14,42s bestanden.
Echte Tastatur-/VBA-/SQL-/TEMP-Refusal-/Retry-Belege; keine neuen CTest-/Katalog-IDs;
456/386/70. Vorstufe65fa8d1 privat remote/Baum geprüft. Abfrageformulare/Provider/
Quellgenerations-/vollständige Orakelparität offen. Letztes geprüftes DEB/Original
weiterhinc67cbd9e vor späteren Änderungen. Ohne Zwischenabschluss weiterarbeiten.


DAO-Aliasse/Record-Clipboard: DAO_EDITABLE_QUERY_ALIASES.md und
RECORD_CLIPBOARD_INSERT_IDENTITY.md dokumentieren getrennte sichtbare/physische
Feldzuordnung für Schreiben, Generatoren, Defaults/Typen/Herkunft/Properties und
Klone. Gemeinsamer Validator prüft WHERE physisch bei abgeschaltetem DQS-Fallback,
ohne Ausführung; mehrdeutige Qt-Fold-Ausgaben nicht schreibbar. Clipboard nutzt
main-Bindung/Quellprüfung und echte RETURNING-Identität auch0. Gemeinsames Bündel:
5/6 bestanden; DAO-Fixtureannahmen zu unbewiesenen Original-Properties und noch
nicht unterstütztem Currency-CreateField korrigiert, ausschließlich DAO wiederholt,
zuletzt1/1 in3,61s. Modell/Currency/Layout13,78s nicht erneut ausgeführt.
Echte VBA-/SQL-/Counter-/Single-/Date-/BLOB-/Property-/Clipboard-Belege; exakte
Currency-Aliasse an unterstützter physischer Tick-Tabelle. Keine neue Currency-
CreateField-/Orakel-/vollständige Query-Paritätsbehauptung. Keine neuen CTest-/
Katalog-IDs;456/386/70. Vorstufe73328ff privat remote/Baum geprüft. Letztes DEB/
Original weiterhinc67cbd9e; nach Sicherung gesammelten Stand paketieren/Original
prüfen und ohne Zwischenabschluss weitere allgemeine Anforderungen implementieren.


DAO Feld-Schreibfähigkeiten: DAO_FIELD_WRITE_CAPABILITIES.md dokumentiert echtes
Boolean DataUpdatable aus Cursor-/Verbindungsschreibfähigkeit für alle aufgelösten
Felder statt Counter-only-Guard. Readonly-/Snapshot-/berechnete ErgebnisseFalse;
native schreibbare Tabellen/AliasfelderTrue. Native Attributes nutzen validierte
Flags1/2/17 plus32 nur bei Schreibfähigkeit; Original-/Providerflags nicht erfunden.
Property-Objekte, Argument450/Readonly383/Closed3420 und Counter-Edit3164 bleiben.
App/DAO gebaut; einmaliges bestehendes DAO-Bündel1/1 in3,69s bestanden,
echte SQL-/VBA-/native Feld-/Property-/Snapshot-/Readonly-Umschaltungsbelege.
Keine neuen CTest-/Katalog-IDs;456/386/70. Vorstufea23c7009 privat remote/Baum geprüft.
Neues DEBa23c7009/Paketprüfung bestanden, Original1/1 in119,54s bestanden;
alle sieben geschützten Hashes unverändert, bekannte Makrogrenze bleibt.
Dieses Paket liegt vor Feldfähigkeiten; letzter Inventarbeleg4ebc3f6c100/78/22;
keine vollständige Windows-/Provider-/Fehlerreihenfolgenparität behaupten.
Ohne Zwischenabschluss weitere allgemeine Anforderungen implementieren.


Dictionary Container-Objektschlüssel: SCRIPTING_DICTIONARY_RUNTIME.md dokumentiert
echte Collection-/Dictionary-Referenzidentität als Schlüssel, gehaltene Original-
referenzen, Keys-/ForEach-Objekte, Lookup/Item/Remove/Umbenennung ohne Scalar-Default.
Besitzgraph umfasst Schlüssel und Werte; Add/Key-Rename prüfen vor Mutation,
direkte/über Schlüssel und Werte/Collection verlaufende Zyklen verweigern3251.
Nothing/Arrays/beliebige Host-/Projektklassen-/COM-Schlüssel weiter explizit begrenzt.
App/Core gebaut; vor Prüfungen versehentliche gemeinsame Textänderung an Collection
auf Dictionary eingeschränkt und Fixture-API-Aufrufe korrigiert; einmaliges
bestehendes Core-Bündel1/1 in5,70s bestanden. Echte VBA-/Objekt-/Weak-
Referenz-/Snapshot-/Zyklusbelege. Keine neuen CTest-/Katalog-IDs;456/386/70.
Vorstufec59e828 privat remote/Baum geprüft; letztes geprüftes DEBa23c7009, Original
1/1 in119,54s und sieben geschützte Hashes unverändert. Dieses Paket liegt vor
Feldfähigkeiten/Objektschlüsseln. Kein vollständiger COM-/Windows-/Lifetimebeleg;
ohne Zwischenabschluss weitere allgemeine Anforderungen implementieren.


Berichtsgruppen: REPORT_GROUP_METADATA.md dokumentiert echte42 aktive BreakLevel-
Blöcke in neun Originalberichten sowie30/29 BreakHeader/Footer. Ausgabegrenze
übersah bislang diese Access2003-Struktur; aktive Quellen/Break-Abschnitte werden
nun ausdrücklich vor vereinfachtem Druck/PDF abgefangen, Ziel-PDF bleibt erhalten.
Quellparser bewahrt bis10 tatsächliche Gruppen und explizite Eigenschaften;
Report.GroupLevel(Index) liefert lebensdauergebundene gecachte native Objekte mit
ControlSource String, SortOrder/Header/Footer Boolean, GroupOn Integer, Interval
Long und KeepTogether Byte. Qualifizierte As-Bindung/Funktionsrückgabe/named Index
funktionieren; fehlende Defaults3251, Mutation383, Arity450, Null94, Index5.
App/Core/Layout gebaut; frühe Ausgabegrenzenprüfung1/1 in13,67s, anschließend
kombiniertes bestehendes Core/Layout-Bündel2/2 in18,77s bestanden. Echte VBA- /
Objektidentitäts-/Metadaten-/Closed-/PDF-Erhalt-Belege; Originalkorpus bindet seine
neun tatsächlichen Quellen ohne neue private VBA-Ausführung. Keine neuen CTest-/
Katalog-IDs;456/386/70. Gruppendruck/-sortierung/-intervalle/-events/-Mutationen und
Windows-Parität bleiben offen. Vorstufe29ed6fd5 privat remote/Baum geprüft;
letztes DEBa23c7009/Original119,54s/sieben Hashes unverändert liegt vor späteren
Feld-/Container-/Gruppenänderungen. Ohne Zwischenabschluss weiterentwickeln.


Build-Abhängigkeit nativer Typnamen: Die unveränderten drei Lookup-Funktionskörper
stehen jetzt in native_object_types.cpp statt inline im breit eingebundenen Header;
CMake bindet diese Datei in dieselbe Laufzeitbibliothek ein. Künftige Vokabular-
änderungen betreffen eine Übersetzungseinheit. Funktionskörper mechanisch gegen
Vorstufe verglichen, App/Core gebaut, bestehendes Core-Bündel1/1 in5,57s bestanden.
Keine neue Laufzeitfunktion/CTest-/Katalog-ID und keine Laufzeit-Benchmarkbehauptung;
456/386/70. Vorstufee8f54673 privat remote/Baum geprüft. Weiterentwickeln ohne
Zwischenabschluss; letztes geprüftes DEB/Original weiterhina23c7009 vor Folgeänderungen.


Native DAO-Currency-Erzeugung: DAO_NATIVE_CURRENCY_CREATION.md dokumentiert echtes
CreateField(...,dbCurrency) für neue Tabellen und Fields.Append an bestehenden.
CURRENCY_REAL nutzt den bestehenden verlustfreien Vierstellenpfad streng innerhalb
+/-2^39, physische CHECK erzwingt Normalisierung; MetadatenType5/Size8/Required und
Default-Drafts/SQL werden tatsächlich genutzt. Defaults Null oder Dezimalliterale
mit höchstens vier Stellen exakt konvertiert, überschüssige Präzision/Exponent/
Aufrufe/Bereich vor Default-Ersatz3251. DefaultValue zeigt normalisierten Currency-
Text statt binärer SQL-Serialisierungsziffern. Fehlendes Fixture-Include vor Checks
korrigiert; App/DAO/Modell/Currency gebaut. Bestehendes Bündel4/5 bestanden, DAO
fand echten physischen Typkonflikt: Speicherprotokoll verlangt REAL. Deklaration
auf REAL korrigiert, logischer Typ CURRENCY_REAL aus Metadaten; nur DAO erneut1/1 in3,50s
bestanden, echte SQL-/Append-/Constraint-/Metadaten-/VBA-Erzeugungsbelege. Keine
neuen CTest-/Katalog-IDs;456/386/70. Vollständiger Jet-Tickbereich/Decimal/Currency-
Rebuild/Validation/Provider-/Windows-Parität bleibt offen. Vorstufe2af79959 privat
remote/Baum geprüft; letztes DEB/Originala23c7009 vor späteren Änderungen.
Ohne Zwischenabschluss weitere allgemeine Anforderungen implementieren.


Native Currency-Schemaerhaltung: DAO_NATIVE_CURRENCY_SCHEMA.md dokumentiert
kanonische Wiederherstellung tatsächlicher Currency-Defaults für bekannte native
Rekonstruktionen. Default/Required, Feldreihenfolge, Index-Append/Delete und Tabellen-
Umbenennung erhalten vorhandene Werte/rowids; bestehende Äquivalenz-/Digest-/
Savepoint-/Rollback-/Metadatenprüfungen unverändert. Feldumbenennung ebenfalls
belegt. App/DAO gebaut; erste DAO-Fixture erwartete falschen Fehler für angehängtes
Field.Type (korrekt383); tatsächlicher Designerpfad separat3251 geprüft. Nur DAO
wiederholt1/1 in3,71s bestanden, unabhängige SQL-Bitvergleiche/NULL-Rollback/
neue Defaults nach Umbenennungen. Keine neuen CTest-/Katalog-IDs;456/386/70.
Vorstufe90f672b9 privat remote/Baum geprüft. Currency-Typkonversion/Seek/FK-Schlüssel/
voller Tickbereich/Originalschema-/Windows-Parität bleiben offen. Letztes DEB/
Originala23c7009 vor Folgeänderungen; ohne Zwischenabschluss weiterentwickeln.


Currency-Feldeditor/-Typwechsel: CURRENCY_FIELD_DESIGNER.md dokumentiert echte
Currency-Auswahl/Anzeige/Anlage und geprüfte numerische Typwechsel. Vorhandene
Werte und Defaults müssen Ziel-Speicher bitgenau numerisch durchlaufen, Currency-
Defaults erst danach kanonisch normalisiert. Rundung/Bereich/Ziel-Currency mit
Regeln/Meldungen verweigert; DAO Field.Type weiterhin readonly383. GUI zeigt
unterstützten Betragsbereich, Currency-Regelbearbeitung bleibt begrenzt. App/DAO/
GUI gebaut; einmaliges bestehendes Bündel2/2 in22,84s bestanden (GUI19,25s bei
unverändert20s-Watchdog, DAO3,59s). Echte GUI-Anlage/Double-Rückwechsel/Default/
SQL-Werte/Löschen und Backend-Matrix mit NULL/Integralen/Präzisions-/Bereichs- /
Regelrefusal belegt; keine neuen CTest-/Katalog-IDs;456/386/70. Vorstufe6786526d
privat remote/Baum geprüft; letzter Paket-/Originalbelega23c7009 vor Folgeänderungen.
Nach Sicherung gesammelten Stand paketieren/Original/Inventar gebündelt prüfen;
keine vollständige Jet-/Currency-/Windows-/Distributions-Parität behaupten.
Ohne Zwischenabschluss weitere allgemeine Anforderungen umsetzen.


Currency im anfänglichen Tabellenassistenten: CURRENCY_TABLE_WIZARD.md dokumentiert
Auswahl/Typ5-Mapping und numerische Vorprüfung, anschließend exakte bereits
implementierte DAO-Default-Prüfung vor TableDefs.Append. Abgewiesene Präzision
hinterlässt keine Teiltabelle und erhält Dialogeingaben; korrigierter Vierstellen-
Default wird tatsächlich gespeichert. App/GUI gebaut; erster GUI-Lauf überschritt20s; erster Lookup-Retry ebenfalls. Phasenmessung
zeigte nur26ms für den Assistenten. Unveränderliche VBA-Syntaxmuster werden nun
einmalig kompiliert. Bestehendes Core/GUI/DAO-Bündel3/3 in15,18s bestanden, GUI in12,29s bei unverändert20s-Watchdog bestanden; echte Dialog-/Metadata- /
SQL-/Currency-Wert-Belege, keine neuen CTest-/Katalog-IDs;456/386/70.
Vorstufeda72e727 privat remote/Baum geprüft: neues DEB/Paketprüfung bestanden,
Original1/1 in120,33s und Inventar1/1 in46,11s, sieben geschützte Hashes unverändert.
Aktueller Inventarbeleg jetztda72e727100/78/22 (73+5 Ausführungen ohne Access-Orakel,
3 Originalspalten/9 fehlende Abhängigkeiten/10 Jet-/Bindinggrenzen). Makrogrenze
bleibt, kein positiver initialisierter Verkauf/Druck. Dieses Paket liegt vor
Currency-Auswahl im anfänglichen Assistenten. Ohne Zwischenabschluss fortsetzen;
keine vollständige Currency-/Jet-/Windows-/Distributions-Parität behaupten.


VBA-Syntaxmuster-Engpass: VBA_IMMUTABLE_SYNTAX_PATTERNS.md dokumentiert62 feste
Regexmuster, deren unveränderte Bytes/Flags jetzt einmalig initialisiert werden.
Keine Modulquellen/Symbolumgebungen/Argumente/Host-/Variablen-/Ergebniszustände neu
gecacht; Match-Ergebnisse bleiben lokal. Primitive Typen überspringen Objekttabellen,
Scripting-Lookup fremde Präfixe. Optionaler GUI-Phasenlogger enthält nur Namen/Zeit.
20s-Watchdog unverändert, keine vorhandenen Prüfungen entfernt; finales bestehendes
Core/GUI/DAO-Bündel3/3 in15,18s (0,55/12,29/2,35) bestanden. Tatsächliche Etappenzeiten,
keine allgemeine Benchmark-/Paritätsquote; keine neuen CTest-/Katalog-IDs;456/386/70.
Ohne Zwischenabschluss weitere allgemeine Anforderungen fortsetzen.


Native Währungs-Beziehungsschlüssel: DAO_NATIVE_CURRENCY_RELATIONS.md. Typ5 ist
für geprüfte native REAL/CURRENCY_REAL-Paare mit expliziten Felddeklarationen
freigeschaltet; bestehende Required-/Unique-/Schema-/Wertdigest-/FK-Prüfungen
bleiben. Verwaiste Datensätze brechen atomar ab, Retry nach Korrektur funktioniert;
echte Update-/Delete-Kaskaden, Null-Fremdschlüssel, Binärpayload, Reorder und
Relation-Delete im bestehenden DAO-Bündel1/1 in2,44s bestanden. App/DAO gebaut.
Keine volle Jet-Currency-Reichweite/Windows-Parität; F014/F055/F121 bleiben offen,
keine neue CTest-/Katalog-ID;456/386/70 unverändert. Vorstufe33b97df privat und
Remote-Head überprüft. Letztes DEB/Original/Inventar bleibt da72e727; dessen
bekannte Makrogrenze ist kein positiver initialisierter Verkaufs-/Drucknachweis.
Unabhängige schreibgeschützte MDB-Kopiebeobachtung:29 Tabellen/316 Spalten,
81 tatsächliche Defaults vollständig übernommen ohne Literalabweichung;
15 zusätzliche Backenddefaults ausschließlich Boolean. Keine VBA-/Geschäftsdaten
für diese Beobachtung geladen oder ausgegeben. Ohne Zwischenabschluss fortsetzen.


Native Währungs-Gültigkeitsregeln: NATIVE_CURRENCY_VALIDATION.md. Erstellung,
Feldmetadaten, tatsächlicher Feldeditor und explizite verlustfreie Typwechsel
unterstützen Vergleiche/And/Or/Between mit exakt darstellbaren Dezimalgrenzen bis
vier Nachkommastellen im vorhandenen nativen Bereich. Keine stille Rundung von
Regelgrenzen; Exponenten/fünfte Stelle/Funktionen/Bereichsüberschreitung werden
vor Änderung abgewiesen. SQL-Constraints mit ValidationText, atomarer Abbruch
bei Bestandsverletzung und kompatible Currency/Double-Regelerhaltung geprüft.
GUI12,45s bestanden (20s-Watchdog unverändert). Erstes DAO-Fixture benutzte nach
Spaltenaddition eine erwartungsgemäß ungültige alte Feldreferenz; Fixture greift
nun neu auf das tatsächliche Feld zu. Nur DAO neu gebaut/geprüft:1/1 in2,44s.
App/GUI/DAO gebaut; keine neuen CTest-/Katalog-IDs;456/386/70. Frühere pauschale
Currency-Regelsperre ist für diesen begrenzten Umfang ersetzt; keine volle
Jet-/Windows-Ausdrucks-/Currency-Parität. Vorstufeb8fee632 privat/Baum geprüft;
sieben geschützte Hashes unverändert. DEB/Original/Inventar weiterhinda72e727.
Ohne Zwischenabschluss bei weiteren allgemeinen Anforderungen fortsetzen.


Native Currency Index/Seek und Beziehungsoberfläche:
NATIVE_CURRENCY_INDEX_AND_RELATION_UI.md. Gemeinsame Prüfung tatsächlicher
nativer REAL/CURRENCY_REAL-Speicherung für DAO-Beziehungen, Indexnavigation und
GUI-Auswahl bei referenzieller Integrität/1:1. Seek akzeptiert jetzt tatsächliche
Currency-Schlüssel und numerische Schlüssel über den vorhandenen Feldcodec;
Text bleibt abgewiesen, Null-/absteigende Bereichsgrenzen bleiben ausdrücklich.
Echte GUI-Beziehung mit Update-/Delete-Kaskade, DAO Gleichheits-/Bereichssuche,
NoMatch/Bookmark und Zurückweisung bei erhaltenem aktuellen Wert geprüft.
App/GUI/DAO gebaut; ein bestehendes Bündel2/2 in14,37s bestanden:
GUI11,91s (20s unverändert), DAO2,46s. Keine neue CTest-/Katalog-ID;456/386/70.
Vorstufe8a198ce2 privat/Baum geprüft; letztes DEB/Original/Inventar bleibt da72e727.
Keine volle Jet-/Windows-Parität; ohne Zwischenabschluss weiterentwickeln.


Berichtsdruckoptionen: REPORT_PRINT_OPTIONS.md. Tatsächlicher Druckdialog bietet
nun Seitenbereich; Renderer berücksichtigt Originalseitennummern, Vorwärts- /
Rückwärtsfolge,1-100 Exemplare und Sortierung. Aufgezeichnete Quellseiten werden
wiederverwendet, keine erneute Berechnung/Eventausführung je Exemplar.
PDF aus dem Druckdialog übernimmt echte Seiteneinrichtung/Optionen und verwendet
nun temporäres Rendern plus atomaren QSaveFile-Commit. Ungültige Bereiche/Render-
fehler erhalten vorherige Datei; Symlink-/NUL-/aktuelle DB-Ziele abgewiesen.
Native Treiberkopien bleiben Treiberaufgabe, sonst kontrollierte manuelle Kopien;
PDF immer explizit. Mehrfache Ausgabe auf10000 Seiten begrenzt. Selection/current
page weiterhin deaktiviert. Gruppierte/andere nicht unterstützte Berichte bleiben
bewusst gesperrt; kein Original-/Hardware-/Windows-Ausgabenachweis.
App/Layout gebaut; bestehendes access_2003_layouts1/1 in10,79s bestanden. Echter
Dialog/PDF/Text beweist beide Kopiensortierungen rückwärts, nurSeite2 mit2/2,
aktuellen Datensatz wiederhergestellt und Bestandsschutz bei ungültigem Bereich.
Erste Fixturefehler waren fehlende Reaktivierung vor geschütztem Druckbefehl und
fehlende NFKC-Normalisierung der PDF-fi-Ligatur; nur betroffenen vorhandenen Check
wiederholt. Keine neue CTest-/Katalog-ID;456/386/70. Vorstufe009be44e privat/Baum
geprüft. DEB/Original/Inventar nochda72e727; gebündelte Aktualisierung folgt.
Ohne Zwischenabschluss weitere allgemeine Anforderungen fortsetzen.


Aktueller Paket-/Original-/Inventarbeleg: Quellcommit299b946ff0aa559f0e4a4893015a9f71b668f605,
Baumc9fa2d1f78ffd08b0102664f079e88743cb491dc, privat remote/Baum geprüft.
/workspace/work/report-print-options/package/linaccess-0.7.0-Linux.deb erstellt,
check-deb bestanden; bestehendes Paket-Integrationsbündel2/2 in168,22s:
Inventar47,60s, real_kassa_dao_event120,62s; sieben geschützte Hashes unverändert.
Dieses Paket enthält initiale Currency-Auswahl, unveränderliche Syntaxmuster,
native Currency-Erstellung/Regeln/Beziehungen/Indexsuche und Berichtsdruckoptionen.
Neues tatsächliches Inventar299b946f:100/78/22 (73 ohne und5 mit technischen
Parametern ausgeführt, jeweils ohne Access-Ergebnisorakel;3 Originalspalten,
9 fehlende Abhängigkeiten,10 Jet-/Bindungsgrenzen). Originalpfad hält bekannte
Makro-/Tastaturgrenzen; weiterhin kein positiver initialisierter Verkauf und kein
Originalbericht-/Hardwaredrucknachweis. Keine vollständige Parität/Distributions-
Matrix/Windows-Referenz neu behauptet.456/386/70 unverändert. Ohne Zwischenabschluss
weiterprogrammieren; nächster allgemeiner VBA-Umfang wird gesondert implementiert.


Native StrConv: NATIVE_STRCONV.md. Gemeinsamer Evaluator unterstützt dokumentierte
ASCII Upper/Lower/ProperCase, Worttrenner NUL/TAB/LF/VT/FF/CR/Space, benannte
String/Conversion/LCID-Bindung, VBA.StrConv und tatsächlichen String-Rückgabewert.
Nicht-ASCII/Codepages/East-Asia/expliziteLCID/Nichttext/Null-Quellvertrag bleiben
mit3251 ausdrücklich außerhalb dieses Umfangs; keine erfundene Windows-Parität.
Native skalare Objekt-Defaults werden einmal gelesen. Keine Host-/Symbol- /
Ergebniscaches neu. App/Core gebaut; bestehendes core_tests1/1 in0,52s bestanden,
inklusive echter VBA-Funktion, sieben Trenner, Getter-Zähler und Schutzgrenzen.
Keine neue CTest-/Katalog-ID;456/386/70. Letztes geprüfte DEB/Original/Inventar
bleibt299b946f (47,60/120,62s, sieben Hashes unverändert), vor dieser Funktion.
Dokumentations-Vorstufe0a045c94 privat/Baum geprüft. Ohne Zwischenabschluss weiter.


Native Rnd/Randomize: NATIVE_RANDOM_GENERATOR.md. Tatsächlicher projektweiter
24-Bit-Zustand im vorhandenen Modulkontext, geteilt von Formular-/Standardmodul-/
Editor-/Klassen-/nativen Modul-SQL-Hosts; getrennte Projekte/Reset/geschlossene DB
geprüft. Rnd liefert Single[0,1), fehlend/positiv nächste,0 aktuellen Wert, negativ
reproduzierbare Folge. Randomize ist Sub mit optionalem benanntem Number; fehlend/
Empty echter Systemtimer, sonst geprüfte Double-Konvertierung. Native Argumente
werden vor Zustandszugriff skalarisiert/konvertiert; fehlerhafte primitive Eingaben
verändern Zustand nicht. Keine Generatorzustände im Syntax-/Quellcache.
Öffentliche MIT-VBMath-Algorithmusreferenz mit Attribution/Lizenz und Installregel;
kein beobachtetes Access-Zufallsfolgen-Orakel, keine Behauptung vollständiger
volatiler Access-/Jet-Ausdrücke (direkt kalkuliertes/Jet-Rnd ausdrücklich3251).
App/Core/GUI gebaut. Bestehendes Bündel: GUI12,37s bestanden; Core fand fehlenden
benannten Randomize-Sub-Fallback. Binder und qualifiziertes argumentloses Randomize
korrigiert; nur Core neu gebaut/geprüft,1/1 in0,49s bestanden. GUI-20s unverändert.
Keine neuen CTest-/Katalog-IDs;456/386/70. Vorstufecde2c756 privat/Baum geprüft.
Letztes DEB/Original/Inventar ist299b946f, vor StrConv/RNG (47,60/120,62s, sieben
Hashes unverändert). Ohne Zwischenabschluss weiter; nächste Berichtsseitenkopf- /
Seitenfußoptionen anhand dokumentierter tatsächlicher Semantik implementieren.


Berichtsseitenkopf/-fuß: REPORT_PAGE_BANDS.md. Tatsächliche Report-Einstellungen
PageHeader/PageFooter0-3 werden streng gelesen, nativ erstellt und im echten
Berichtsentwurf gespeichert/wieder geöffnet. Dokumentierter Default AllPages0.
Bit1 unterdrückt jeweiligen Bereich bei Berichtskopf, Bit2 bei Berichtsfuß und
erzwingt für diesen eine eigene Seite (dokumentierte Access-Regel). Gemischte
Einstellungen bleiben unabhängig; ohne Berichtsfuß keine künstliche Fußseite.
Pagination/Detailteilung nutzen tatsächliche aktive Bereichshöhen, ursprünglicher
Datensatz/Layout werden restauriert; Ranges/Kopien nutzen weiter Quellseiten.
Keine neue Gruppierungs-/Format-/Print-Event-/VBA-Property-Mutation-Parität.
App/GUI/Layout gebaut. Bestehendes Layout1/1 in10,72s bestanden, tatsächliche PDF-
Texte für alle4 Modi und gemischte Kombinationen, Seitenzahlen/Inhaltsgrenzen/
Restaurierung und ungültige Quellen. GUI hatte alte ForceNewPage1-Erwartung,
Fixture verwendet bewusst0 und neue Fußseitenregel. Vorhandenen Wiederöffnungs-
Check angepasst statt zweiten neuen Dialoglauf; nur GUI wiederholt1/1 in12,07s,
20s-Watchdog unverändert. Keine neuen CTest-/Katalog-IDs;456/386/70.
Vorstufeace05fb7 privat/Baum geprüft. DEB/Original/Inventar noch299b946f, vorStrConv/
RNG/Seitenbereichen. Ohne Zwischenabschluss nächste allgemeine Gruppierung umsetzen.


Native allgemeine Berichtsgruppierung: NATIVE_REPORT_GROUPING.md. Echter Gruppen-
und Sortierdialog mit bis10 unterschiedlichen Quellfeldern, Auf/Ab-Reihenfolge,
Gruppenköpfen und Fußberechnungen; Speichern/Wiederöffnen/Quellgenerierung/Seiten-
vorschau/PDF. AST hält tatsächliche Gruppenbereiche/Controls getrennt mit belegter
Stufenbindung. Renderer verlangt alle7 expliziten Gruppenproperties, GroupOn0,
Intervall1, KeepTogether0; keine fehlenden importierten Defaults geraten. Andere
Druckgrenzen bleiben. Stabile temporäre Sortierung, Null-Gruppen, native Text-/
Zahl-/Date-/Currency-Schlüssel, eindeutige Gruppenaggregate und Berichtssummen;
Zeilen/IDs/Snapshot/Scope/aktueller Datensatz/Layout restauriert. Native Text-
Vergleiche sind keine volle Jet-Locale-Parität; binäre/gemischte/fortgeschrittene
Schlüssel/Intervalle/Repeat/KeepTogether/Events bleiben außerhalb dieses Umfangs.
App/GUI/Layout gebaut: vorhandene GUI12,68s bestanden (20s unverändert), gezielt
Layout nach Fixturekorrekturen1/1 in9,99s. Tatsächliche verschachtelte PDF-Summen/
Counts/Null/stabile Marker/First-Last/Restaurierung, Date-Rand, Grenzen/Dateierhalt.
Fixture nutzte zunächst bestehendes SQL-CStr3085 und fehlenden RecordSource-
Override; nun vorhandene Verkettung und tatsächliche Paketquelle wie die App.
Wichtig: Original-Toolbar-Prüffälle lesen nun bisher ausgelassene Gruppencontrols.
Ohne Datenbank verweigern alle3 Layout-PDFs wegen numerischem Format auf berechnetem
Text13. Exakter bekannter Maskenhash/Quelle fehlt/Typ/Error/Dateierhalt geprüft;
aktuell0 positive Layout-PDFs und3 belegte Verweigerungen. Drei Original-Handler
weiter ausgeführt; keine neuen privaten VBA-Aufrufe, keine Initialisierungs- /
Original-Berichtsparität behauptet. Frühere unvollständige positive Layout-
Beobachtung ist für diesen Stand ersetzt. Keine neuen CTest-/Katalog-IDs;
456/386/70. Vorstufe6578f056 privat/Baum geprüft. DEB/Original/Inventar noch299b946f,
vorStrConv/RNG/Seitenbereichen/Gruppierung; neue gebündelte Paketprüfung folgt.
Ohne Zwischenabschluss weiter. Allgemeine SQL-CStr-Lücke separat beheben.


Paket-/Originalbeleg nach nativer Gruppierung (2026-11-10): Anwendungscode
62d00929; korrigierte öffentliche Rechte der neuen MIT-Lizenzverzeichnisse und
strikte Paket-Allowlist. DEB unter work/native-report-grouping/package-retry/
package/linaccess-0.7.0-Linux.deb: Inhalt/Rechte/Build-ID/SQLite/Metadatenhelfer
bestanden. Vorhandenes Inventar49,42s:100/78/22, unveränderte Kategorien73/5/3/9/10,
kein Access-Ergebnisoracle. Vorhandener Originalablauf nach gezielter Korrektur
des Summenvergleichs1/1 in116,96s bestanden; sieben geschützte Hashes unverändert.
Neue Gruppencontrols hatten Annahmen des alten Prüfers offengelegt: direktes
SQLite umgeht die typisierte Quellausführung; die Regex umfasste auch kombinierte
Summen. Separate schreibgeschützte Quellzeilen-/Aggregatbindung rechnet vollständige
Ausdrücke mit unabhängiger Summenakkumulation nach, ohne Berichtscache/-Gruppenscope.
Keine Anwendungsgrenze gelockert, keine neuen CTest-/Katalog-IDs. Original-Layout-
PDFs bleiben0 positiv/3 belegte Verweigerungen, bekannte Makro-/Tastaturgrenzen
bleiben; keine initialisierte Verkaufs-/Windows-/Gesamtparität behauptet.456/386/70.
Weiter mit allgemeinem SQL-CStr, anschließend weiteren offenen Anforderungen.


Allgemeines natives SQL-CStr (2026-11-10): docs/NATIVE_SQL_CSTR.md. Gemeinsamer
VBA-Auswerter und belegte Feld-/Parametertypen erhalten Boolean/Single/Date/
Currency in typisierten Projektionen; Text einschließlich NUL, Null94, echte
BLOB-Felder13, nicht exakt darstellbare große SQLite-Integer6, Arity450.
Tokenbasierte Übersetzung, keine Geschäftsdatenersatzwerte oder Zustands-Caches.
App/DAO gebaut, bestehendes dao_execution1/1 in2,43s bestanden. Fixture erwartete
zunächst1 trotz vorherigem NativeScale-UPDATE auf2 und verwendete SQLite-X-Hex
statt eines echten Binärfeldes; gezielt korrigiert, keine Laufzeitgrenze gelockert.
Kein kompletter Jet-Locale-/Join-/Subquery-Typnachweis, keine neuen CTest-/Katalog-
IDs;456/386/70. PRIVAT/Baum geprüfte Vorstufe5ec76beb; Paket noch Anwendung62d00929
mit Verpackungskorrektur, vor SQL-CStr. Original116,96s/Inventar49,42s/7Hash gleich.
Ohne Zwischenabschluss weiter mit laufenden nativen Berichtssummen je Gruppe.


Native laufende Berichtssumme je Gruppe (2026-11-10): REPORT_RUNNING_SUM.md.
Echter Designer speichert running_sum_mode1/2; alte Definitionen bleiben2.
Ungültige Werte/Typen und fehlender tatsächlicher Gruppenkopf/-fuß werden
verweigert. Detailpräfixe je innerster tatsächlicher Gruppenstufe einmal in
Drucksortierung berechnet, bei Gruppenwechsel zurückgesetzt; Sortierstufen ohne
Gruppenband gelten nicht als Gruppenwechsel. Seiten/Neuzeichnen/Kopien zählen
nicht doppelt. Bestehende Null-/Detail-/CanGrow-/importierte Metadatengrenzen
bleiben. App/GUI/Layout gebaut: vorhandene GUI12,75s (20s unverändert), gezielt
Layout nach zeilenbegrenzter PDF-Regex1/1 in10,19s. Acht exakte Currency-Präfixe
bei verschachtelten/Null-Schlüssel-Gruppen, richtige Gruppen-/Gesamtsummen,
Speichern/Wiederöffnen/Restaurierung geprüft. Keine neuen CTest-/Katalog-IDs;
456/386/70. Vorstufe1579977d privat/Baum geprüft. Paket noch app62d00929 mit
Verpackungskorrektur5ec76beb; vor SQL-CStr/OverGroup. Original116,96s, Inventar
49,42s,7Hash gleich. Ohne Zwischenabschluss weitere offene Anforderungen umsetzen.


Allgemeine native SQL-Zeichenfolgen (2026-11-10): NATIVE_SQL_STRINGS.md.
CStr/Str/Left/Right/Mid einschließlich $-Varianten und Replace verwenden gemeinsame
VBA-Auswertung und belegte Typen aller Argumente. UTF-16-Positionen, optionale
Argumente, Str-Vorzeichen/Fraction, Null94/Stringvarianten, Replace Start/Count/
Compare; NUL bleibt erhalten, Binär/inexakte Integer verweigert. Typisierte
Argumente erreichen innerhalb der bisherigen Quellenresolver-Grenzen auch
Filter/Gruppier-/Sortierausdrücke; ASC/DESC/mehrere Schlüssel bleiben. Vorhandene
numericOnly-Guards für Date-/Currency-Operatoren nicht gelockert, keine allgemeine
Join-/Subquery-/HAVING-/Jet-Locale-Parität. App/DAO/Currency gebaut: vorhandenes
Bündel2/2 in3,02s (DAO2,95/Currency0,07); vorher StringfamilieDAO2,78s. Reale
DAO-Pfade inkl. Boolean-Gruppierung/Sortierung, Date-Filter und vierstelliger
Currency nahe2^39, UTF-16/Null/Optionen geprüft. Keine neuen CTest-/Katalog-IDs;
456/386/70. Vorstufec2c5be8c privat/Baum geprüft; DEB noch62d00929 mit5ec76beb,
vor CStr/OverGroup/Stringadaptern. Gebündelte neue Paket-/Original-/Inventarprüfung
folgt. Ohne Zwischenabschluss weitere offene Anforderungen implementieren.


Native SQL-Suche/Vergleich/Trim (2026-11-10): NATIVE_SQL_SEARCH_TRIM.md. InStr
mit UTF-16-Positionen liefert Long, StrComp Integer; Trim/LTrim/RTrim inklusive
$-Varianten, belegte Argumenttypen, Null/Varianten und optionale Vergleiche über
gemeinsamen VBA-Auswerter. ASCII-/Locale-/LCID-/Surrogat-/Quellenkontextgrenzen
bleiben ausdrücklich bestehen. App/DAO gebaut, vorhandenes dao_execution1/1
in2,79s (gesamt2,80) bestanden, inkl. vorheriger String-/Kontextprüfungen. Keine
neuen CTest-/Katalog-IDs;456/386/70.
Aktueller geprüfter Paket-/Original-/Inventarbeleg: Quellcommit678a7fd82320afe5a255158a4e5d7f9e26b58944,
Baum0fbe2a242c8c908120f6928ae6290d1bed7d1c30, privat verifiziert. DEB
work/sql-string-adapters/package-proof/package/linaccess-0.7.0-Linux.deb: Inhalt/
Rechte/App-/Helper-Build-ID/SQLite bestanden; vorhandenes Bündel2/2 in163,22s
(Original116,95/Inventar46,27).7 geschützte Hashes unverändert.100/78/22, bekannte
Kategorien73/5/3/9/10, kein Access-Ergebnisoracle; Original-Layout-PDF0 positiv/
3 Verweigerungen, bekannte Makro-/Tastaturgrenzen bleiben. Das Paket enthält
CStr/OverGroup/Stringadapter und liegt vor dieser Such-/Trim-Erweiterung.
Ohne Zwischenabschluss weitere allgemeine offene Anforderungen implementieren.


Wiederholte native Gruppenköpfe/Abschnitts-API (2026-11-10):
NATIVE_REPORT_REPEAT_HEADER.md. Echter Gruppendesigner speichert repeat_header,
Quellgenerator explizites RepeatSection und tatsächliche native Abschnittsnamen.
Planer wiederholt aktive Gruppenköpfe mit belegter Gruppenschlüsselbindung,
berücksichtigt Höhe/Seitenrollen und beendet geschlossene Bereiche. Beliebige
zeilenabhängige Headerausdrücke, andere RepeatSection-Kontexte/Spalten/Unterberichte
bleiben außerhalb dieses Umfangs; bestehende Guards/Dateiatomizität bleiben.
Section-Indizes5..24 für belegte Gruppenpaare, öffentliche erste4 AcSection-
Gruppenkonstanten, Name/Index-Identität, Boolean RepeatSection lesen/schreiben
mit Vorabprüfung/Neuzeichnung, Null94/geschlossen2450. Tatsächliche VBA-Umschaltung
und dreiseitige PDFs an/aus, Identitäten zweier Stufen, Restaurierung/Dateierhalt
geprüft. GUI12,75s bestanden (20s unverändert), nach gezielter Fixturekorrektur
Layout1/1 in10,45s; kein neuer kombinierter2/2- oder Windows-Nachweis.
Keine neuen CTest-/Katalog-IDs;456/386/70. Vorstufe00f8e4ef privat/Baum geprüft.
DEB/Original/Inventar weiterhin678a7fd8 (116,95/46,27s,7Hash gleich), vor Such-/Trim
und dieser Erweiterung. Ohne Zwischenabschluss weitere offene Anforderungen.


Typisierte native SQL-Groß-/Kleinschreibung und Rückwärtssuche (2026-11-10):
NATIVE_SQL_CASE_REVERSE.md. LCase/UCase samt $-Varianten, StrReverse/InStrRev
verwenden gemeinsame belegte Argumenttypen; Boolean-Text, exakte begrenzte
Currency-Ziffern und Long-Ergebnistyp von InStrRev. Bestehende ASCII-/Locale-/
UTF-16-Grenzen bleiben. StrReverse(Null)94, nullable Such-/Case-Varianten und
String-Result-Grenzen korrekt unterschieden; keine Windows-Parität behauptet.
App/DAO gebaut, vorhandenes dao_execution1/1 in2,96s bestanden. Keine neuen
CTest-/Katalog-IDs;456/386/70. Vorstufee64854cf privat/Baum geprüft. Paket weiterhin
678a7fd8 (Original116,95/Inventar46,27s/7Hash gleich), vor Such-/Trim-/RepeatSection/
Case-Erweiterungen; neue gebündelte Paketprüfung folgt. Danach echte modellose
DoCmd.OpenModule-Ausgabe für allgemeine VBA-Entwicklung nachvollziehen/umsetzen.
Keine Zwischenabschlüsse; weitere offene Anforderungen kontinuierlich bearbeiten.


Echtes natives modelloses DoCmd.OpenModule (2026-11-10): MODELESS_MODULE_EDITOR.md.
Formular-/Standardmodul-Hosts, optionale/benannte ModuleName/ProcedureName, belegte
Prozedursuche/Quellposition, Deklarationsansicht, vorhandene Fenster/Entwürfe
wiederverwenden. Gemeinsamer bisheriger Editoraufbau heap-eigen mit wertgebundenen
Callbacks; modaler GUI-Einstieg bleibt. Eigene Quellen bearbeitbar, importierte
Standardquellen hier schreibgeschützt ohne Speichern/Löschen/Run/Debug. Laden
führt Prozeduren nicht aus. Optimistische native Katalogrevision schützt vor
Überschreiben; Close mit Speicheroptionen, echter Abbruch, Busy-/Projektwechsel-/
Reentranz-/QObject-Lebensdauergrenzen. Keine volle VBE-/Klassen-/Formmodul-/Library-/
MDI-/Windows-Fehlergleichheit. App/GUI gebaut, vorhandenes mainwindow_documents
1/1 in13,50s bestanden, 20s unverändert. Tatsächliche VBA-Fortsetzung, Navigation/
Wiederverwendung, Revision/Save/Discard/Cancel, geladene Deskriptoren und unberührtes
technisches Importarchiv geprüft; keine neuen CTest-/Katalog-IDs.456/386/70.
Aktueller geprüfter Paket-/Original-/Inventarbeleg:5c82c36546f744621cb1523e2c1d37d3fef15258,
Baum20206cfc219378aadf8127937ceb3a9d2962cac8 privat verifiziert. DEB
work/sql-case-reverse/package-proof/package/linaccess-0.7.0-Linux.deb bestanden;
Original118,77/Inventar46,44s, Bündel2/2 in165,21s,7Hash gleich.100/78/22 mit
Kategorien73/5/3/9/10, kein Access-Ergebnisoracle; bekannte Original-Grenzen bleiben.
Dieses Paket liegt vor OpenModule. Ohne Zwischenabschluss weiterprogrammieren.


Offene native Modulfenster auswählen/neuzeichnen (2026-11-10):
MODELESS_MODULE_EDITOR.md. SelectObject acModule mit explizitem offenem Namen
über Formular-/Standardmodul-Hosts; keine Wiederöffnung oder Entwurfsersetzung,
kein falscher Datenblattkontext. Verbindungs-/Datei-/Generations-/Modal-/Closing-
Guards. Minimierung entfernen, maximierten/Vollbildzustand erhalten; explizites/
aktives Repaint erkennt modellose Module. App/GUI gebaut; vorhandenes
mainwindow_documents1/1 in13,18s bestanden (20s unverändert); tatsächliche VBA-/
Servicepfade, Zustand/Entwürfe/Neuzeichnen/geschlossenes Ziel geprüft. Keine neuen
CTest-/Katalog-IDs;456/386/70. Vorstufe65844569 privat/Baum geprüft. DEB bislang
5c82c365 (118,77/46,44s,7Hash gleich); neue Paket-/Original-/Inventarprüfung folgt.
Weiter mit der dokumentierten VBA-Bezeichner-Typzeichenlücke; keine Zwischenabschlüsse.


Native skalare Bezeichner-Typzeichen (2026-11-10):
VBA_IDENTIFIER_TYPES.md beschreibt %/&/!/#/@/$ an lokalen Dim-/Static-
Deklarationen, skalaren Parametern, Modulvariablen und Konstanten; kanonische
Variablenbindung, deklarationsbasierter Typabgleich und ByRef-Erhalt. Kein
Quelltext-Strippen, kein Variant-Subtype-Ersatz, kein veränderlicher Rahmen im
Syntaxcache. As-Doppelangabe wird verweigert. Vorhandenes core_tests nach
Korrektur der Konstantensyntax/ByRef-Bindung1/1 in0,53s bestanden. Funktionsnamen,
Array-/ReDim-Deklarationen, Schleifen, implizite Typzeichen und qualifizierte
Projektmember bleiben offen. Keine neue CTest-/Katalog-ID;456/386/70 unverändert.
Nächste Implementierung: Typzeichen an Funktionsnamen samt Signatur-/Rückgabebindung.


Typisierte lokale Funktionsnamen (2026-11-10): VBA_IDENTIFIER_TYPES.md ergänzt.
Function-Suffix/kanonischer Name/Rückgabevariable/Signatur-Typprüfung; direkte
Ausführung mit Typzeichen wird validiert, Metadateninventar behält kanonische
Namen und Quellzeilen. Long-Aufruf vor Klammer erst anhand lokaler Signatur als
Typzeichen gebunden; kein globales Entfernen von &. Vorhandenes core_tests1/1
in0,57s bestanden. Property-/Projekt-/Library-Gesamtnamensräume bleiben offen.
Privater skalarer Vorstand4aa7885f74844d3e41e95c0a2b28bb320f9a14f2,
Baum2cb0d4f0f474adce57e6595b59e3b24ab0fb7831 verifiziert.456/386/70 unverändert.
Weiter mit impliziten skalaren Typzeichen, danach Arrays/ReDim; kein Zwischenabschluss.


Implizite und Array-Typzeichen (2026-11-10): VBA_IDENTIFIER_TYPES.md ergänzt.
Skalare implizite Bindung ohne Option Explicit, Vorrang vor Def-Typen,
Literalbezeichner-Suffixe ignoriert; deklarationsbasierter Typabgleich bleibt.
Lokale/Modul-Arrays, ByRef-Arrayparameter, Elemente, ReDim/Preserve und Erase
verwenden kanonische Namen/Elementtypen. Neue ReDim-Ziele auch mit Option Explicit
prozedurweit kanonisch gebunden. Statischer Funktionszustand identisch bei
Aufrufen mit/ohne Suffix. Bestehender alter statischer Suffix-Guardtest aktualisiert,
weil dieses Verhalten nun implementiert ist; keine Wache pauschal entfernt.
Anwendung/Kern/DAO gebaut; bestehendes Bündel2/2 in3,54s (Kern0,57,DAO2,97s).
Keine neuen CTest-/Katalog-IDs, F040/F122 bleiben offene Vollanforderungen;
456/386/70 unverändert. Funktionsvorstand4e68489b948cf8344467e94ae42c60e4e7f44421,
Baum59d5837f3f8651ee9b31901701fbd9a7c6e9c1e4 privat verifiziert.
Nächste Implementierung: Typzeichen an For/Next, direkten Aufrufen und Stringzielen.


Typzeichen in Anweisungspositionen (2026-11-10): VBA_IDENTIFIER_TYPES.md ergänzt.
For/Next und unterstütztes For Each, direkte lokale Function-Aufrufe mit Call/
Anweisungsform, Mid$-Zuweisung und LSet/RSet auf denselben kanonischen Namen.
Typisierte Arrayquellen behalten Live-Mutationswachen. Lokale Signaturprüfung
nach Namensraumtreffer; eingebaute Dollar-Funktionen bleiben verfügbar (zunächst
bestehende Fixed-String-Prüfung fehlgeschlagen, genau diesen Filter korrigiert).
Double# an Namen/Zahlen verwechselt Anweisungsdoppelpunkte nicht mit Datumliteralen;
Uhrzeitliteral geschützt. Vorhandenes core_tests1/1 in0,53s bestanden.
Kein neues CTest/Katalog-Item.456/386/70. Vorstufe88fdf1f0b8a26c9e6769172dc34e9b7f4141a0e3,
Baum158eeec6c3759bd7be7aae583255916945dfdb24 privat verifiziert; App/Kern/DAO
auf dieser Vorstufe2/2 in3,54s. Nächster Paket-/Original-/Inventarbeleg folgt,
danach typisierte Konstanteninitialisierer und weitere offene native Anforderungen.


Konstanten-/Property-Typzeichen (2026-11-10): VBA_IDENTIFIER_TYPES.md ergänzt.
Getrennte quellgebundene Konstantentypen für Initialisierer/Enums/Optional/
Arraygrenzen/Stringlängen, mit Inline-Kontext und ohne Variant-Subtype-Inferenz.
Get-Namenssuffix und passende Let-Werttypen; kanonische Referenzen/Selektoren.
Schreibschutz typisierter Konstanten vor RHS-Effekten; Loop-Vorprüfung verwendet
kanonische Namen/Suffixtyp vor Def. Bestehendes core_tests1/1 in0,54s bestanden;
gezielte Fixes für Property-Inventarfilter und Syntax von Name&() enthalten.
Keine neue CTest-/Katalog-ID;456/386/70. F040/F042/F043/F122 bleiben breit offen.
Paketbeleg bleibt53e91932 (Original115,97/Inventar46,72s;2/2 in162,70s,7Hash gleich).
Weiter mit deklarationsbasierter Typbindung zwischen Standardmodulen.


Typisierte fremde Standardmodulfunktionen (2026-11-10): VBA_IDENTIFIER_TYPES.md.
Kanonische Hostauflösung, deklarationsbasierter Typabgleich, Ausdruck/Call und
ByRef-/Objektargumentmetadaten über tatsächliche Projektquelle; lokale/Instanz-
Priorität und Sichtbarkeits-/Mehrdeutigkeits-/SQL-Wachen bleiben. Fremde Properties
weiter3251. Statisches Inventar: threadlokal128/16MiB geschätzt, voller Quelltext
als Identität, kein Host-/Katalogzustand im Cache. App/Kern/DAO gebaut; bestehendes
Bündel2/2 in3,42s (0,42/3,00), echte getrennte SQL-Module/Integer2/ByRef/Call und
Verweigerung vor Zustandsänderung belegt.456/386/70, keine neuen IDs.
Vorstand89088dd8402cc62741d30fc1ae7a367f97015eac/Baum24a0fe404fa6fde6c1a7cae13ff38986bcb02b4e
privat verifiziert. Paketbeleg bleibt53e91932. Weiter mit Eval-/Event-Referenzen.


Typisierte Eval-/Event-Funktionsbindung (2026-11-10): VBA_IDENTIFIER_TYPES.md.
Eval: tatsächliche öffentliche Signatur/kanonischer Name, Dollarform; Access-/Jet-
Lexikmodus unverändert. Ereignisse: lokale/öffentliche fremde Signaturen und
Typzeichen, benannte Argumente geordnet, Optional-Missing erhalten, keine
fingierten Adressalias-Bindungen. Private fremde Funktionen/Bibliotheksrückfall
abgegrenzt. App/Kern/DAO2/2 in3,50s (0,41/3,09); nach Sichtbarkeitsabschluss
betroffenes core_tests1/1 in0,42s. Analyzer-Include beim Bau korrigiert.
Fester Katalog456/386/70 unverändert, keine neuen IDs. Weiter offenF040:
Single-!-Namen vor Wortoperatoren/Bang-Feldmehrdeutigkeit, vollständige Bindung.
Vorstand89824952415203a67e181c48fa6935cc96c3442c/Baumece81dd9b9b89507f8be5ad14334501906dddd3d
privat verifiziert. Paketbeleg bleibt53e91932, frischer Paket-/Originallauf folgt.


Single-! vor Wortoperatoren (2026-11-10): VBA_IDENTIFIER_TYPES.md erweitert.
Kontextbezogene Tokennamensbindung gegen echte Variablen-/Ergebnistypen, keine
Ausführung in Metadaten-/Syntaxprüfung; Bang-/Member-/Literalnamen erhalten.
Gebundene Syntax vor dem Körper; kontextbezogene Case/Argument/Bounds/Print-Gänge.
Implizite Single-Typen nach erster textlicher Verwendung, Zweige/Def/OptionExplicit
und vorheriger Variant-Deklaration; Intrinsics/Host-/Projektname nicht verdrängt.
App/Kern/DAO gebaut; finales bestehendes Bündel2/2 in3,69s (0,47/3,23s) bestanden:
Deklarationen/Const/Parameter/Modul/hoisted/implizit/Def/Fn/Get/Event/Call/For/Case/
ReDim, Funktion nur1x und echtes DAO-Feld[And] belegt; ungültiger Funktionsbang
vor Körpereffekten verweigert. Keine neuen IDs,456/386/70. F040 bleibt breit offen.
Paketbeleg780a709 (117,91/47,08s,7Hash gleich). Stand780a709/Baum579ff0b2 privat
verifiziert; neue Single-Erweiterung liegt nach diesem Paket.


Typisierte SQL-Kalenderfamilie (2026-11-10): DATE_RUNTIME.md erweitert.
17 Kalenderfunktionen mit nachgewiesenen Date/Boolean/Single/Currency/integralen
Argumenttypen, tatsächlichen Date/Integer2/Integer4/Text-Ergebnissen, geschachtelten
Ausdrücken und unterstützten Filter-/Gruppier-/Sortierpfaden. Sechs Vergleiche
zwischen nachgewiesenen Date-Werten liefern Boolean -1/0; Datumsliteral-Herkunft
bleibt im Lexer erhalten; Date IS NULL/IS NOT NULL behält gebundene Kinder.
Kein TEXT-Inhalt als Date-Typ geraten. Ungültiges UTF-8 im gemeinsamen Scalar-
Adapter vor ersetzender Dekodierung abgewiesen. Bestehende DAO-/Query-Modell-
Prüfung2/2 in3,47s (3,22/0,25s), App gebaut. DateAdd-Typverlust im Filter und
bestehende SQLite-1-Erwartung korrigiert; bestehender ungültiger UTF-8-Fall belegt.
Keine neuen CTest-/Katalog-IDs.456/386/70 unverändert; F067/F068/F071 sowie
regionale Kalender-/Format-/komplexe Quellenanforderungen bleiben offen.
Kein neuer Windows-/Access-Ergebnisoracle.

Letzter Paketbeleg vor dieser Kalendererweiterung: privat veröffentlichter
Stand9413428d938c27ef19cdeca1dcfc7d5bedaee49c,
Baum61740b61e261d6ea954b3347240193e3be67ac57.
work/single-bang-identifiers/package-proof/package/linaccess-0.7.0-Linux.deb:
Bau/CPack/check-deb bestanden, bestehendes Original-/Inventarbündel2/2 in164,29s
(117,39/46,90s),7 geschützte Hashes unverändert. Bekannte Referenzgrenzen bleiben.


Typisierte SQL-Mathematik (2026-11-10): TYPED_SQL_ARITHMETIC.md.
Sqr/Exp/Log/Sin/Cos/Tan/Atn erhalten nachgewiesene Date/Boolean/Single/Currency-
Argumente und deklarierte Double-Ergebnisse in Projektion/Filter/Gruppierung/
Sortierung. Tatsächliches DAO dbDouble ist7. Native Null94/Domain5/Typ13/
Overflow6 bleiben erhalten; kein Windows-Library-Orakel behauptet. App/DAO/
Query-Modell gebaut; Modell0,25s bestanden, betroffener abschließender DAO-Lauf
1/1 in3,36s. Technische Test-Erwartungen dbDouble und Null korrigiert; keine
Produktregel dafür gelockert. Keine neuen IDs;456/386/70, F067/F068/F071 offen.
Vorstand51b8e75ce1ee55827a3c7da243671a821d645fd5/Baumaa57c12dd3cdc7046009a39076fde7c00a7168f6
privat veröffentlicht und exakt verifiziert. Letzter Paketbeleg weiterhin9413428.


Typisierte SQL-Konvertierungen (2026-11-10): SQL_NUMERIC_CONVERSIONS.md.
CBool/CByte/CInt/CLng/CSng/CDbl erhalten nachgewiesene Date/Boolean/Single/
Currency/integrale Argumente, verschachtelte Ergebnistypen und echte DAO-
Metadaten. Projektion/Filter/Gruppierung/Sortierung geprüft; Null94/Overflow6/
Typ13 erhalten. CByte(True)=255 gemäß bestehendem VBA-Kern, nicht roher
SQLite-Integer1; technische Fehlererwartung dazu korrigiert, Produktregel
unverändert. App/DAO/Query-Modell gebaut; Modell0,29s bestanden, betroffener
abschließender DAO-Lauf 1/1 in3.35s bestanden. Keine neuen IDs;
456/386/70 und offene F068/F071/F105/F122 bleiben. Kein neuer Access-Oracle.
Vorstand82c0e230ad72024b7fd5a1df44fd0dae0ebde989/Baum9c6ffccfacd6d23f2afae0ddd119c295d28cb380
privat veröffentlicht und exakt verifiziert; Paket weiterhin9413428.


Typisierte SQL-Formatierungsfamilie (2026-11-10): NATIVE_FORMAT_RUNTIME.md.
Format/Format$, FormatNumber/Currency/Percent und FormatDateTime erhalten
nachgewiesene Date/Boolean/Single/Currency-Argumente einschließlich Optionen.
Projektion/Filter/Gruppierung/Sortierung sowie genaue int64-Eingabe mit typisierter
Nachkommastellenoption geprüft. Strikte Textmasken13/Nullmasken94 und getrennte
Null-Konventionen erhalten. Bestehende Single-Formatpolitik7 signifikante Stellen
und berechnetes TEXT ohne proven Text/Memo-DAO-Type (3251) bleiben; technische
Testannahmen hierzu korrigiert, keine Produktregel gelockert. Tatsächliche Text-
Variants geprüft. App/DAO/Modell gebaut, Modell0,35s bestanden, abschließender
betroffener DAO-Lauf1/1 in3.48s bestanden. Keine neuen CTest-/Katalog-IDs;
456/386/70 unverändert, F067/F068/F071/F121/F122 weiter offen, kein Access-Oracle.
Vorstandf856c97fa7eadc1df38f6e80fad0784b32bc8160/Baum3bec4863b184a479dc1b92d2ad5ceb0813074925
privat veröffentlicht/exakt verifiziert. Frisches Paket folgt; letzter Beleg9413428.


Paketbeleg für native typisierte Scalar-Familien (2026-11-10):
Privater veröffentlichter Stand2d7fa91e0ee8010c1bcc258bb6d84d0df829372e,
Baum812ba0beafe3a152b033afa6a4f8d61139c53938. DEB unter
work/typed-scalar-families/package-proof/package/linaccess-0.7.0-Linux.deb:
Integrationsbau/CPack/check-deb bestanden; vorhandenes Original-/Inventarbündel
2/2 in164,85s (Original118,63/Inventar46,21s). Alle7 geschützten Hashes
unverändert. Inventar100/78/22, Kategorien73/5/3/9/10 unverändert. Die78
Ausführungen haben weiterhin keinen Access-Ergebnisoracle; bekannte Original-
Makro-/Ablaufgrenzen bleiben. Paket enthält Kalender-/Mathematik-/Konvertierungs-/
Formatierungsbindung, liegt vor den neuen Zeichen-/Alias-/Längenänderungen.
456/386/70 unverändert, keine zusätzliche Vollanforderung geschlossen.


Typisierte Zeichen-/Basis-/Längenfamilie (2026-11-10):
CHARACTER_FUNCTIONS_RUNTIME.md und NATIVE_SQL_STRINGS.md erweitert. Hex/Oct,
Space/String, Asc/AscW/Chr/ChrW sowie Len/LenB binden tatsächliche Date/Boolean/
Single/Currency/Integer/Long-Quelltypen; Integer-/Long-Basisbreiten und tatsächliches
DAO dbInteger für Asc geprüft. Hex$/Oct$/Space$/String$/Chr$/ChrW$ mit echten
SQL-Callbacks und getrennten Null-Regeln ergänzt. Len(Date)=8 erhalten, LenB
nicht-textuell3251 statt ISO-Textinterpretation. Unterstützte verschachtelte
Projektion/Filter/Gruppierung/Sortierung; ASCII/Surrogat-/Output-/Quellgrenzen
bleiben. App/DAO/Query-Modell gebaut, bestehendes Bündel2/2 in4,54s (4,25/0,29)
bestanden. Keine neuen IDs;456/386/70 unverändert, keine Windows-Oracle-Behauptung.
Vorstand2d7fa91e0ee8010c1bcc258bb6d84d0df829372e/Baum812ba0beafe3a152b033afa6a4f8d61139c53938
privat verifiziert; dessen Paketprüfung/Originalbündel164,85s/7Hash gleich siehe
Paketbeleg oben. Neue Zeichen-/Längenbindung liegt nach diesem Paket.


Typisierte skalare SQL-Finanzfamilie (2026-11-10): NATIVE_FINANCIAL_FUNCTIONS.md.
FV/PV/Pmt/NPer/Rate, SLN/SYD/DDB und IPmt/PPmt erhalten tatsächliche Date/
Boolean/Single/Currency/integrale Argumente; deklarierte Double-Ergebnisse und
DAO dbDouble7 sowie unterstützte Verschachtelung/Filter/Gruppierung/Sortierung
geprüft. Unabhängige Nullzins-/Zweiperioden-Abschreibungswerte, Single-Rundung,
Null94, Text13 und negative Asset-/Timinggrenzen3251 belegt. Vorhandene Solver/
Domain-/Iterations-/Präzisionsgrenzen unverändert, keine SQL-Arrays hinzugefügt.
App/DAO/Query-Modell gebaut; bestehendes Bündel2/2 in4,19s (3,89/0,30) bestanden.
Keine neuen IDs,456/386/70 unverändert; kein neuer Access-Oracle.
Vorstandf8bb8a6b1016b5224ff19b86652c89f8b9e27885/Baum01cc89b8c8d4424e7778caba70eeb6a4b9685aa0
privat veröffentlicht/exakt verifiziert. Letzter Paketbeleg2d7fa91 (164,85s,
7Hashes gleich) liegt vor Zeichen-/Längen-/Finanzbindung.


Typisierte SQL-Farb-/Partition-Familie (2026-11-10): SQL_RGB_RUNTIME.md,
QBCOLOR_RUNTIME.md, NATIVE_PARTITION.md. RGB/QBColor/Partition binden reale
Date/Boolean/Single/Currency/integrale Quelltypen; Farbprojektionen INTEGER4,
tatsächliches DAO dbLong4/Long-Variant. Partition-Padding/Buckets, Date-Kanäle,
Boolean-Negativ-/Null-/Textfehler sowie Filter/Gruppierung/Sortierung geprüft.
Native Palette/Rundung/Null-/Extrem-/Quellgrenzen bleiben, keine Windows-Oracle-
Behauptung. App/DAO/Query-Modell gebaut; bestehendes Bündel2/2 in4,21s
(3,91/0,30) bestanden. Keine neuen IDs;456/386/70 unverändert.
Vorstande9051bce801df5dd03097e54512f435e80d2ac8e/Baum4475bc5f5bf527cb138642b11a99b65c5711c31d
privat veröffentlicht/exakt verifiziert. Letzter Paketbeleg2d7fa91 (164,85s,
7Hashes gleich) liegt vor Zeichen-/Längen-/Finanz-/Farb-/Partition-Bindung.


Typisierte SQL-Rundung/Normalisierung (2026-11-10): NUMERIC_FUNCTIONS_RUNTIME.md,
ROUND_RUNTIME.md und SQL_NUMERIC_CONVERSIONS.md. Abs/Int/Fix/Sgn/Round/CCur
binden nachgewiesene Quelltypen und geprüfte polymorphe Resultate; Int/Fix/Sgn
als native SQL-Callbacks ergänzt. Integer/Long-Abs-Overflow6, Currency-/Single-
Resultate mit tatsächlichen DAO-Typen, Null-/Präzisionsgrenzen und rohe exakte
SQLite-Abs-int64-Berechnung erhalten. Round(Single) bleibt3251. Unterstützte
Projektionen/Klauseln geprüft; sonstige Date/Currency-Operatorgrenzen bleiben.
Tatsächlichen Fehler13 bei CCur(Date) behoben: MS-VBAL5.5.1.2.3/6.1.2.3.1.3
fordert den Date-Double-Serienwert durch den vorhandenen Currency-Konverter.
Positive/negative Bruchteile im Kern und Date-Feld in SQL geprüft. App/Kern/DAO/
Query-Modell/Currency-Ziel gebaut; Modell0,31s und Currency-Ergebnisse0,07s
bestanden. Nach Date-Fix betroffenes Kern-/DAO-Bündel2/2 in4,47s (0,40/4,07)
bestanden. Keine neuen IDs,456/386/70 unverändert, kein Access-Oracle behauptet.
Vorstand3d2ff1f12bc33fa3b103cc8633f79bef7b5d27d6/Baum7da7aa9200277987742a3a1a6d06b5d7bc819051
privat veröffentlicht/exakt verifiziert; letzter Paketbeleg2d7fa91 liegt davor.
