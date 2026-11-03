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

21 originale AutoNumber-Markierungen bleiben erhalten; 16 einzelne INTEGER-
Primärschlüssel verwenden AUTOINCREMENT, fünf Zähler ohne Primärschlüssel
persistente native Standardwert-Generatoren. Das ist weiterhin keine vollständige
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
| `.mdb`-Tabellen | MDBTools: typisierte INTEGER/REAL/TEXT-Kopie inklusive NULL, BLOB-Spalten ohne Nutzdaten; **29/29** Tabellen der echten Kassen-MDB in SQLite-Kopie importiert | Vollständiger Jet-Schemaimport | Standardwerte, echte OLE-Daten, exakte Jet-Typen, Encoding |
| `.accdb`-Tabellen | Nur soweit installierte MDBTools-Version liest | Zusätzlicher Importadapter | Versions- und Feature-Abhängigkeit |
| Tabellen-/Datensatzeditor | Ja, für SQLite | Constraints, Indizes, Migrationen | Noch rudimentär |
| Access-Abfragen | Jet-Adapter mit typisierten PARAMETERS/QueryDefs, grundlegenden Strings/LIKE/TOP einschließlich Sortiergleichständen, Datumswerten und eager IIf; Navigator-Ergebnisansicht. 100 Originalabfragen klassifiziert, 77 ausgeführt (davon fünf mit Datenparametern) ohne Access-Referenzergebnis | Jet/ACE-SQL vervollständigen | Funktionen, implizite Formularparameter, Crosstab, genaue Typ-/Ausdruckssemantik; [Details](QUERY_ADAPTER.md) |
| Access-Formulare | 47 Originalformulare mit 879 geparsten Controls / 86 sichtbaren CommandButtons; unterstützte Qt-Widgets aller 47 Formulare instanziiert. Direkte Tabellen editierbar, gespeicherte Abfragen begrenzt schreibgeschützt. Gebundene Lifecycle-/Validierungsereignisse getestet | Access-Steuerelemente und Layouttreue ausbauen | ActiveX, verknüpfte Unterformulare, Layoutdetails, vollständige Ereignisse; [Events](FORM_EVENTS.md) |
| Access-Berichte | Original-Layoutvorschau mit einseitigem Qt-Druck/PDF; drei originale Open-/Close-Paare automatisch ausgeführt, tatsächliche CommandBars-Aktionen angepasst; 6/9 Quellen binden 54 nicht-NULL-Textfelder | Report-Modell vervollständigen | Keine vollständige Daten-/Gruppen-/Seitenverarbeitung oder Access-Druckgleichheit; [Druckadapter](PRINT_PREVIEW_ADAPTER.md) |
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
