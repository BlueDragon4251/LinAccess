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
