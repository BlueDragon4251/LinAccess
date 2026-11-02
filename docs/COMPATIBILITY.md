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
| VBA-Ausführung | **Experimentelle Teilmenge**: Variants, Ausdrücke, If, For/For Each/Do While, DAO, parameterisierte Sub/Function-Aufrufe, ByRef-Aliase, Label-Fehlerbehandlung/Resume und gebundene Ereignisse. SQL-Aktionen ändern die Projektkopie tatsächlich | Vollständige VBA-Syntax und Semantik | Klassen, AST/Grammatik, Modulzustand, Arrays, Optional/ParamArray, Date/Decimal und vollständige Currency-Koercionen fehlen; [Laufzeit](DAO_RUNTIME.md) |
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
