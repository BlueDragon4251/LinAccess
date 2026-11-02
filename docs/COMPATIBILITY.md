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
