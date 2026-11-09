# Bounded native macros

Native projects can create and reopen macros through **+ Makro** and the Makros
explorer group. The real action editor adds/removes/reorders actions and saves
its definition in the project's `_la_objects` catalog. It has no prefilled
business records or SQL. Stored object names are limited to64 UTF16 units and
contain no dot; grouped calls use ObjectName.SubmacroName.

`DoCmd.RunMacro(MacroName, RepeatCount)` loads the actual saved definition on the
current project connection. Named arguments use the existing VBA binder. Supported
document actions delegate to the existing DoCmd implementations and their database,
transaction, document, modal and error guards:

| Action | Supported arguments |
|---|---|
| RunSQL | SQLStatement, optional Boolean UseTransaction |
| OpenTable | TableName, optional numeric View/DataMode |
| OpenQuery | QueryName, optional numeric View/DataMode |
| OpenForm | FormName, optional numeric View |
| OpenReport | ReportName, optional numeric View |
| RunCommand | numeric Command |
| SelectObject | numeric ObjectType, ObjectName, optional Boolean InDatabaseWindow |
| Close | optional numeric ObjectType, ObjectName, numeric Save |
| DeleteObject | explicit numeric ObjectType and ObjectName |
| Rename | NewName, explicit numeric ObjectType and OldName |
| Requery | optional ControlName |
| RunCode | one supported Function call, including parentheses; return ignored |
| StopMacro | none; terminates this RunMacro invocation, including repetitions |
| StopAllMacros | none; terminates the active native macro invocation tree |
| Echo | Boolean EchoOn, optional single-line status text |
| CancelEvent | none; cancels a supported current bound event |
| RunMacro | MacroName, optional RepeatCount/RepeatExpression; bounded nested call |

The editor uses the Access action names, ordinary argument cells and Ja/Nein for
Boolean arguments. Required text cannot be blank. Numeric arguments in persisted
actions must be exact signed32-bit integers. Optional arguments are trailing;
there is no expression parser hidden in these cells. A SQL argument is passed as
SQL to the existing checked executor. Literal action fields do not imply support
for Access macro argument expression evaluation.

The entire native definition is structurally validated before its first action:
version1, known root/action attributes,1..256 actions, strict argument types and
arity, optional condition expression syntax, valid UTF16/no NUL,262144 UTF16 units per text argument and1MiB serialized
size. Unknown actions/attributes refuse3251; malformed values refuse
13/450/5 as appropriate. Duplicate case-insensitive catalog names refuse3251.
Runtime RepeatCount uses existing Long coercion, then requires1..1000 and at most
10000 planned actions per count-only call. Predicate-driven calls enforce the actual visited-action limit. Nested calls share a10000-visited-action limit
(including skipped conditions) and a16-active-call limit. Native macro groups are described below. Stops/lifetime checks run before each action.

A macro is a sequence of actions, not a single database transaction. Each RunSQL
keeps its existing UseTransaction contract. If a later action fails, earlier
successful actions remain committed. The original runtime error number and error
metadata are retained, with macro name, one-based action number and action name
added to the diagnostic. A missing macro reports2485. Unknown action syntax at
the end of a definition prevents earlier SQL from executing; a runtime failure
in a valid later action does not roll back earlier completed actions.

`CurrentProject.AllMacros` uses the existing guarded live project catalog:
Count, Item, enumeration, readonly Name/IsLoaded and project-generation lifetime.
SysCmd GetObjectState accepts acMacro4 for real tagged native designer state.
Enumeration of a stored source is metadata, not proof that its content executes.

This is the native project format, not an Access2003 SaveAsText/binary macro
importer or an MS Access execution oracle. Complete Access condition/continuation semantics,
embedded/data macros, full action vocabulary/expressions,
complete macro-bound event dispatch, AutoExec, library lookup, a macro debugger and complete
Access timing/error identity remain open under F088/F050/F057/F148. Catalog456 /
384 confirmed partial items /72 open full requirements remains unchanged.

Public RunMacro contract consulted before implementation:
https://learn.microsoft.com/en-us/office/vba/api/access.docmd.runmacro
It is current documentation, distinct from a Windows Access2003 observation.

Validation: existing mainwindow_documents/access_2003_layouts2/2 (17.16s),
matching app build. Actual editor/reopen, saved catalog, named VBA invocation,
repeated SQL effects, StopMacro, structural preflight without prefix SQL effects,
error3078 at action2 with prior committed effects and live metadata are covered.
No new test target, full-suite run, platform matrix or Windows oracle is claimed.


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


## Event cancellation

CancelEvent sets the current bound event's cancellation flag for supported Open,
Unload, NoData, Dirty, BeforeInsert, BeforeUpdate and DblClick dispatch. Nested
events restore the prior scope. The macro keeps running; cancellation is passed
to existing event validation/close/open handling. Calls outside supported scopes
refuse3251, with no claim about complete Access error precedence. Keyboard and
right-button MouseDown cancellation remain unsupported. GUI/layout2/2 (17.08s)
verify actual validation/open/unload cancellation and scope cleanup. The existing
original corpus additionally passes1/1 (114.46s), after fixing obsolete raw-display
assumptions and restoring underlying invalid-WhereCondition error3075 before
derivative initialization failures. No original sale completion or type/rendering
oracle is implied. Full macros/event/Access equivalence remains open.

Contract consulted: https://learn.microsoft.com/en-us/office/vba/api/access.docmd.cancelevent


Native form designers now assign supported event bindings through a real dialog;
see [NATIVE_FORM_EVENT_DESIGNER.md](NATIVE_FORM_EVENT_DESIGNER.md). Macro name lookup
and save conflict checks use the existing native Qt case-insensitive convention,
including Unicode, with a1M-name/5-second forward scan and exact definition lookup.
Ambiguous names refuse before execution. This does not establish Jet collations.
GUI/layout2/2 (18.42s); app built; no new target or Windows oracle.


## Conditional actions and nested macros

Each action may carry an optional string `condition`; blank means unconditional.
The real action editor saves/reopens/reorders this column with the action.
Nonblank conditions are limited to4096 UTF16 units, checked for supported native
expression syntax before any action starts, then evaluated exactly once before
that action on each repetition. False/Null skip it using existing reduced native
Eval/If semantics. Expressions run in the caller's existing Eval context, including
supported control references and public standard-module Functions. This does not
add general CodeBehind-function or complete Access database-expression scope.
Function side effects remain ordinary real effects; the condition is not evaluated
again for Boolean conversion. Lifecycle is rechecked before performing the action.

RunMacro actions use the same loader and argument contracts. StopMacro exits only
the current invocation (including its repetitions); the parent can continue.
Nested calls share execution limits; exceptions and normal return restore depth
and root budget. Earlier successful SQL is retained on later runtime/budget
failure, consistently with existing per-action transaction boundaries. Syntax/type
preflight of a complete stored definition precedes that invocation's effects;
this is not recursive static validation of every future dynamic target.

Validation: app built, existing layout target passed1/1 (11.21s) in the initial bundled run;
after moving the technical condition Function to an actual public standard-module
artifact, only the affected GUI target was repeated and passes1/1 (6.43s). Actual
condition editor persistence, True/False/Null, once-per-repetition Function SQL
effects, invalid-later-condition preflight with no prefix writes, nested repetition,
child StopMacro with continuing parent, recursive depth refusal/recovery and real
BeforeUpdate blocked-versus-allowed SQL persistence are verified. No new test
target, Windows oracle or full-suite claim; F088/F041/F050/F057/F148 stays open.

The current .deb also passes the existing payload/permissions/Build-ID/metadata
helper checks and the actual SQLite3.54.1 backend self-check. No reference database
or source is shipped. This package check does not claim a fresh platform matrix.


## RunCode Function actions

RunCode is a saved macro action, not a DoCmd method. Function Name requires one
unqualified Function invocation with parentheses and no leading equals sign.
The outer-call scanner handles nested argument calls, doubled string quotes,
bracketed identifiers and date literals; the existing expression parser checks
the argument grammar. Trailing arithmetic/additional expressions are refused
before any action of the macro executes. This bounded implementation accepts
ASCII procedure identifiers and at most65536 UTF16 expression units.

In an actual form/report event binding, the caller's CodeBehind Function takes
precedence, including a local Private Function. Outside an event, lookup uses
public standard-module Functions. A Sub is refused3251; a missing/private standard
Function reports2517. A standard-module name colliding with the target Function
is refused3251 rather than choosing an ambiguous namespace. Event-context markers
are restored after nested dispatch. The return value is discarded. Supported
native expression/argument/compiler, shared-state and lifetime limits still apply.
Raised errors retain their number/source/help and receive the existing macro/action
diagnostic context. No imaginary DoCmd.RunCode interface is exposed.

Public contract consulted: Microsoft Learn, RunCode macro action,
https://learn.microsoft.com/en-us/office/client-developer/access/desktop-database-reference/runcode-macro-action .
This is current documentation, not an additional Access2003 execution observation.

Validation: existing GUI/layout checks2/2 passed (18.17s), app and matching
DEB checked. Actual Function effects, nested macro/event scope, quoted/nested
arguments, Sub/Private refusal, structural preflight and error source/action
metadata are covered. Original uploads3/3 and fixtures4/4 hashes unchanged.


## Run from the database window

Select a native macro in the explorer and choose **Ansicht → Ausgewähltes Makro
ausführen**. This uses a real project/database application host without creating
a hidden form. The same runner as event macros enforces repeat/condition/action
preflight, nested16-call/10000-action budgets, StopMacro and diagnostic metadata.
A VBA execution scope covers the complete macro. Earlier successful actions remain
committed when a later action fails.

The standalone DoCmd proxy provides RunMacro/RunSQL, table/query documents,
stored native normal forms and report previews (explicit View2), RunCommand and
Requery. Saved Functions run in the real standard-module/DAO context. Actual
active forms receive document commands; otherwise existing table/query APIs run.
RunSQL preserves UseTransaction/savepoint behavior and refuses SELECT2342.
CancelEvent requires a real event; an unbound standalone action refuses3251.

Opening forms/reports in this host is currently limited to native stored sources
and the name/view arguments. Imported archive-only documents, named filters,
WhereCondition/OpenArgs/data/window modes and complete Application identity/model
remain open here. Modal editors block opening independent UI documents2046.
Report normal/print View0 remains unsupported3251. New document initialization
failure or cancellation is surfaced rather than reported as a successful macro.
This is still not an Access macro importer, AutoExec or complete macro parity.

Validation: existing GUI/layout2/2 passed (19.15s), app and matching DEB checked.
A real menu invocation exercises SQL/Function effects, nested StopMacro, conditions,
Table/Query/Form/Report/Requery/RunCommand, event SQL and actual bound row count.
Later-action persistence, recursion cleanup, SELECT refusal and real cancelled
Form_Open are exercised. No new test target or Windows Access2003 oracle.

Native project [AutoExec and Shift bypass](NATIVE_AUTOEXEC_RUNTIME.md) are now
implemented for stored native macros; archive macro importing remains open.


## RepeatExpression

RunMacro now evaluates a supported RepeatExpression before each attempt. With
only RepeatCount it runs at most that many times; with only RepeatExpression it
continues until the expression becomes false. With both, the first limiting
condition wins. Both omitted still run once. Initially false/Null expression
results run no actions. StopMacro still terminates the current invocation and its
remaining repetitions. Parent macros continue after a nested StopMacro.

The shared guarded evaluator is used in real form/event and standalone project
contexts. A text predicate is limited to4096 UTF16 units with valid expression
syntax before executing the called macro. API scalar numeric/Boolean constants
are converted through the native String/expression path; full Access Variant and
expression scope/coercion remain bounded. Predicate errors retain their original
number/source/help and add macro name/repetition context. Public-function Eval
now preserves those error fields and explanation instead of dropping them.

Explicit repeat counts remain1..1000. Predicate-only calls permit at most1000
completed iterations; a further true attempt fails3251, while a terminating false
predicate returns normally. The existing shared16-active-call/10000-visited-action
budget also applies, including nested and skipped actions. Predicates can have
real supported side effects; the terminating false evaluation is observable.
There is no global macro transaction or rollback of prior completed iterations.

The editor accepts a RepeatExpression in RunMacro argument3 with argument2 blank.
Only this optional count slot uses JSON null as Missing in the native project
format. It is not an Access macro JSON format or a general Null-to-Missing rule.
Stored predicates are syntax-checked with the whole definition before its first
action, including when a later RunMacro action contains an invalid predicate.

Contracts consulted: Microsoft Learn,
https://learn.microsoft.com/en-us/office/vba/api/access.docmd.runmacro and
https://learn.microsoft.com/en-us/office/client-developer/access/desktop-database-reference/runmacro-macro-action .
These document the optional count/predicate combinations, not a fresh Access2003
execution oracle or full macro compatibility proof.

Validation: existing layout1/1 passed (11.44s); after fixing the actual Eval
error-metadata defect, core/GUI2/2 passed (9.32s), matching app build. Predicate
SQL effects/evaluation count, argument combinations, false-first, iteration bound,
preflight without prefix effects, error/help metadata and actual editor save/reopen
with omitted count are exercised. No new CTest target or Windows oracle.


## Native macro groups

Version1/actions remains the single-macro format. Version2/groups is an ordered
array of objects with name/actions. It permits1..64 submacros, at most256 total
actions and1MiB serialized program. Names are valid UTF16, nonempty/trimmed,
at most64 units and contain no dot. Case-insensitive duplicate labels refuse3251.
The entire group, including inactive submacros, is structurally preflighted before
any selected actions execute. A malformed later submacro cannot allow a prefix
SQL effect from an earlier selected submacro.

RunMacro ObjectName starts only the first ordered submacro. ObjectName.SubmacroName
selects it case-insensitively; a missing submacro reports2485. Nested calls, StopMacro,
RepeatCount/RepeatExpression and shared action/depth budgets retain their existing
contracts. Native form/control event bindings and their designer accept these
qualified macro calls. Group members do not become separate stored project objects.

The actual editor adds a group mode, submacro name/selection, add/remove and ordering.
It preserves raw drafts across switches, including invalid not-yet-saved cells.
Selection/rendering is transient and does not mark saved source dirty. Converting
a multi-submacro draft to a single macro is refused until the user reduces it to
one submacro. Save validates all groups. Active table-cell editing is committed
before serialization. Renamed labels do not automatically refactor callers,
VBA strings or event bindings. This remains a native project format, not an Access
binary/SaveAsText macro importer or full macro compatibility claim.

The RunMacro action documentation linked above specifies group-qualified calls
and first-submacro execution when only the group name is supplied. No fresh
Windows Access2003 oracle has been added.

Validation: existing GUI/layout2/2 passed (20.10s), app and matching DEB checked.
Actual editor/CtrlS/reopen/switch/reorder with changed default execution, Unicode
selection, nested parent/StopMacro/repeat, real Form_Load binding and whole-group
preflight/duplicate refusal without SQL effects are exercised. No new test target.

[Macro storage/deletion](NATIVE_MACRO_STORAGE.md) now guards names, revisions
and active invocations; native names are fixed after save. Imported archives
remain outside write/delete ownership.

[Native macro project commands](NATIVE_MACRO_PROJECT_COMMANDS.md) provide real
Close/DeleteObject/Rename operations under project and document execution guards.


## StopAllMacros and Echo

StopAllMacros is a saved action with no arguments, not a DoCmd VBA method. It stops
the current nested native invocation tree, subsequent actions and repetitions.
The shared state is checked after condition/repeat-expression evaluation and
before the next action. It clears when the outermost invocation unwinds, so later
independent macros run normally. Earlier committed SQL is not rolled back.

The implemented stop concerns macro dispatch. Native VBA Function code that called
a nested macro can finish its own statements; its parent macro still stops before
later actions. Complete Access/VBA stop propagation and arbitrary reentrant roots
remain open. No application-Quit/global VBA cancellation is fabricated.

Echo actions and the standalone DoCmd/Application Echo members use the existing
real GUI update suppression/status surface. StopAllMacros restores that Echo state,
preserving previously disabled surfaces. The public contract also mentions
SetWarnings restoration; full native SetWarnings state is not implemented and no
such restoration/equivalence is claimed here.

Contract consulted: Microsoft Learn,
https://learn.microsoft.com/en-us/office/client-developer/access/desktop-database-reference/stopallmacros-macro-action .
This is documentation, not a new Windows Access2003 execution observation.

Validation: existing GUI/layout2/2 passed (19.55s), matching app build. Actual dialogs
observe suppressed updates, then nested stop/Echo restoration in both hosts. SQL
order/VBA continuation, omitted parent actions/repetitions, later independent start,
stop from repeat evaluation, false conditional stop, named Application.Echo and
invalid status preflight without UI mutation are exercised. No new test target.

## Complete native action arguments and MessageBox

The macro designer retains arguments beyond the three summary columns. Its
Argumente… dialog edits every argument declared by the selected native action;
the condition remains separate. Extra values survive group switches and row
moves. Cancelling the argument dialog preserves the original draft. Changing to
an action with fewer arguments retains surplus drafts for explicit clearing or
validation rather than silently discarding them.

Native MessageBox accepts Message, optional Beep (default true), Type and Title.
The native JSON Type is an explicit text token: None, Critical, Warning?, Warning!
or Information (case insensitive); this is not an imported Access macro binary
encoding. A blank Type means None. A blank Title uses LinAccess. The designer
allows a blank Beep argument before later arguments and serializes its documented
true default. Message argument source is limited to 255 UTF-16 code units and may
start with = for actual evaluation in the current form/standard-module context.
Expression errors propagate with the macro action diagnostic. Result text is
bounded and validated before GUI display. @ separates a bold heading, body and
spaced final section; literal markup is escaped. The action shows an actual
OK-only Qt dialog with the requested icon and optional system beep. It is a macro
action, not a fabricated DoCmd.MessageBox member or a VBA MsgBox replacement.

Whole-definition preflight checks expression syntax and known Type tokens before
any action runs. General imported Access macro decoding, exact Access 2003 dialog
rendering/formatting/localized enum encodings and arbitrary macro expressions
remain open. Public contract consulted:
https://learn.microsoft.com/en-us/office/client-developer/access/desktop-database-reference/messagebox-macro-action .
This is documentation, not a new Windows execution observation.

Validation: existing GUI/layout2/2 passed (19.30s), matching application build.
Actual argument editing/cancel/group persistence and both real host dialogs are
covered; no new CTest target or Access 2003 rendering oracle.

## Four-argument GoToRecord and multiline argument editing

The native GoToRecord action accepts optional ObjectType, ObjectName, Record,
Offset. Native JSON null represents an omitted argument for this action; it does
not inject VBA Null. The designer emits these holes for blank numeric options.
Direct VBA Null arguments still use normal runtime conversion/error semantics.
The corresponding standalone Application.DoCmd.GoToRecord member supports named
arguments and dispatches to actual open table/query documents or an actual open,
visible project form. Unnamed calls use the active form or data document; explicit
names do not create missing documents. It uses existing record-save/event,
query-readonly, row-order, fetch and navigation-boundary rules. Reports and other
unsupported types are refused. Navigation under a modal designer remains blocked
by the existing context guard; arbitrary hidden/modal form discovery is not added.

The detailed action editor uses multiline plain text widgets for text arguments,
including SQL and MessageBox source. Numeric/Boolean arguments use single-line
inputs. No newline flattening or rich-text interpretation is required to edit an
existing source argument. All values are still validated by the common serializer
and whole-definition parser at Save.

Public method contract:
https://learn.microsoft.com/en-us/office/vba/api/access.docmd.gotorecord .
This is documentation, not a new Windows Access 2003 observation.

Validation: app built; layout1/1 (11.40s) and corrected readonly-query GUI1/1
(7.61s) passed. Actual table/query/form positions, named module calls, macro
omissions in both hosts, invalid offset preservation and multiline text covered.

## Seven-argument FindRecord from modules and macros

The standalone DoCmd.FindRecord member uses FindWhat, Match, MatchCase, Search,
SearchAsFormatted, OnlyCurrentField and FindFirst, including named VBA arguments.
It dispatches to the active actual project form or existing table/query search;
it does not fabricate a caller form. The argument parser is shared with the
existing form host, preserving typed value normalization, option validation,
Null refusal and wildcard checks. Exact Currency/Date/numeric/Boolean search
remains subject to the existing supported field types and whole-field bounds.
Formatted searches still require text. Existing selection, validation and search
bounds apply; complete Access no-match/error/timing equivalence remains open.

The native FindRecord macro action has seven arguments. Its FindWhat is literal
text in the native JSON action format; this does not claim arbitrary Access macro
expression or original macro-binary decoding. Optional JSON null values are
omissions. The complete argument dialog retains all seven options and emits
omissions for blank optional numeric/Boolean fields. Typed FindWhat values are
available through the actual VBA member rather than invented JSON subtype tags.

Contract consulted:
https://learn.microsoft.com/en-us/office/vba/api/access.docmd.findrecord .
No new Windows Access2003 observation is claimed.

Validation: app built; existing layout1/1 (11.23s) and corrected GUI1/1 (7.91s)
passed. This also fixed missing initial table current-field selection and made
whole-record selection independent of the initial Qt selection anchor. The new
search fixture now specifies its expected ascending order explicitly.
