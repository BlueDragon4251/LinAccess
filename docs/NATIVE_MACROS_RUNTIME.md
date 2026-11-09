# Bounded native macros

Native projects can create and reopen macros through **+ Makro** and the Makros
explorer group. The real action editor adds/removes/reorders actions and saves
its definition in the project's `_la_objects` catalog. It has no prefilled
business records or SQL. Names are limited to64 UTF16 units; macro-group names
containing a dot are outside this implementation.

`DoCmd.RunMacro(MacroName, RepeatCount)` loads the actual saved definition on the
current project connection. Named arguments use the existing VBA binder. Supported
actions delegate to the existing DoCmd implementations and their database,
transaction, document, modal and error guards:

| Action | Supported arguments |
|---|---|
| RunSQL | SQLStatement, optional Boolean UseTransaction |
| OpenTable | TableName, optional numeric View/DataMode |
| OpenQuery | QueryName, optional numeric View/DataMode |
| OpenForm | FormName, optional numeric View |
| OpenReport | ReportName, optional numeric View |
| RunCommand | numeric Command |
| Requery | optional ControlName |
| StopMacro | none; terminates this RunMacro invocation, including repetitions |
| CancelEvent | none; cancels a supported current bound event |
| RunMacro | MacroName, optional RepeatCount; bounded nested call |

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
10000 planned actions per call. Nested calls share a10000-visited-action limit
(including skipped conditions) and a16-active-call limit. RepeatExpression and
macro groups remain unsupported. Stops/lifetime checks run before each action.

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
importer or an MS Access execution oracle. Complete Access condition/continuation semantics, macro groups,
embedded/data macros, RepeatExpression, full action vocabulary/expressions,
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
