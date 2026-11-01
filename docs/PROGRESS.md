# Verified progress

Overall completion is not quantified: the complete Access/VBA/UI specification
remains open. Counts below are measurements, not a percentage compatibility claim.

| Measure | Verified |
|---|---:|
| Original Kassa tables compared cell-by-cell, including NULL | 29/29 |
| Original Bridge artifacts hash-checked | 192/192 |
| VBA files/procedures included in syntax inventory | 32/187 |
| Original register keypad button/parameter-helper executions | 11 |
| Original DAO insertion button execution | 1 |
| Original bound Report_Open / Report_Close procedures executed | 3 / 3 |
| Original simple Form_Load / verified property assignments | 1 / 6 |
| Original complex Form_Load / independently verified dynamic Tags | 1 / 15 |
| Available embedded sources in that original form | 7 |
| Calculated controls / remaining error values in that form | 34 / 1 |
| Original report sources classified / currently bound | 9 / 6 |
| Original report non-NULL textbox values compared with prepared reader | 54 |
| Original form layouts scanned | 47 |
| Parsed controls / command buttons (81 declared visible before VBA) | 879 / 86 |
| Original query artifacts classified / executed without Access result oracle | 100 / 77 (72 ordinary + 5 parameterized) |
| Original date-filter fields exercised through actual calendar selection | 2 |
| Original generated record-search procedures / native selections / missing control-reference failures | 2 / 2 / 1 |
| Original public Eval function executions on prepared native numeric/Null states | 20 |
| Original deletion procedure / direct invocation / bound native click on synthetic rows | 1 / 1 / 1 |
| Original duplicate handler / handled manual-key conflicts via invocation and native click / inserted rows | 1 / 2 / 0 |
| Original counter flags / single-PK generators / managed non-PK generators on new imports | 21 / 16 / 5 |
| Original clock Timer handler / independently checked timestamp targets on prepared layout | 1 / 2 |
| Original report-preview handler / direct invocation / prepared bound click | 1 / 1 / 1 |
| Current local regression tests | 21 passed |

Recent completed work includes native combo/platform/backend fixes, control event
prefixes, InputBox and DAO clone/bookmark/find form navigation. Positional
Sub/Function calls, ByVal/ByRef aliases, function expressions, top-level label
error handling and original keypad regression remain covered.

Native Eval/Application.Eval now executes bounded Access expressions and public
Functions, including real Forms references, SQL writes, statement invocation and
guarded form self-close followed by a modal dialog. Private Functions and Subs
are rejected before their effects execute. The unchanged original module with
15 Eval references executes in 20 prepared numeric/Null/mixed control states.
Counted real Forms lookups verify the additional lookup per Null branch; complete
project context and an Access 2003 result oracle remain open. See EVAL_RUNTIME.md
for the exact regression boundary.

Current completed query package: typed named parameters, imported/temporary DAO
QueryDefs, basic Jet strings/LIKE/DISTINCTROW/TOP ties, canonical dates and bounded
execution. Full Jet/ACE SQL and original Kassa operation remain open.

Completed event package: bound-form lifecycle, validation/cancellation, navigation
save, close handling and synchronized bindings (FORM_EVENTS.md).

Completed typed-field package: Qt/native form/DAO values, all 16 Kassa Boolean
columns, True=-1 normalization and explicit numeric/date validation (TYPED_FIELDS.md).

Completed query UI package: translated imported SELECT/action queries with typed
parameter prompts, a read-only result model, and bounded saved-query form sources.
DAO read-only results share the same executor. TOP, NULL, parameters, error clearing,
DML rejection in result models and named TOP form sources have regression coverage.

Completed iteration package: native For Each/Next, Exit For, DAO Fields collection,
nested loops and error-handler loop restoration; schema counts verified on all
29 original Kassa tables. Original collection-dependent procedures remain open.

Completed conditional SQL package: native eager IIf argument evaluation with
selected-value type/NULL preservation and connection-reopen regression. Other
SQLite/Jet expression semantics remain explicitly incomplete.

Completed imported-form object package: project-scoped Forms and Controls
collections, real cross-form data writes/saves, guarded Qt object lifetimes and
bang/property namespace separation (FORM_OBJECT_MODEL.md). All 47 original form
layouts instantiate in the structure regression; full execution remains open.

Completed aggregate package: First/Last with NULL endpoints and copied per-group
values; original query executions increased to 38/100 without an Access oracle.
Cross-engine ordering and remaining dialect dependencies remain open.

Completed print-preview adapter package: real Qt print/PDF actions and Office
CommandBars/Controls Caption/Visible/Enabled mapping. Three original Report_Open
procedures execute explicitly; one-page temporary PDFs are verified. Full
report pagination and Access print parity remain open
(PRINT_PREVIEW_ADAPTER.md).

Completed report-event package: navigator module binding, automatic Open/Close,
NoData versus failed-source separation, cancellation and distinct Reports collection.
Original data inventory now measures six of nine report sources bound to 54 non-NULL
controls; referenced filter forms and other gaps remain (REPORT_EVENTS.md).

Completed bounded date-format package: shared native SQL/VBA Format/Format$,
Gregorian masks, explicit locale/NULL/error behavior, DST-neutral timestamps and
real values from ten original date columns. Subsequent concatenation support resolves those dependencies (DATE_FORMAT.md).

Completed SQL concatenation package: arithmetic/Boolean/clause precedence, NULL
and dynamic ANSI-89 LIKE conversion. Original query execution increases to 49/100,
and report binding to 6/9 sources / 54 non-NULL controls (SQL_CONCATENATION.md).

See DEVELOPMENT_CHECKLIST.md for all requirements and DAO_RUNTIME.md and
MDB_IMPORT_REGRESSION.md for evidence and limitations. Private syntax inventory
outputs remain under build and contain no raw VBA literals or fixture rows.

Completed SQL object-value package: project-scoped live Forms/Formulare and
Reports/Berichte control references in queries and saved views, typed Value
parameters, nine original references across five queries and six original Value
parameter declarations (SQL_OBJECT_VALUES.md). Complete report workflows remain open.

Completed object-reference context package: Set and assignment owners retain
actual control/DAO Field references; one unchanged original Form_Load executes
six Qt property assignments. Bare Date/Now/Time calls and catchable argument
errors are tested. Full default-member/Variant semantics remain open.

Completed project command-bar package: bare CommandBars and actual enclosing
Qt toolbar/button collections, guarded lifetimes and one-based identity tests.
Subsequent embedded-form work completes its Form_Load; the entire startup still
has unresolved calculations/activation; subsequent menu-bar, mixed-Variant,
SaveAsText and domain-aggregate work completes Form_Open and Form_Current.

Completed embedded-form execution package: real child forms, persistence and
close cancellation, typed combo columns, simple calculated aggregates, Error
Variants and complete SaveAsText string fragments (NATIVE_FORM_RUNTIME.md).
Qualified projection labels, ranked TOP dependency views, guarded TOP 1 boundary
queries and bounded two-table RIGHT-to-LEFT JOIN adaptation and subsequent origin/derived-table work execute 74/100
original queries. Nested SQL budgets remain effective. Overall compatibility,
classes/AST, remaining startup events and exact monetary types remain open.

Earlier local validation at that stage: 18/18 regressions pass, the Qt-free runtime build/test
passes, and the Debian package payload/permissions check passes. The original
29-table comparison also passes five consecutive times with MDBTools 1.0.0.
All four original fixture SHA-256 checks remain unchanged. This does not explain
the historical intermittent EOF failures or verify remote distro CI execution.

Completed active-menu package: actual project QMenuBar/QMenu/QAction objects,
popup controls and guarded lifetimes; the unchanged original complex Form_Open
completes. Subsequent domain-aggregate work reduces seven calculated Error Variants to
one and completes OnCurrent. Subsequent native activation work also completes
the original OnActivate; this is not a complete workflow claim.

Completed domain-aggregate/SaveAsText package: DSum and DCount on actual tables
and query views, optional criteria, Null/count handling and catchable errors.
Backslash-escaped property quotes are decoded correctly. Two unchanged original
DSum controls match independent record sums. Original Form_Current completes;
one of 34 calculated controls still produces an initial Null conversion error
after subsequent CCur implementation. Exact Jet Currency storage remains open.

Completed bounded Currency execution package: exact four-place values and decimal
string conversion, checked Currency arithmetic/comparison, subtype introspection
and unbound control retention. The original CCur expression executes with a
synthetic non-Null input; its initial Null branch remains error 94 under eager
IIf. Full mixed coercion/locales and exact Jet MONEY storage remain open.
See CURRENCY_RUNTIME.md.

Completed native activation package: actual Qt document events, deferred first
Current, form-to-form Deactivate/Activate, accepted/cancelled close and embedded
form exclusions. The unchanged original Open/Load/Activate/Current all complete.
Remaining initial calculated error 94, full workflows and Access event-oracle
comparison remain open (NATIVE_ACTIVATION.md).

Completed bounded query-origin package: Jet bracketed derived tables, all-field
DISTINCTROW semantics, proven unique base-column lineage and pure previous-alias
expansion. Corpus execution rises from 58 to 74 (69 ordinary and five positive
parameter cases). One original reference requests a field absent from its
six-column source view; it remains explicitly unresolved (QUERY_ORIGINS.md).

Completed DAO parameter-collection package: real Count/Item/For Each, shared
Parameter identity and explicit stale/closed handle errors (3420). SQL replacement
keeps the collection current while invalidating previous parameter generations.
Original six-query declarations and five positive parameter results are tested.

Package validation optionally compares the packaged executable's GNU build ID
with the current linked binary. Both Debian/Ubuntu CI package checks require it.
An actually stale local package was rejected and the fresh package passed, in
addition to existing permissions/content/fixture-exclusion checks.

Completed native filter package: real Filter/FilterOn properties and UI, saved
FilterOnLoad, typed edits, readonly query filters and recovery/cancellation.
Twelve bound original textbox values are checked after filtering actual records.
Me.Requery reloads external DB changes and Current is dispatched for bound
sources. The unbound original Current/Requery cycle is prevented by respecting
its absent record source (NATIVE_FILTERS.md). Full advanced semantics remain open.

Completed import version/library diagnosis: byte-identical pre-fix sources copy
9/29 when the backend library is 1.0.1, including with CLI-reported 1.0.0, and
29/29 with library 1.0.0. Twenty standalone PK clauses explain the parser failures.
The current full 29-table schema/data test passes all four combinations. EOF was
not reproduced and stays explicitly unresolved (MDB_IMPORT_REGRESSION.md).


Completed native MsgBox package: actual selected buttons return VBA results,
optional positional arguments and styles are validated, and an unchanged original
missing-input handler displays/focuses real native controls (MESSAGE_BOX_RUNTIME.md).
Completed native Close/Dropdown package: original Close button execution,
Unload cancellation, real SQL-backed popup rows and typed mouse selection
(CLOSE_COMMAND_RUNTIME.md).

Completed conditional alias package: pure IIf projection chains and leading-dot
numeric literals; three further unchanged queries return positive non-NULL results.
Current query coverage is 77/100, without an Access result oracle (QUERY_ORIGINS.md).
A reproduced nested Currency ByRef mutation loss is fixed, including original
expression value interoperability and error/alias regressions (CURRENCY_RUNTIME.md).

Completed bounded module-state package: procedure locals, CodeBehind instances
and connection-owned standard modules have separate lifetimes and lookup scopes.
Original public declarations provide actual shared state; full classes/types remain
open (MODULE_STATE_RUNTIME.md).

Completed original calendar package: both unchanged function-expression date-field
bindings execute the original Screen/global helper, native calendar page handlers,
Click/Error/Resume path and DoCmd.Close. Both selected values feed a query of actual
MDB records with an independent positive count. This uses a limited MSCAL-to-Qt
adapter, not general COM/ActiveX execution. The original helper exposed and regressed
a string-literal-versus-syntax lexer defect. Full generic VBA/default-property
semantics and OCX parity remain open (CALENDAR_RUNTIME.md). Latest syntax inventory:
182 accepted / 187 inventoried, explicitly not runtime execution coverage.


Completed bounded OpenForm package: four positional arguments, omitted defaults,
acNormal and WhereCondition load actual restricted records before Open. Both direct
DoCmd and Application.DoCmd dispatch are tested. Missing targets, initialization
errors and cancellation propagate to VBA. An unchanged original wizard handler now
opens its original target. Its initially failing target exposed a missing native
OptionGroup value (2465); numeric defaults and exclusive native choices fix that
regression. Separate original-layout/data and synthetic interaction tests verify
filtering, cells, option selection, events and persisted SQL values. This remains
modal preview infrastructure, not full Access window/view/filter-name semantics
(OPEN_FORM_RUNTIME.md).


Completed bounded Access bracket/UTF-8 identifier package: native bracket names
read typed current-form values through a separate context, preserve calculated
control read-only behavior and do not resolve same-named local variables. The
unchanged original bracket condition executes in an isolated original-layout
probe; its full surrounding handler is not executed by that probe. A reproduced
DAO regression fixes inconsistent ASCII-only statement rules. An unchanged
original Unicode-field assignment copies actual imported values and Null through
native DAO into disposable tables, with independent SQL assertions. Current
syntax inventory: 182/187 accepted, not execution coverage
(BRACKET_IDENTIFIER_RUNTIME.md).


Completed bounded TableDefs package: actual user-table Count/Item/For Each,
case-insensitive named lookup, cached native object identity, TableDef.Name and
real dynaset/snapshot OpenRecordset. The unchanged original standard-module
existence function executes successfully for all 29 original tables and returns
a typed False for an absent name; all 29 record counts match independent SQL.
This exposed and fixed single-line If/colon Exit For losing its enclosing loop
context, plus incorrect Null function-result defaults. Fields/Indexes/Relations
metadata and complete DAO catalog semantics remain open (TABLEDEFS_RUNTIME.md).


Native With/SendKeys increment (WITH_KEYBOARD_RUNTIME.md): three unchanged
keyboard handlers now complete; actual queued receivers/text are checked.
Generated dotted aliases, displayed calculated values, local SELECT INTO and
per-module Option Compare fixed reproduced runtime failures along that path.
The independent original make-table probe has 14 columns and zero sale rows;
positive/NULL copies are separately tested. Two subsequent asynchronous
AfterUpdate events remain expected 3075/1023 failures with no active-sale ID.
This is not a completed sale regression. Syntax inventory is now 185/187,
explicitly separate from execution coverage. Optional REA CLI is installed outside
the app; Ghidra/Windows Access reference execution is unavailable.
Validation: Release build and 18/18 CTests (71.91 s); GUI-disabled core 1/1
(3.00 s); fresh .deb payload/Build ID check; four original SHA-256 checks.


Native modeless OpenForm continuation: reproduced failing immediate Forms access,
replaced synchronous preview exec with a persistent native window, reused unique
same-project instances and preserved canonical catalog Name. Imported Modal
blocks UI without suspending VBA. Immediate Screen access now follows logical
native activation; its pre-fix regression failed. Real SQL persistence, caller
close cancellation, guarded references and self-close then MsgBox are tested.
Native deletion waits for the outermost VBA call to unwind. Original bound form,
wizard and both calendar selections remain tested over unchanged artifacts.
Caller-tree ownership, independent application lifetime/MDI, reentrant opening,
complete active-document DoCmd/keyboard routing and full window options remain
open. This does not complete the original sale workflow.
Validation: all 18 Release CTests (75.58 s), fresh .deb payload/Build ID and all
four original hashes passed. The GUI-disabled runtime code is unchanged.


Bounded active-document command increment (ACTIVE_DOCUMENT_COMMANDS.md): the
pre-fix SQL persistence assertion failed after OpenForm. Global GoToControl,
save/navigation/Maximize/Requery now use the invoking project's active document;
explicit Me remains bound to its instance. Application.DoCmd paths and actual SQL
actions are tested, including Null/default-Null/Error/overflow rejection. Queued
keys follow a subsequent window switch and independently opened same-project
windows, without saving another project's dirty record. Original-layout real-MDB
focus/Last/Requery cells are independently compared in the existing explicitly
CodeBehind-free binding probe. Full sale setup, remaining commands/subform
semantics, Access warnings/transaction flags and independent document lifetime
remain open.
Validation: 18/18 Release CTests (71.98 s), GUI-disabled core 1/1 (2.85 s), fresh
.deb payload/Build ID and all four original hashes passed.


Independent-document / Quit / dynamic-combo increment (31 October 2026): the
old uncommitted workspace changes and Quit logs were absent. Development resumed
from the remotely verified private branch at 753ffc8; no missing feature was
assumed present. OpenForm targets now belong to their project window rather than
the invoking form, retaining actual command bars. The unchanged original opener
and its target events complete; closing/destroying the caller leaves the target
open. The first new original test reproduced error 5 when a parentless target
lost ActiveMenuBar; project ownership fixes that specific failure.

Project switch/window close and application Quit now validate native Access
forms/reports, save pending records and honor Unload cancellation before logical
closure. Tests cover retained references, retry after cancellation, cross-project
isolation, reentrant shutdown rejection, exactly one finished notification and
deferred deletion. Successful Quit is exercised inside a real QApplication event
loop, which must terminate; code following statement Quit must not execute.
See DOCUMENT_LIFETIME_RUNTIME.md for limits, including other native editors.

Dynamic RowSource, RowSourceType, BoundColumn, ColumnCount and ColumnWidths now
operate on typed multicolumn Qt models with independent bound/display values,
hidden popup columns and actual row selection/persistence. Invalid property/query
changes preserve prior metadata/rows/selection. See DYNAMIC_COMBO_RUNTIME.md.

Validation: final Release build; 18/18 CTests passed (48.26 s), GUI-disabled
Release core build and 1/1 test passed (0.39 s), fresh Debian package matched the
current executable Build ID and passed content/permissions/fixture-exclusion
checks. All four original fixture SHA-256 checks passed before and after testing.
No source/fixture literals, rows or credentials were added to public artifacts.
Remote distro CI remains separately tracked, not inferred from these local tests.
The full initialized sale/snapshot/report workflow, exact Access event/window
parity and previously documented two empty-sale AfterUpdate errors remain open.

Remote follow-up: code commit 121afa58c46c55b2cb623ebbbbeec2c7b81b4dc2 was pushed
to development/access-compatibility and verified with ls-remote. Repository
visibility was rechecked as private. GitHub Actions run
https://github.com/BlueDragon4251/LinAccess/actions/runs/41250623462
failed before running job steps. The Ubuntu and Debian 12 check annotations
explicitly report account payment/spending-limit restrictions preventing job
startup. This is an external runner/billing blocker, not a compiler/test failure;
Debian 12/13, Ubuntu 24.04 and Windows PowerShell CI validation remain unperformed.
Account/billing settings were not changed. Re-run the private workflow after the
account restriction is resolved, then continue the positive initialized-sale /
snapshot / report workflow. No fixture upload or external analysis is required.


Native keyboard/events increment: the pre-fix Qt regression reproduced Escape
closing a form. Native Escape now restores the active field, then pending record
changes, preserving NULL, duplicate bindings and invalid numeric edits. Shift+Enter
saves with validation; F9 recalculates, Shift+F9 requeries, F2 toggles editor
selection, Enter advances focus and Ctrl+W/Ctrl+F4 close through normal lifecycle.
Form.Recalc and Form.Undo execute native operations. Report-preview Escape closes
the preview. This is a bounded subset, not Access 2003 shortcut/UI equivalence.

Bound KeyDown/KeyPress/KeyUp, KeyPreview, ByRef cancellation, ASCII replacement and
native modifier masks have actual Qt/SQL event-order tests. A child form receives
and saves its own input without routing through the parent's KeyPreview or saving
its pending record. Unsupported physical/locale key mappings fail explicitly.
Three original KeyPress procedures execute as explicit calls; only one has a
matching original binding. The two renamed controls' missing procedures remain
ProcedureMissing rather than being guessed. Full original-layout login separately
fails at its five-column combo's sixth-column read (error 5, source line 34), using
synthetic non-credential input. No password is read, serialized or reported.
See KEYBOARD_RUNTIME.md for exact scope, original source hashes and limitations.

Validation: final Release build; 18/18 CTests passed (49.12 s); GUI-disabled
Release core build/test 1/1 (0.37 s); fresh .deb matched the current executable
Build ID and passed payload/permission checks. All four original hashes passed.
Remote CI remains subject to the separately recorded account startup restriction;
these local results do not verify Debian 12/Ubuntu or Windows Access behavior.
Full sale/login/report operation, remaining VBA/DAO/Jet behavior, Access 2003
appearance/designers and exhaustive keyboard parity remain open.

Native combo/platform increment (2026-11-01): initial choice loading now happens
on the actual Qt popup gesture (mouse or Alt+Down), with typed numeric-ID search
drafts retained. Synthetic GUI regressions assert real model rows, mouse
selection, SQL persistence, event-loop responsiveness and failed-RowSource
recovery. QComboBox::setModel already deletes its previously owned model;
removing the subsequent dereference fixes the Qt 6.4 use-after-free, with a
QPointer regression on both supported Qt versions. MDBTools receives a UTF-8
child locale while the parent environment remains untouched; the new C-locale
regression reads original Unicode table schemas from a Unicode-named test copy
and rechecks the original hash. Explicit Qt value includes fix clean Qt 6.4 builds.

Clean distribution validation exposed a separate SQLite 3.45.1 engine crash in
Ubuntu 24.04 while preparing nested JOIN-view column metadata, before stepping
records. Private schema-only native reduction reproduced the crash outside Qt.
A missing-dependency preflight preserves error 3078 for bounded transitive view
references but alone does not fix the engine defect. Both original failing
regressions passed unchanged against current official SQLite 3.54.1. The GUI
build now compiles that unmodified public-domain source archive offline, verifies
its published SHA3-256, and packages libsqlite3.so.0 in a private app directory.
No system library is replaced. Qt/native/compiled source IDs must agree before
native connection access; an older library injected by a workspace search path
was explicitly rejected with error 3251. The isolated host validation prepends
the matching build backend to its existing Qt sysroot search path.

Release results with the pinned backend: host Qt 6.8.2 20/20 CTests (49.70 s),
Debian 12 Qt 6.4.2/GCC 12 container 20/20 (51.54 s), Ubuntu 24.04 Qt 6.4.2/GCC 13
container 20/20 (47.70 s); GUI-disabled Release core 1/1 (0.44 s). All three fresh
DEBs matched executable Build IDs and exact backend bytes, passed payload and
permission allowlists and ran the extracted native JOIN/NULL/backend-identity
self-test. Debian/Ubuntu packages were also installed into fresh containers;
their installed binaries resolved the private backend, passed the same self-test,
and stayed running for the bounded offscreen GUI startup check. The package
contains only the executable, desktop entry, private public-domain backend and
backend notice; no original fixtures, VBA sources or credentials. All four
original fixture SHA-256 checks passed after testing on every build platform.

The CI workflow now installs the package and verifies its backend, records final
fixture hashes and uploads complete DEBs rather than an incomplete bare binary.
Remote GitHub runner startup remains blocked by the previously recorded account
billing/spending restriction; local container success is not a remote-CI claim.
Measured query inventory remains 100 original query artifacts / 77 executions without
an Access result oracle; original full sale/login/report operation, UI/designers,
remaining VBA/DAO/Jet and exhaustive Access 2003 parity remain open.

Remote verification follow-up: private GitHub run
https://github.com/BlueDragon4251/LinAccess/actions/runs/41298553788 completed
successfully for pushed backend commit a63d003b55bbf8be7c19d161e8bf9ccd25fa553e.
Ubuntu 24.04, Debian 12, Debian 13 and Windows PowerShell syntax jobs all ran and
passed. Linux jobs include full CTest, DEB content/build identity checks, installed
backend self-test and final fixture hashes. The previous account startup block
did not recur in this run. This verifies the Linux CI matrix and bridge syntax;
it is not Windows Access runtime or rendering verification.

Event-prefix increment: native control dispatch now honors the stored
EventProcPrefix independently of Name, including supported click/double-click,
keyboard and validation events. All 47 original form layouts are scanned with
68 stored prefixes preserved; rendered control objects expose their values.
Synthetic GUI/SQL regressions verify actual prefixed writes, key order/ASCII
replacement, validation cancellation, VBA-driven reassignment, default naming,
unbound events and explicit ProcedureMissing without name guessing. Existing
original keyboard-binding/login limitations remain unchanged. FORM_EVENTS.md
records scope and unverified Access 2003 naming/design-view edges.

Validation: Release 20/20 CTests on host Qt 6.8.2 (63.84 s), Debian 12 Qt 6.4.2
(56.60 s) and Ubuntu 24.04 Qt 6.4.2 (52.94 s). Fresh DEBs passed executable/backend
identity, exact payload/permissions and extracted backend regressions. Both distro
packages passed installation, private-backend resolution and bounded offscreen
GUI startup in fresh containers. Four original hashes passed after testing.

Remote event-prefix verification: private GitHub run
https://github.com/BlueDragon4251/LinAccess/actions/runs/41299117595 for commit
b10a8b645e99fae4a5e8068fd0e36a4499ad904e completed successfully across Ubuntu 24.04,
Debian 12, Debian 13 and Windows PowerShell syntax. This includes the prefix
GUI/SQL regressions in Linux jobs; no Microsoft Access runtime was involved.

InputBox increment: both GUI hosts delegate InputBox/InputBox$ to a shared native
Qt dialog; the app generates no answer without user input. Positional optional
arguments, defaults, TWIP coordinates, text/empty cancellation return and explicit
argument/host/Windows Help errors are implemented. Actual modal GUI tests enter
and save quoted text, Escape/Cancel empty text, length-limited text, one-line
clipboard paste and Unicode IME input through an imported bound form into real
SQLite rows. Core tests check omissions, conversions and numeric error classes.
Original source inventory identifies two calls in one DAO/form-dependent procedure;
that complete original procedure is not yet verified. INPUT_BOX_RUNTIME.md records
this boundary and unimplemented named/statement/help/rendering behavior.

Validation: Release host 20/20 CTests (53.80 s), GUI-disabled core 1/1 (0.49 s),
Debian 12 20/20 (51.77 s), Ubuntu 24.04 20/20 (48.76 s). Fresh packages passed
executable/backend/payload/permissions checks and extracted backend self-tests;
both distro DEBs passed installed backend and offscreen GUI startup checks in
fresh containers. Four original fixture hashes passed after distro testing.
No original credential was used in GUI input tests or copied to artifacts/logs.

Remote InputBox verification: private GitHub run
https://github.com/BlueDragon4251/LinAccess/actions/runs/41300019379 for commit
f694c81397c4c8221035753f640ed8e7f4f811cf completed successfully on Ubuntu 24.04,
Debian 12, Debian 13 and Windows PowerShell syntax. The run verifies the InputBox
increment on Linux; it does not verify a Windows Access input-dialog oracle.

DAO cursor/search increment: clone cursors now share actual dataset changes with
independent positions/edit buffers, typed opaque byte-array bookmarks, bounded
Find methods and NoMatch. Stale/foreign/deleted/reused-rowid bookmark rejection,
strict unknown-field handling, ordered query snapshots and independent close
have native SQL regressions. Form.Recordset exposes a restricted clone/search
view; Form.Bookmark selects through native save/validation/Current processing.
Saved table values refresh shared clone snapshots without losing surviving row
bookmarks; explicit requery/filter reload invalidates the form family. Str/Str$
implements the numeric SQL-fragment conversion used by the original generated
handlers. General arrays, editable form Recordset, locking and complete criteria/
formatting/provider semantics remain open; see DAO_CURSOR_RUNTIME.md.

The unchanged original module and original layout (source hash in that document)
run against a prepared synthetic table and prepared real Qt lookup models. Both
original automatic Open/Load procedures complete. Two original generated search
procedures pass positive/no-match/Null cases (six explicit calls) and select the
expected record through actual combo mouse input (two bound event executions).
A third original procedure references an absent exported control, including after
case/space normalization. Its direct call and actual bound combo event fail with
2465 without navigation. No renamed alias is invented. This is not a successful
third search or verification of the complete original lookup/business dataset.

Validation: Release host 20/20 CTests (68.29 s), Qt-free core 1/1 (0.46 s), Debian 12
20/20 (54.37 s), Ubuntu 24.04 20/20 (50.86 s). Fresh DEBs passed executable/backend
identity, exact payload/permission allowlists and extracted backend regressions.
Both distro packages passed installed-backend resolution and bounded offscreen
GUI startup in fresh containers. Four original fixture hashes passed after tests.

Remote cursor/search verification: private GitHub run
https://github.com/BlueDragon4251/LinAccess/actions/runs/41303749658 for commit
f523639a73b18de4d3625700599c244f86520cd7 passed Ubuntu 24.04, Debian 12, Debian 13
and Windows PowerShell syntax. The source's missing third search control remains
an explicit expected error, not a counted successful original workflow.

Built-in row-source increment: Value List/Wertliste now populates actual
multicolumn Qt rows from semicolon-separated text with quoted delimiters, doubled
quotes and Unicode. Actual native popup selection persists its bound integer
value in SQLite. Failed parses/row-limit violations preserve the prior source,
model and value. Empty/blank/padding semantics are explicitly documented as the
current implementation, pending an Access 2003 edge-case oracle. Field List/
Feldliste lists table or SELECT projection names without evaluating/fetching its
records; declared parameters do not require values for description. Shared query
preparation/error handling preserves ordinary result execution and suppresses
internal TOP rank columns in metadata. Tests describe a SELECT with an eager
runtime overflow expression and independently confirm actual execution still
fails, proving metadata extraction does not execute its data expressions.

The original fixtures contain 17 explicit Table/Query declarations and no explicit
Value/Field List declaration; these general native GUI/schema regressions imply
no additional original-project success. Custom row-source callback functions,
ColumnHeads, AddItem/RemoveItem and full ListBox behavior remain open.

Validation: host Release 20/20 CTests (72.75 s); GUI-disabled core 1/1 (0.37 s);
Debian 12 20/20 (55.40 s); Ubuntu 24.04 20/20 (52.40 s). Fresh packages passed exact
executable/backend/payload/permissions and extracted backend checks. Both distro
packages passed installation, private-backend resolution and bounded offscreen
GUI startup in fresh containers. All four original fixture hashes passed after
testing. DYNAMIC_COMBO_RUNTIME.md records precise scope and reference limits.

Eval increment: native text evaluation, Application.Eval, statement invocation,
public Function calls, single-quoted strings and Access Is Null/Is Not Null.
Synthetic native UI/SQL tests verify actual effects, forbidden private/Sub bodies,
typed argument errors, recursion recovery and modal self-close lifetime. The
unchanged original module identified in EVAL_RUNTIME.md completes in 20 prepared
numeric/Null states; counted host accesses verify each original Null branch's
additional Forms lookup. This is not a full-project or Access-oracle comparison.

Validation: host Release 20/20 CTests (56.80 s); GUI-disabled core 1/1 (0.58 s);
Debian 12 20/20 (58.45 s); Ubuntu 24.04 20/20 (54.20 s). Fresh DEBs pass payload,
permissions, executable/backend identity and extracted SQL-backend checks. Both
distro packages also passed isolated installation, private backend resolution and
bounded GUI startup. All four original fixture SHA-256 checks pass. Full Eval
grammar, Access name resolution/coercion and complete original workflows remain
open; original source and data stay in private fixtures.

ComboBox ListCount increment: a read-only property now loads an uninitialized
real row source and returns the native model's row count. Tests cover Value List,
SQL rows (including a NULL-valued row), empty lists, explicit Requery/cache
behavior, unchanged selected value/Dirty state, invalid arguments/assignment and
a real VBA procedure inserting its result into SQLite. Both original ListCount
tokens occur in comments, so no additional original caller execution is claimed.

Validation: host Release 20/20 (58.14 s), Debian 12 20/20 (59.86 s), Ubuntu 24.04
20/20 (55.18 s); freshly generated host/distro packages pass exact payload/backend
checks. Distro packages pass isolated installation and bounded offscreen startup.
All four original fixture hashes pass. Core-only source is unchanged from the
preceding verified build. ColumnHeads/full ListBox and Access 2003 list-count
edge semantics remain open; see DYNAMIC_COMBO_RUNTIME.md.

Single-record command increment: native RunCommand SelectRecord/DeleteRecord and
bounded Access-95 DoMenuItem Select/Delete compatibility execute against the
active document. Real highlighted selection, read-only SelHeight, AllowDeletions,
pending-edit validation, bound Delete/BeforeDelConfirm/AfterDelConfirm and actual
confirmation run through savepoint-backed deletion/rollback. Tests verify pending
SQL visibility, programmatic/user cancellation, invalid/Null response, surviving
and stale bookmarks, FK protection, empty/new/read-only targets, guarded callback
edits, nested navigation/deletion/close and separate caller/target contexts.

The unchanged original deletion procedure identified in RECORD_DELETE_RUNTIME.md
executes once explicitly and once through its actual bound button with prepared
synthetic rows and native confirmation. No original source/data/credentials are
printed or packaged. All 187 original procedures are now syntax-accepted after
omitted DoMenuItem argument support; this remains syntax-only inventory. The
original legacy Copy/PasteAppend procedure still fails unsupported operations;
full sale workflows, complete writable sources and Access event/UI parity remain
open.

Validation: host Release 20/20 (59.78 s); GUI-disabled core 1/1 (0.59 s); Debian 12
20/20 (61.66 s); Ubuntu 24.04 20/20 (56.76 s). Fresh host/distro DEBs pass exact
payload, permissions, executable and backend checks. Both distro packages pass
isolated installation, backend resolution and bounded GUI startup. All four
original fixture hashes remain unchanged. New record command regressions are in
record_delete_tests.cpp and the original search/deletion regression.

AutoNumber import increment: preserve MDBTools' explicit counter modifiers through
structural metadata rather than discarding them. All 21 original flags survive;
16 single INTEGER primary keys use native AUTOINCREMENT and five non-primary-key
counters are explicitly metadata-only. Imported IDs remain unchanged. Tests
cover explicit values, next-ID generation, deleted-highest-key no-reuse, higher
explicit IDs, quoting/malformed metadata and the unsupported non-PK case. One
eligible copied original table is probed with insert/delete/insert under a rolled
back savepoint; all 29 copied tables still match their exported cells.

Validation: host Release 20/20 (62.03 s), Debian 12 20/20 (60.45 s), Ubuntu 24.04
20/20 (54.23 s). Fresh DEBs pass exact payload/backend checks, isolated distro
installation and bounded GUI startup. All four fixture hashes pass; GUI-disabled
core source is unchanged from the preceding verified build. The source's stored
counter seed, five non-PK generators, random modes and exact Jet allocation/
overflow semantics remain open; existing projects are not silently rebuilt.
AUTONUMBER_IMPORT.md records the precise new-import scope.

Single-record clipboard increment: real system clipboard, typed native MIME plus
quoted TSV, Copy/PasteAppend and selected-record Ctrl+C. Tests preserve Unicode,
multiline text, NULL/empty distinction, binary data and an exact 64-bit integer
through actual SQL insertion. Verified AUTOINCREMENT columns generate destination
keys; manual keys are retained. Malformed/oversized/ambiguous data, incompatible
shapes, disabled additions, unsupported event/counter contexts and excluded form
filters fail without inserting partial records. Active-target scope and saved
snapshot bookmarks are tested. Native AllowAdditions also controls new-record
navigation/selector availability. Deletion SQL now has a budget limited to SQL
execution, leaving modal user wait outside that budget.

The original duplicate handler's actual 14-column source schema has a manual PK,
no AutoNumber. Direct append reports 3022; one unchanged original invocation and
one real bound click run its native error-message path with the row count still
one. Completed handler status is not recorded as successful duplication. Full
insertion event bindings, multi-record/external clipboard formats and exact Access
conversion/rendering remain open (RECORD_CLIPBOARD_RUNTIME.md).

Validation: host Release 20/20 (76.85 s), GUI-disabled core 1/1 (0.66 s), Debian 12
20/20 (62.56 s), Ubuntu 24.04 20/20 (57.43 s). Fresh DEBs pass exact payload,
permissions and backend checks; isolated installed distro packages pass backend
resolution and bounded offscreen startup. All four original fixture hashes pass.
Original fixtures remain read-only; all mutation probes use synthetic data/copies.

Managed counter increment: the five non-primary-key counters now have persistent
DEFAULT-function allocation with transactional sequence state and explicit-insert/
update seed triggers. All 21 original counters have bounded native generation on
new imports (16 AUTOINCREMENT, five managed); existing projects are not rebuilt.
Form save and DAO AddNew omit unset generated fields, and PasteAppend verifies
actual generated defaults. Counter values remain separate from rowid/bookmarks.
Tests cover constraint/import rollback, higher explicit values, no-reuse, reopen,
quoted identifiers, invalid/exhausted sequences and two simultaneous connections
allocating 40 distinct values. Native form/DAO/clipboard and counter-only forms
are exercised. The original 29-table copy still compares every exported cell;
separate isolated generation probes cover original single-PK and non-PK tables.

The backend CLI now performs a fixture-free managed-counter regression, so both
extracted and installed DEBs execute the actual callback/allocation path. All
20 host tests pass (63.70 s); the added CLI probe passes separately (0.04 s).
GUI-disabled core: 1/1 (0.62 s). Final Debian 12 and Ubuntu 24.04 suites pass
20/20 (63.20 s / 59.29 s), including the expanded backend CLI. Fresh DEBs pass
payload, permissions, binary/backend identity, isolated installed allocation and
GUI startup checks. All four original fixture hashes pass.

Exact source seeds, random/replication modes, signed-Long overflow, allocation
timing/gaps, full read-only/DAO Attributes, schema rename/drop cleanup and direct
SQL explicit-NULL allocation remain open. AUTONUMBER_IMPORT.md records the scope;
this is not complete Jet counter or full Access 2003 parity.

Joined UPDATE increment: one INNER JOIN can update either direct physical table
through DAO/QueryDef.Execute. Materialization preserves self-join inputs; a native
typed aggregate rejects differing candidates with 3251 and statement rollback.
Tests cover arithmetic, parameters, NULL/text/binary values, identical duplicates,
ambiguous duplicates, frozen self-joins, aliases, empty matches and unsupported
views/outer joins/multiple targets/row identities. The new joined_update_contracts
CTest brings the fixture-enabled suite to 21 tests.

One unchanged original action is additionally executed against actual imported
schemas with synthetic unique matches: one expected target value changes. With
full original copied data its candidates conflict; the second action requires an
outer join and saved-query target. Both remain explicitly bounded errors, so the
full original corpus remains 77/100 without an Access result oracle. No source
query, fixture rows or VBA text were added to documentation or logs.

Validation: host Release 21/21 CTests (63.39 s), GUI-disabled core 1/1
(0.54 s), Debian 12 21/21 (58.71 s), Ubuntu 24.04 21/21 (54.62 s).
Fresh host/distro DEBs pass payload/backend/permission/fixture-exclusion checks;
all four original fixture hashes pass.
Isolated installed Debian/Ubuntu packages also pass allocation/backend and GUI
startup checks without source-tree or fixture mounts.

Active navigation increment: DoCmd.GoToRecord now supports all six AcRecord
constants for the active native form, positive next/previous offsets, absolute
saved-record numbers and explicit acActiveDataObject. Selection follows actual
save/validation, including a dirty newly saved row; cancellation raises 2501.
Out-of-range positions raise 2105 rather than silently clamping. Next from last
opens the new row, and Previous returns to the last saved row. Tests exercise
statement/object dispatch, offsets in both directions, new-row boundaries and
actual SQL persistence with BeforeUpdate cancellation. Named-object navigation
remains explicitly unsupported; MicrosoftDocs' GoToRecord/AcRecord contracts
were checked without using private fixture content in external requests.

Validation: host Release 21/21 (63.16 s), GUI-disabled core 1/1 (0.76 s),
Debian 12 21/21 (59.74 s), Ubuntu 24.04 21/21 (55.82 s). Fresh DEBs
pass exact payload/backend/permission/fixture-exclusion checks and all four
original fixture hashes pass.
Isolated installed distro packages pass the native backend/counter and GUI startup checks.

Named-form navigation increment: GoToRecord acDataForm now resolves a live
case-insensitive same-project form name, including hidden windows, without
activating/reopening that form. Tests execute statement and Application.DoCmd
calls while a separate controller remains active; actual displayed records
change and Me stays with the controller. Closed/foreign forms raise 2450 and
ambiguous instances raise 5. Native table/query/report targets remain unsupported.
The targeted access_2003_layouts regression passes (6.64 s).

Validation: host Release 21/21 (61.45 s), GUI-disabled core 1/1 (0.55 s),
Debian 12 21/21 (60.21 s), Ubuntu 24.04 21/21 (55.68 s). Fresh host/distro
DEBs pass payload/backend/permission/fixture-exclusion checks; all four original
fixture hashes pass. ACTIVE_DOCUMENT_COMMANDS.md now reflects the implemented
record actions/offsets and independent document ownership.
Isolated installed Debian/Ubuntu packages pass backend/counter execution and GUI startup.

Control Requery increment: DoCmd.Requery ControlName now resolves the active
form's named supported control. It reuses actual row-source/subform/calculation
adapters; bound non-list data controls reload the form after normal save
validation. Tests verify externally changed SQL rows, separate caller/target
combo models, instance-specific Me.pick.Requery and BeforeUpdate cancellation
without losing dirty input or writing rejected values. Targeted layouts pass
(6.60 s). Full Access focus/event/reload parity is not claimed.

Validation: host Release 21/21 (57.89 s), Debian 12 21/21 (58.76 s),
Ubuntu 24.04 21/21 (54.01 s). Fresh host/distro DEBs pass exact payload,
backend, permissions and fixture-exclusion checks; all original fixture hashes
pass. GUI-disabled core remains unchanged from the preceding 1/1 validated
commit. QUERY_ORIGINS.md now states the current 77-query coverage and joined
action boundaries directly instead of retaining the earlier 74-query summary.
Installed Debian/Ubuntu packages also pass native backend/counter execution and GUI startup.

Insertion events increment: native control changes on a new direct-table form
record now dispatch BeforeInsert before Dirty, and a successful actual INSERT
dispatches AfterUpdate then AfterInsert. VBA assignment does not fire insertion
events. Native tests cover text/checkbox input, exact traces, actual SQL values,
field initialization from the handler, Cancel/errors/Undo and validation/UNIQUE
retry. An additional regression exposed and fixed Undo being ignored during
BeforeInsert while Dirty was still false. Rejected activations do not recommit
values after cancellation. Targeted layouts pass (7.24 s).

This runs after the editor value mutation; it does not establish Access's
BeforeInsert/Change/Text/Value visibility. INSERT_EVENT_RUNTIME.md states that
remaining timing/oracle boundary and the separate unsupported PasteAppend
insertion-event pipeline. Full event parity is not claimed.

Validation: host Release 21/21 (74.06 s), Debian 12 21/21 (60.50 s),
Ubuntu 24.04 21/21 (54.69 s). Fresh host/distro DEBs pass executable/backend,
permissions, exact payload and fixture-exclusion checks. All four original hashes
pass. The GUI-disabled core has no source changes in this increment.
Isolated installed Debian/Ubuntu packages pass backend/counter execution and GUI startup.

Change event increment: imported text/combo OnChange bindings now dispatch
through EventProcPrefix on native edits. A new native Qt/SQL regression checks
KeyDown/KeyPress/BeforeInsert/Change/KeyUp, existing-record edits, combo editing,
KeyPress cancellation, handler effects/error propagation and actual saved values.
VBA Value assignment, bound-mirror synchronization and loading/requery do not
fire Change. The targeted layout suite passes (6.84 s). Text/Value buffering,
Text-property setters, cascading Change and LimitToList/NotInList remain open.

Validation: host Release 21/21 (73.78 s), Debian 12 21/21 (59.76 s),
Ubuntu 24.04 21/21 (56.00 s). Fresh host/distro packages pass exact payload,
backend, permissions and fixture-exclusion checks. All four original fixture
hashes pass. GUI-disabled core sources are unchanged from their preceding
validated 1/1 build.
Isolated installed Debian/Ubuntu packages pass native backend/counter execution and GUI startup.

Text selection increment: native focused text/ editable-combo wrappers expose
SelStart/SelLength/SelText. Native VBA tests replace selected text, check actual
Change callbacks and SQL values, focus/type/Null/range failures, selection-only
clean state and UTF-16 surrogate-pair replacement. SelText without selection is
Null, matching the published Microsoft contract. The targeted layouts pass
(6.96 s). A read-only token scan of 34 restored VBA files found no selection
property calls; this adds no original-handler execution claim. Full Text/Value
buffering and formatted/multiline/input parity remain open.

Validation: host Release 21/21 (62.66 s), Debian 12 21/21 (68.49 s),
Ubuntu 24.04 21/21 (62.92 s). Fresh host/distro DEBs pass payload/backend,
permissions and fixture-exclusion checks; all four original fixture hashes
pass. GUI-disabled core has no changes in this increment.
Isolated installed Debian/Ubuntu packages pass backend/counter execution and GUI startup.

Form timer increment: imported/runtime TimerInterval and OnTimer now use a
form-owned precise QTimer and real VBA/expression dispatch. Tests assert SQL
effects, Load configuration, three-tick self-stop, hidden restart, cancelled
Unload, successful close/no further ticks, self-close, failures and expression
bindings. Strict boolean deadline waits are asserted in the standalone runner.
The targeted layout suite passes (7.66 s).

One unchanged original Kassa clock module executes its original Timer binding
on a prepared layout with unrelated lifecycle bindings removed and a shortened
interval. Both Now assignments replace staged old values in the form caption and native
field and match independent timestamp bounds. Hashes protect the original module/layout; all mutations are
inside an isolated project/savepoint. This is one procedure (the SaveAsText
occurrence duplicates the VBA export), not full startup or Access display-locale
parity. TIMER_RUNTIME.md records the guards and remaining scheduling/oracle gaps.

The first full timer run exposed a bug in the new original probe: Me.Caption
was read as a control rather than a form property. The probe is corrected and
now stages old timestamps before checking both actual changes. Targeted layouts
and real Kassa integration pass together (32.03 s). No runtime timer defect was
masked or bypassed; the original handler remains unchanged.

Final timer validation: host Release 21/21 (60.25 s), Debian 12 21/21
(63.89 s), Ubuntu 24.04 21/21 (58.98 s). Fresh host/distro DEBs pass exact
payload/backend/permission/fixture-exclusion checks and all four original hashes
pass. GUI-disabled core sources are unchanged in this increment.
Isolated installed Debian/Ubuntu packages pass native backend/counter execution and GUI startup.

Outer joined UPDATE increment: single LEFT/RIGHT [OUTER] JOIN updates now accept
the preserved direct-table target. Native tests verify missing-source NULLs,
Is Null filtering, actual affected counts, required-field rollback (3314),
ambiguity rollback (3251) and rejection of nullable-side/FULL targets. The
materialized candidate source and existing typed ambiguity guard are preserved.
Targeted joined contracts/query inventory pass (6.38 s). Corpus coverage remains
77/100: the original remaining RIGHT action targets its nullable side and a
saved query. JOINED_UPDATE_RUNTIME.md records the exact boundary; no general
Jet updateability/result-oracle parity is claimed.

Validation: host Release 21/21 (58.81 s), Debian 12 21/21 (60.34 s),
Ubuntu 24.04 21/21 (56.05 s). Fresh host/distro DEBs pass exact payload,
backend, permissions and fixture-exclusion checks; installed distro packages
pass backend/counter execution and GUI startup. All four original hashes pass.
GUI-disabled core sources are unchanged.

OpenReport increment: explicit native preview supports omitted arguments,
WhereCondition, imported report CodeBehind, actual data binding and independent
lifetime. Tests check SQL event effects, Open/NoData cancellation, missing-source
3078, handler errors, wrong-kind/missing-report 2103, reuse and caller closure.
Default/direct printing is explicitly rejected rather than opening preview.
Other view/filter/window/OpenArgs and pagination gaps remain documented in
OPEN_REPORT_RUNTIME.md.

One unchanged original preview handler also opens its original report via direct
call and a prepared bound-button click over actual imported MDB data. Native
fields/Nulls match a separate prepared reader. The old corpus setup omitted
reports from its artifact catalog; it now imports them like the app. The original
full-layout login and other known workflow boundaries are not bypassed. Targeted
layouts/Kassa integration pass together (34.55 s); GUI-free core passes 1/1
(0.55 s). A token inventory found four preview and four direct-print call sites;
the latter are not implemented/claimed by this preview increment.

Requested theoretical status bars use the unweighted checked/total items in
DEVELOPMENT_CHECKLIST.md as a catalog indicator. Tasks differ greatly in size;
this is not a measured percentage of Access functions, behavior or visual parity.

Final OpenReport validation: host Release 21/21 (81.65 s), Debian 12
21/21 (62.07 s), Ubuntu 24.04 21/21 (56.52 s), GUI-disabled core 1/1
(0.55 s). Fresh host/distro DEBs pass exact payload/backend, permissions and
fixture-exclusion checks; isolated installed distro packages pass backend/
counter execution and GUI startup. All four original hashes pass.

String-unit increment: valid Unicode Len/Left/Right/Mid now index UTF-16 rather
than UTF-8 bytes. Dollar substring variants, Null distinctions, bounded rounded
Long arguments and strict malformed-text errors have native tests. Numeric
Null conversion now returns 94; leading-dot decimals and E/e literals are
recognized with malformed/overflow diagnostics. A shared counter also implements
SQL Len (including embedded NUL and pure alias expansion).

Tests execute expressions and an actual VBA procedure with On Error/Err.Number,
verify single scalar getter evaluation and compare SQL values for Unicode/Null/
embedded NUL/invalid data. Targeted core/query tests pass (0.50 s); the original
query inventory remains 77/100. No direct Len/Left/Right/Mid calls were found in
the restored original VBA/query token scan, so no original-handler count is added.
UTF16_STRINGS.md documents isolated-surrogate, typed Len, code-page and locale
limits; this is not full BSTR/VBA string parity.

Validation: host Release 21/21 (78.99 s), Debian 12 21/21 (61.43 s),
Ubuntu 24.04 21/21 (56.57 s), fresh GUI-disabled core 1/1 (0.56 s).
Fresh host/distro DEBs pass exact payload/backend/permission/fixture-exclusion
checks; all four original fixture hashes pass.
Isolated installed Debian/Ubuntu packages pass native backend/counter execution and GUI startup.

Trim-family increment: native Trim no longer removes tab/newline data or turns
Null into an empty String. Trim/LTrim/RTrim preserve Variant Null, their dollar
variants raise 94, and all remove only end spaces as requested. Tests assert
actual VBA assignment/error handling, other-whitespace preservation, scalar
default getter effects and arity. GUI-disabled core passes 1/1 (0.46 s). A
read-only original token scan found no direct Trim-family calls; original
execution counts remain unchanged.

Validation: host Release 21/21 (62.94 s), Debian 12 21/21 (66.44 s),
Ubuntu 24.04 21/21 (61.70 s), GUI-disabled core 1/1 (0.46 s). Fresh
host/distro DEBs pass exact payload/backend, permissions and fixture-exclusion
checks. All four original fixture hashes pass.
Isolated installed Debian/Ubuntu packages pass native backend/counter execution and GUI startup.

String-search/comparison increment: InStr/StrComp now honor Option Compare and
explicit comparison constants, UTF-16 positions/binary order, Null and bounded
start/error semantics. Numeric type mismatch reports 13. Core tests include
actual VBA execution and supplementary-character ordering. Full Windows/Jet
linguistic matching, variable-length collation and LCIDs remain unsupported.
GUI-disabled core passes 1/1 (0.49 s).
Validation: host 21/21 (61.46 s), Debian 12 21/21 (63.75 s), Ubuntu 24.04
21/21 (58.66 s). Fresh host/distro DEBs pass exact payload/backend/permission
and fixture-exclusion checks; isolated installed packages pass backend/counter
execution and native GUI startup. All four original fixture hashes pass.

Native Text/Value increment: focused Text reads/writes now use actual editors;
user typing and Text/SelText changes preserve prior Value until control update.
BeforeUpdate receives the new value, failed validation retains drafts, and
successful update refreshes bound mirrors/calculations before AfterUpdate.
Programmatic Value replaces pending input without control Change/update events.
Tests verify native VBA event values, SQL save boundaries, cancellation/retry,
Escape/Undo, new Null records, combo draft/model loading and cascading Change.
Targeted access_2003_layouts passes (7.83 s). No direct original Text-property
member tokens were found; original execution counts remain unchanged.
Validation: host 21/21 (61.56 s), Debian 12 21/21 (63.93 s), Ubuntu 24.04
21/21 (59.69 s), GUI-disabled core 1/1 (0.77 s). Fresh host/distro DEBs pass
exact payload/backend/permission/fixture-exclusion checks. Isolated installed
Debian/Ubuntu packages pass backend/counter execution and GUI startup. All four
original fixture hashes pass. A dedicated Undo/calculation regression verifies
that record reload clears draft visibility before recalculating fields.
Remaining timing/reentrancy limits are documented.

Control Undo/OldValue increment: bound OldValue reads the original field buffer
and remains unchanged through control and form AfterUpdate. Text/combo Undo
restores a pending edit without discarding other changed fields; Change and
BeforeUpdate handlers can undo without an incorrect AfterUpdate callback.
Self-modifying BeforeUpdate Value/Text/SelText raises 2115, while AfterUpdate
Value assignment is allowed. Reentrant new drafts survive cancelled save for
explicit retry. Native SQL/event/focus/new-record tests pass in
access_2003_layouts (7.62 s). Original OldValue/control-Undo token counts are zero;
no new original-handler execution count is claimed. Currency/Decimal/Date field
representation, general update reentrancy and Undo events remain separate work.
Validation: host 21/21 (76.53 s), Debian 12 21/21 (63.05 s), Ubuntu 24.04
21/21 (57.47 s), GUI-disabled core 1/1 (0.51 s). Fresh host/distro DEBs pass
exact payload/backend/permission/fixture-exclusion checks. Isolated installed
Debian/Ubuntu packages pass backend/counter execution and GUI startup. All four
original fixture hashes pass.

DAO transaction increment: native DBEngine/Application.DBEngine and the single
project default workspace/database now coordinate real BeginTrans/CommitTrans/
Rollback savepoints. Tests independently read a temporary project through a
second connection and verify Recordset.Update rollback, nested commit/outer
rollback, flags/depth limits, stale fields/bookmarks, snapshot refresh safety,
clone preflight and close/reopen generations. Typed Workspace VBA helpers execute
real SQL. A native button commits, explicitly rolls back and handles a duplicate
key with rollback before refreshing its bound control. CurrentDb.Execute now
dispatches its argument list to Database.Execute, including dbFailOnError.
Targeted core/DAO/layout tests pass (8.33 s); rebuilt GUI-disabled core passes
1/1 (0.49 s). Original API token counts are zero, so original execution counts
are unchanged. Nested rollback, clones in transactions, forced OS flush,
Workspace.Close/custom/multiple-database workspaces and full Jet cursor/locking
semantics remain open.
Validation: host 21/21 (77.13 s), Debian 12 21/21 (65.79 s), Ubuntu 24.04
21/21 (59.47 s), freshly rebuilt GUI-disabled core 1/1 (0.49 s). Fresh host/
distro DEBs pass exact payload/backend/permission/fixture-exclusion checks;
isolated installed Debian/Ubuntu packages pass backend/counter execution and
GUI startup. All four original fixture hashes pass.

Imported control-state increment: NotDefault now uses the property's actual
supported default rather than generic truthiness. Enabled/Visible persist through
reloads and the legacy visibility path updates the same model. TextBox Locked
blocks native typing while preserving VBA writes. Source read-only protection no
longer misreports Enabled=false; native editor/checkbox/combo/group input is
protected. TabStop and flat TabIndex drive real traversal and live reordering,
with renumbering and validation. Native SQL/state/focus tests and both original
layout/calendar-heavy regressions pass (34.62 s); freshly rebuilt GUI-disabled
core passes 1/1 (0.50 s). One unchanged original flat layout also passes actual
forward/backward Tab traversal without executing original VBA.

The original declared-visible command-button count is corrected from 86 to 81
(879 total parsed controls). The original calendar's exported disabled state is
now honored: its event regression explicitly enables the real control first.
This qualifies the earlier calendar workflow evidence; no unprepared OCX/Access
compatibility is inferred. CONTROL_STATE_RUNTIME.md records remaining cycle/
focus/nesting/default/Locked/visual boundaries.
Validation: host 21/21 (83.29 s), Debian 12 21/21 (66.31 s), Ubuntu 24.04
21/21 (60.89 s), freshly rebuilt GUI-disabled core 1/1 (0.50 s). Fresh host/
distro DEBs pass exact payload/backend/permission/fixture-exclusion checks;
isolated installed Debian/Ubuntu packages pass backend/counter execution and
GUI startup. All four original fixture hashes pass. The original flat-layout
keyboard probe records two forward and two backward native Tab transitions.

Password-mask increment: imported TextBox InputMask is preserved; Password uses
native immediate asterisk echo while actual Text/Value/SQL remain unchanged.
VBA reveal/remask, unsupported/Null preflight, Unicode, Null and navigation are
tested. Generic imported masks are marked unsupported and user input is blocked.
One hash-verified original layout is probed only with synthetic input in an empty
database, without original VBA/credentials/authentication; no new original
handler execution is claimed. The separate original login failure remains open.

A synthetic native layout PDF was visually checked after rasterization and shows
asterisks. The printer emits a widget bitmap, so an empty text extraction was
rejected as an unsuitable content test. GUI/original-heavy regressions pass
(35.99 s). PASSWORD_INPUT_MASK.md documents the bounded scope.
Validation: host 21/21 (81.92 s), Debian 12 21/21 (64.84 s), Ubuntu 24.04
21/21 (60.05 s), GUI-disabled build/core 1/1 (0.50 s). Fresh host/distro DEBs
pass exact payload/backend/permission/fixture-exclusion checks; isolated installed
Debian/Ubuntu packages pass backend/counter execution and GUI startup. All four
original fixture hashes pass.

Form Cycle increment: standalone flat forms now honor Cycle All Records (0) and
Current Record (1). Native Tab validates before moving focus, persists/navigates
records when required, and retains input/focus on control or record cancellation.
New rows save through real SQL without empty extra inserts; Shift-Tab returns
to an existing row. Unbound forms wrap without SQL. VBA handler focus changes
are respected. A control AfterUpdate closes its actual dynamically opened form
and enters a real MsgBox; the form survives the nested loop and is deleted
after unwinding, with one saved row. Actual Tab-key routing is separate from Qt's
implicit focus traversal: hiding/disabling a dirty focused control does not save
or navigate its record. Targeted GUI/original-heavy tests pass (34.91 s).
Original exports contain no explicit Cycle settings; original
execution counts are unchanged. Page, nested/subform and general focus/event
semantics remain open (FORM_CYCLE_RUNTIME.md).
Validation: host 21/21 (82.58 s), Debian 12 21/21 (65.55 s), Ubuntu 24.04
21/21 (60.42 s), GUI-disabled build/core 1/1 (0.67 s). Fresh host/distro DEBs
pass exact payload/backend/permission/fixture-exclusion checks; isolated installed
Debian/Ubuntu packages pass backend/counter execution and GUI startup. All four
original fixture hashes pass.

Select Case increment: multiline list, inclusive range, relational (optional Is)
and Else clauses execute through the native interpreter. The selector/default
value is resolved once and subsequent/inactive branches skip expression effects.
Nested If/For/Exit For, selector/case Resume and bounded Resume Next are tested.
The existing Qt transaction-button regression now executes Case branches,
including Option Compare Text: real SQL commit, rollback and handled duplicate-key
rollback retain their independently asserted balances. Core/layout/original-heavy
targeted tests pass (37.97 s). All 32 original modules have zero Select Case starts;
no additional original procedure execution is claimed. Full grammar/type/locale
and Access error-oracle boundaries remain open (SELECT_CASE_RUNTIME.md).

Validation: host 21/21 (86.07 s), Debian 12 21/21 (65.05 s), Ubuntu 24.04
21/21 (59.69 s), freshly rebuilt GUI-disabled core 1/1 (0.90 s). Fresh host/
distro DEBs pass exact payload/backend/permissions/fixture-exclusion checks;
isolated installed Debian/Ubuntu packages pass backend/counter execution and
GUI startup. All four original fixture hashes pass.

Do/While increment: Do Until, postcondition Loop While/Until, unconditional Do/Loop,
While/Wend and Exit Do now execute natively. Static exit binding removes inner
If/Select/With/For Each frames while preserving the outer object and For loop.
Tests cover entry/exit and Null conditions, innermost exit, mixed Exit For/Do,
body/condition error retry and the execution budget. The real Qt transaction
button performs both account updates through Do Until/Select Case; commit,
rollback and handled duplicate-key rollback preserve asserted SQL/control values.
Targeted core/layout/original-heavy tests pass (35.98 s). Original inventory has
seven existing Do While starts and none of the newly added loop starts; original
handler execution counts are unchanged. DO_LOOP_RUNTIME.md records the bounds.

Validation: host 21/21 (63.64 s), Debian 12 21/21 (66.30 s), Ubuntu 24.04
21/21 (60.17 s), rebuilt GUI-disabled core 1/1 (0.98 s). Fresh host/distro
DEBs pass exact payload/backend/permissions/fixture-exclusion checks; isolated
installed Debian/Ubuntu packages pass backend/counter execution and GUI startup.
All four original fixture hashes pass.

Control font increment: imported family, point size, weight, italic and underline
now seed real native text widgets from declared control-type templates with
per-property instance overrides. Template blocks do not become extra controls.
A regression exposed Qt retaining a 9-point combo editor beside a 12-point combo;
explicit editor and requeried popup font application fixes it. Native font metrics,
sibling styles, SQL save/mirroring and record navigation are tested. One original
form and one report layout are hash-verified and instantiated with an empty DB
and no original VBA; no additional original handler/dataset execution is claimed.
Targeted layout/original-heavy tests pass (35.10 s). Dynamic font APIs, missing
families and exact Windows rendering remain open (CONTROL_FONT_RENDERING.md).

Validation: host 21/21 (80.67 s), Debian 12 21/21 (68.45 s), Ubuntu 24.04
21/21 (60.40 s), GUI-disabled core 1/1 (0.92 s). Fresh host/distro DEBs
pass exact payload/backend/permissions/fixture-exclusion checks; isolated
installed Debian/Ubuntu packages pass backend/counter execution and GUI startup.
All four original fixture hashes pass.

Runtime font increment: native VBA text-control objects now read/write FontName,
FontSize, FontWeight, FontBold, FontItalic and FontUnderline. One renderer handles
imported and runtime styles, including combo editors/popups. A bound Qt/VBA button
changes all six while preserving pending masked Text, old Value, selection, focus
and SQL. Scalar defaults resolve once, numeric rounding/ranges and handled error
preflight are tested. A fresh instance restores imported fonts; an actual SQL view
remains read-only and clean while fonts change. Generic property probes style an
original form label and reject report assignment with 3251 until OnPrint exists,
without original VBA/table-row execution. Targeted tests pass (36.43 s). Qt defaults/fallback, full Windows/font semantics and design
persistence remain open (CONTROL_FONT_RENDERING.md).

Microsoft reference calibration found that Qt's bold() classification includes
600 while Access FontBold is false below 700. The getter now checks >=700;
600/699/700 are regressed. Documented unexported point-size defaults are 8/10,
not the host's previous 9. Report setters require OnPrint, which is not yet
implemented, and now reject before mutation instead of succeeding in preview.
Static rendering/getters still work. Font contract links and limits are recorded
in CONTROL_FONT_RENDERING.md; no Windows Access runtime oracle is claimed.

Runtime increment validation: host 21/21 (70.55 s), Debian 12 21/21 (69.71 s),
Ubuntu 24.04 21/21 (63.59 s), GUI-disabled core 1/1 (0.89 s). Fresh host/distro
DEBs pass exact payload/backend/permissions/fixture-exclusion checks; isolated
installed Debian/Ubuntu packages pass backend/counter execution and GUI startup.
All four original fixture hashes pass. These measurements follow the documented
Bold/default-size/report-Print corrections above.

ElseIf increment: multiline chains share one conditional frame and execute only
the first matching body. Function-side-effect counters prove later and inactive
conditions are skipped; Null is false. Nested Select/Do/For exits, condition/body
Resume and invalid syntax before effects are regressed. The native Qt transaction
button displays an independent branch mode while actual SQL account transfers
commit, rollback or roll back a handled duplicate-key failure. Branch/control
values and independently queried balances are asserted. Targeted tests pass
(36.71 s). Original modules have zero line-start ElseIf clauses; original business
execution counts are unchanged. ELSEIF_RUNTIME.md records grammar/oracle limits.

Validation: host 21/21 (65.43 s), Debian 12 21/21 (67.65 s), Ubuntu 24.04
21/21 (60.80 s), rebuilt GUI-disabled core 1/1 (1.29 s). Fresh host/distro
DEBs pass exact payload/backend/permissions/fixture-exclusion checks; isolated
installed Debian/Ubuntu packages pass backend/counter execution and GUI startup.
All four original fixture hashes pass.
