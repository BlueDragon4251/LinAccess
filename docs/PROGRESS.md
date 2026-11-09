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
| Desktop CTest targets, last complete suite | 48 passed |

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

AllowEdits increment: imported false/NotDefault and native Boolean properties now
protect bound clean existing records while preserving unbound/new-row input.
Programmatic Value changes temporarily enable current-record edits until save/
Undo, with validation cancellation and pending drafts retained. Locked/backend/
addition policies remain independent; record loading clears dirty state before
applying input policy. Deletion and in-command policy changes preflight safely.
Native checkbox Click is connected; grouped members dispatch only group Click,
correcting an older child-event regression against Microsoft documentation.
Callback scopes retain closing documents across nested dialogs and suppress a
following Click after AfterUpdate closes the form. Actual Qt/VBA/SQL tests cover
these paths (targeted 37.28 s). All five original disabled flags bind correctly;
zero original writable-source policy transitions occur because their native query
sources are already read-only. No original business-handler coverage is added.
ALLOW_EDITS_RUNTIME.md records the remaining lifecycle/query/oracle boundaries.

Validation: host 21/21 (83.77 s), Debian 12 21/21 (68.34 s), Ubuntu 24.04
21/21 (62.50 s), GUI-disabled core 1/1 (0.88 s). Fresh host/distro DEBs pass
exact payload/backend/permissions/fixture-exclusion checks; isolated installed
Debian/Ubuntu packages pass backend/counter execution and GUI startup. All four
original fixture hashes pass. Five original flag-retention probes and zero
original writable-source policy transitions are recorded separately.

Control default increment: DefaultValue now inherits declared control templates
and exposes native expression-string metadata getters/setters. Numeric/string/
Null/Boolean literals initialize unbound controls and writable-table new rows
without data events, Dirty changes, inserts or sequence advances. Existing rows
and current Value remain independent of metadata changes. Shared fields mirror
one default; conflicts and unsupported expressions raise explicit errors and
cannot insert a partial row. Manual input can resolve a failed default. Native
VBA/Qt/SQL tests cover Unicode, empty-versus-NULL, locked fields, overrides, error/
scalar/UTF-16 preflight and actual ID/row persistence (targeted 39.04 s).
All 34 original nonempty metadata entries are retained and 15 unbound literal
initializations match the separate native expression parser, using an empty DB
and no supplied original CodeBehind/table data. No original bound-new-record or
business-handler execution claim follows. Computed/Jet table/default timing/
typing boundaries remain open (CONTROL_DEFAULT_RUNTIME.md).

Validation: host 21/21 (87.53 s), Debian 12 21/21 (67.46 s), Ubuntu 24.04
21/21 (62.39 s), GUI-disabled core 1/1 (1.05 s). Fresh host/distro DEBs pass
exact payload/backend/permissions/fixture-exclusion checks; isolated installed
Debian/Ubuntu packages pass backend/counter execution and GUI startup. All four
original fixture hashes pass. Original counts remain 34 metadata / 15 unbound
literal initializations / zero unsupported adapters for these entries.

## 2026-11-02 — Literal native label captions

Removed Qt automatic HTML interpretation from imported form/report Label widgets.
HTML-looking Caption text is displayed literally at construction and after native
VBA Caption assignment. The existing Qt/SQLite font regression covers literal tags,
entity syntax and newlines while preserving fonts, getter values and clean data
state. This corrects rendering within the existing catalog scope; the checklist
count is unchanged. Full mnemonic and visual parity remains open.

Validation: host 21/21 (67.57 s), Debian 12 21/21 (69.62 s), Ubuntu 24.04
21/21 (63.46 s), GUI-disabled core 1/1 (1.15 s). Fresh packages pass exact
payload/backend/permissions/fixture-exclusion checks and isolated Debian/Ubuntu
backend and GUI startup. All four original fixture hashes remain unchanged.

## 2026-11-02 — Native horizontal text alignment

Imported per-control/template TextAlign now drives native TextBox/Label widgets.
Explicit zero overrides templates; General uses retained value/field types rather
than guessing from text or evaluating code. Native VBA updates only instance
appearance, preserving pending input, selection, focus, fonts and SQL permissions.
Distributed alignment and unsupported runtime report context reject explicitly.
The initial targeted pair passed 2/2 (39.17 s); an additional read-only-query
regression is included in final validation. An empty-DB/no-CodeBehind probe checks
24 hash-verified original form/report layouts with 208 Labels and 226 TextBoxes.
No additional original business execution or complete visual parity is claimed.
See TEXT_ALIGNMENT_RUNTIME.md for boundaries and public contracts.

Final validation: host 21/21 (87.10 s), Debian 12 21/21 (68.95 s), Ubuntu
24.04 21/21 (62.38 s), GUI-disabled core 1/1 (1.17 s). Fresh host/distro
DEBs pass exact payload/backend/permissions/fixture-exclusion checks. Isolated
installed Debian/Ubuntu packages pass backend/counter execution and GUI startup.
All four original fixture hashes pass; the repository remains private.

## 2026-11-02 — Native foreground colors and VBA RGB

Imported/inherited ForeColor metadata now colors native text adapters and combo
editors/popups. RGB and the two original system-color intents resolve correctly
from independent application palettes; other encodings and runtime report
setters reject explicitly. Native RGB(...) enables actual VBA color expressions.
Qt stylesheet initialization exposed palette resets: canvas background now uses
a native palette, password styling avoids redundant reapplication, and explicit
colors survive control-state updates/navigation. Native label/canvas pixels and
masked pending input/selection/focus, SQL save, read-only permissions, popup
requery, scalar/errors and fresh instance/report preflight are tested.
The initial targeted trio passed 3/3 (41.97 s); additional canvas/default-getter
checks are included in final validation. Empty-DB/no-CodeBehind original probes
check 45 layouts with 271 RGB and 277 system-color controls. Original 32-module
lexical inventory has zero RGB calls / zero .ForeColor references, establishing
no extra original business execution. See FOREGROUND_COLOR_RUNTIME.md.

Final validation: host 21/21 (86.73 s), Debian 12 21/21 (72.41 s), Ubuntu
24.04 21/21 (64.26 s), GUI-disabled core 1/1 (1.17 s). Fresh host/distro
DEBs pass exact payload/backend/permissions/fixture-exclusion checks. Isolated
installed Debian/Ubuntu packages pass backend/counter execution and GUI startup.
All four original fixture hashes pass. Earlier text-alignment commit b43e659
also completed private CI successfully.

## 2026-11-02 — Native control backgrounds

TextBox/Label BackColor/BackStyle import, per-type inheritance, zero overrides
and native VBA getters/setters now drive opaque or transparent interiors. The
shared color decoder adds window-interior and button-face system intents. Native
composite pixels prove Normal/Transparent behavior while pending masked input,
fonts/foreground/selection/focus and SQL permissions remain intact. Explicit save,
record navigation/requery, instance reset, scalar/error and report preflight are
covered. Targeted pair: 2/2 (38.30 s). An empty-DB/no-CodeBehind probe checks
55 original layouts: 326 Normal / 587 Transparent / 414 retained colors. Original
32-module lexical inventory has zero .BackColor/.BackStyle references; original
business coverage is unchanged. See CONTROL_BACKGROUND_RUNTIME.md for active/
inactive palette scope, unsupported report/other adapters and open visual gaps.

Final validation: host 21/21 (92.70 s), Debian 12 21/21 (72.02 s), Ubuntu
24.04 21/21 (65.21 s), GUI-disabled core 1/1 (1.02 s). Fresh host/distro DEBs
pass exact payload/backend/permissions/fixture-exclusion checks. Isolated
installed Debian/Ubuntu packages pass backend/counter execution and GUI startup.
All four original fixture hashes pass; repository privacy is unchanged. Earlier
foreground-color commit b067f9c completed private CI successfully.

## 2026-11-02 — Native SQL RGB and DAO color expressions

Access SQL RGB(...) now registers a real SQLite adapter and shares its channel
conversion with native VBA. Compiler function/purity recognition handles alias
reuse and compiled joined-source views, retaining quoted RGB field identifiers.
SQL error transport now preserves 450 (previously generic 3146). Actual DAO VBA
reads calculated color fields and persists updates; NULL failures roll back the
statement, including materialized joined updates. Conflicting source-view colors
produce 3251 without modifying targets. Core/DAO/query-model/join targeted tests
pass 4/4 (1.07 s), including parameters, types, arity, clamping and result reset.
Original 100-query artifact / 32-module lexical RGB inventories are zero; no
additional original business execution or full Jet parity is claimed. See
SQL_RGB_RUNTIME.md.

Final validation: host 21/21 (71.72 s), Debian 12 21/21 (73.02 s), Ubuntu
24.04 21/21 (66.93 s), GUI-disabled core 1/1 (0.90 s). Fresh host/distro DEBs
pass exact payload/backend/permissions/fixture-exclusion checks. Isolated
installed Debian/Ubuntu packages pass backend/counter execution and GUI startup.
All four original fixture hashes pass. Earlier control-background commit
714fa54 completed private CI successfully.

## 2026-11-02 — Native classic QBColor palette

The shared native color module now executes QBColor in VBA and SQL, using the
fixed 16-color values confirmed by Microsoft's public reference table. This
requires no .NET dependency and does not provide an Access runtime oracle.
Core tests cover all indices, scalar/rounding/type/range/arity behavior; real
Qt ForeColor/Eval changes preserve pending masked input and data. SQL tests
preserve quoted fields and parameter/alias behavior, while native DAO palette-
view joined updates persist valid colors and roll back invalid source indices.
Targeted core/DAO/query-model/join/GUI set: 5/5 (9.75 s). Original lexical QBColor
call inventory is zero across 100 query artifacts / 32 VBA modules; original
business coverage is unchanged. Full Variant/coercion/theme/Windows parity
remains open; see QBCOLOR_RUNTIME.md.

Final validation: host 21/21 (71.39 s), Debian 12 21/21 (69.67 s), Ubuntu
24.04 21/21 (63.52 s), GUI-disabled core 1/1 (1.97 s). Fresh host/distro DEBs
pass exact payload/backend/permissions/fixture-exclusion checks. Isolated
installed Debian/Ubuntu packages pass backend/counter execution and GUI startup.
All four original fixture hashes pass. Earlier SQL RGB commit 1c82bcb completed
private CI successfully.

## 2026-11-02 — Main-window imported-document lifetime

Source-inspector form/report opening now closes the details dialog and uses
managed native windows with imported Modal metadata, instead of stack previews
inside exec. Repeated opening reuses the same project/kind/name instance; read
cursors finish before construction. Deferred disposal protects closing documents
through outer VBA calls. The new mainwindow_documents test compiles actual GUI
sources and drives the real project/file-dialog/explorer path on temporary
SQLite projects, proving Load once, independent nonmodal form/report lifetime,
actual SQL save, cancelled project switch/Unload, modal/Open-cancel behavior and
self-close/MsgBox retention. Initial targeted trio: 3/3 (38.93 s). Full fixture-
enabled desktop suite now has 22 CTests. No original business-coverage increment
is claimed. The legacy partially converted/generated form editor retains its
modal opener pending caller/renderer migration, since changing its child alone
would block that child. Full F11/MDI/acDialog/focus parity remains open; see
MAINWINDOW_DOCUMENT_RUNTIME.md.

Final validation: host 22/22 (89.26 s), Debian 12 22/22 (72.86 s), Ubuntu
24.04 22/22 (64.99 s), GUI-disabled core 1/1 (0.94 s). Fresh host/distro DEBs
pass exact payload/backend/permissions/fixture-exclusion checks. Isolated
installed Debian/Ubuntu packages pass backend/counter execution and GUI startup.
All four original fixture hashes pass; repository privacy is unchanged. Earlier
QBColor commit 6920d36 completed private CI successfully.

## 2026-11-02 — Classic F11 and database-window context

F11 activates the actual main project explorer and restores collapsed width,
matching the archived 2003 bring-to-front contract rather than modern hide/show.
It follows native KeyPreview/BeforeUpdate handling, respects modal windows and
does not call form-record save. Logical database-window context clears stale
Screen targets and blocks default record/GoToControl operations; unsupported
Close/SendKeys targets reject explicitly. Named/explicit-owner scopes retain
existing behavior, including Me.Dirty=False while the explorer is active.
A first-activation marker preserves early Current initialization. Actual GUI/
SQL tests cover all these paths, repeated F11, Deactivate/Activate, modal/control
cancellation and the existing close/MsgBox/project-switch regressions. Final
targeted trio: 3/3 (39.79 s). Original lexical vbKeyF11/SelectObject inventories
are zero across 32 modules, with no new business-execution claim. Full WM/MDI/
keyboard/error identity remains open; see DATABASE_WINDOW_RUNTIME.md.

Final validation: host 22/22 (90.58 s), Debian 12 22/22 (72.36 s), Ubuntu
24.04 22/22 (66.48 s), GUI-disabled core 1/1 (1.08 s). Fresh host/distro DEBs
pass exact payload/backend/permissions/fixture-exclusion checks. Isolated
installed Debian/Ubuntu packages pass backend/counter execution and GUI startup.
All four original fixture hashes pass. Earlier main-window document commit
da225b8 completed private CI successfully.

## 2026-11-02 — Native SelectObject and object grouping

SelectObject now uses actual project-tree metadata for table/query/form/report
selection and native open Form/Report activation. Imported source queries, forms
and reports join their corresponding groups, preserving source status/details.
Optional empty names choose groups; case-insensitive/scalar identities resolve
once, with explicit errors for missing or ambiguous targets instead of guessed
aliases. Control validation and modal blocking preflight selection; dirty form
records do not save merely on object selection. Real GUI/VBA/SQL regressions
include a saved query artifact, hidden-form restoration, two-window activation,
input cancellation, scalar access and intentional source/user-name collision.
Targeted core/GUI/layout/original-heavy set: 4/4 (41.01 s). Original 32-module
SelectObject lexical inventory is zero, so original business coverage is
unchanged. Open table/query window and macro/module targets remain unsupported;
see SELECT_OBJECT_RUNTIME.md for scope and public contract.

Final validation: host 22/22 (82.81 s), Debian 12 22/22 (72.48 s), Ubuntu
24.04 22/22 (65.76 s), GUI-disabled core 1/1 (1.24 s). Fresh host/distro DEBs
pass exact payload/backend/permissions/fixture-exclusion checks. Isolated
installed Debian/Ubuntu packages pass backend/counter execution and GUI startup.
All four original fixture hashes pass. Earlier F11/context commit 6967113
completed private CI successfully.

## 2026-11-02 — Project open/create keyboard shortcuts

Ctrl+O/Ctrl+N now route to real project dialogs after native KeyPreview/control
validation. Real production GUI tests cover both dialog cancellations, modal
blocking, Unload cancellation, successful open/create and byte-identical
existing-file refusal. Synchronous native SendKeys cannot replace its executing
module's project. Targeted trio: 3/3 (39.37 s). Scope remains SQLite project
operations, with direct MDB/ACCDB editing and full keyboard/File-menu parity
open; see PROJECT_SHORTCUTS_RUNTIME.md. No original business increment.

Final validation: host 22/22 (74.54 s), Debian 12 22/22 (73.35 s), Ubuntu
24.04 22/22 (66.01 s), GUI-disabled core 1/1 (1.04 s). Fresh host/distro DEBs
pass payload/backend/permissions/fixture-exclusion checks. Isolated installed
Debian/Ubuntu packages pass backend/counter execution and GUI startup. All four
original fixture hashes pass. SelectObject commit c675477 completed private CI
successfully. Repository privacy remains unchanged.

## 2026-11-02 — Native project window cycling

Ctrl+F6/Ctrl+Shift+F6 cycle actual main database/form/report windows in a bounded
creation-order ring. Actual keyboard/focus/SQL regressions cover both directions
and wrapping, KeyPreview/control/modal cancellation, hidden/foreign/closed
exclusion, minimized restoration and a target report closed from BeforeUpdate.
Native execution guards and weak targets prevent stale activation. Initial
main-window regression: 1/1 (1.31 s). Full MDI/window-order/table/query/editor
parity remains open, with no original business coverage increment; see
WINDOW_CYCLE_RUNTIME.md.

Final validation: host 22/22 (87.58 s), Debian 12 22/22 (72.23 s), Ubuntu
24.04 22/22 (64.39 s), GUI-disabled core 1/1 (0.95 s). Fresh host/distro DEBs
pass payload/backend/permissions/fixture-exclusion checks. Isolated installed
Debian/Ubuntu packages pass backend/counter execution and GUI startup. All four
original fixture hashes pass. Repository privacy is unchanged.

## 2026-11-02 — Focused lookup keyboard refresh

F9 reloads the focused native combo adapter's actual row model, then recalculates
form expressions. Other focus and Shift+F9 keep their existing scopes. Native
keyboard cancellation and execution guards apply. Actual GUI/SQL tests cover
F4 opening/cancellation, F9 cancellation and external updates, retained bound ID,
pending text/focus/dirty state, failed-query model preservation, source recovery
and Escape undo without SQL saving. Targeted trio: 3/3 (41.19 s). Complete
ListBox/popup/Windows timing remains open; no original business increment.
See LOOKUP_KEYBOARD_RUNTIME.md.

Final validation: host 22/22 (74.18 s), Debian 12 22/22 (80.56 s), Ubuntu
24.04 22/22 (71.75 s), GUI-disabled core 1/1 (1.02 s). Fresh host/distro DEBs
pass payload/backend/permissions/fixture-exclusion checks. Isolated installed
Debian/Ubuntu packages pass backend/counter execution and GUI startup. All four
original fixture hashes pass. Repository privacy remains unchanged. Earlier
project-shortcut commit c571fe6 completed private CI successfully.

## 2026-11-02 — Native ComboBox list introspection

Read-only ListIndex/ItemData expose actual selected indices and typed bound row
values, including ordinal BoundColumn=0. GUI/VBA/SQL regressions cover lazy load,
null/unmatched/empty selections, scalar-once/index conversions and errors,
read-only properties, pending-buffer/report boundaries and removal/requery.
The tests found and fixed NULL-to-empty-string matching in the shared widget
setter across load, assignment, mirroring and requery. Native ListBox and exact
Windows pending-buffer/error/subtype behavior remain open. Targeted trio: 3/3
(40.11 s). No original business execution increment; see
COMBO_LIST_PROPERTIES_RUNTIME.md.

Final validation: host 22/22 (77.77 s), Debian 12 22/22 (74.25 s), Ubuntu
24.04 22/22 (67.12 s), GUI-disabled core 1/1 (0.92 s). Fresh host/distro DEBs
pass payload/backend/permissions/fixture-exclusion checks. Isolated installed
Debian/Ubuntu packages pass backend/counter execution and GUI startup. All four
original fixture hashes pass. Earlier window-cycle commit f09241c completed
private CI successfully. Repository privacy remains unchanged.

## 2026-11-02 — Persistent native single-selection ListBox

ListBox now uses a persistent QTreeView over the shared real typed row-source
model. SQL/Value/Field Lists, bound/unbound/default/NULL values, hidden widths,
actual mouse/arrow selection, control validation/cancellation and SQL save/Undo
are tested. Locked/report display blocks user value edits; F9/source assignment
reloads rows and failed initial sources recover. Old selection models are
released through Qt ownership. MultiSelect metadata inherits templates and
explicit zero overrides; modes 1/2 show an explicit unsupported diagnostic.
Targeted trio: 3/3 (40.87 s). Complete headers/multi-selection/ListBox API/Windows
visual/event identity remains open. No original business coverage increment;
see NATIVE_LISTBOX_RUNTIME.md. Earlier ComboBox introspection commit 85d7436
completed private CI successfully.

Final validation: host 22/22 (89.70 s), Debian 12 22/22 (74.62 s), Ubuntu
24.04 22/22 (66.04 s), GUI-disabled core 1/1 (0.91 s). Fresh host/distro DEBs
pass payload/backend/permissions/fixture-exclusion checks. Isolated installed
Debian/Ubuntu packages pass backend/counter execution and GUI startup. All four
original fixture hashes pass. Repository privacy remains unchanged.

## 2026-11-02 — Native column headings and header-aware ListCount

ColumnHeads now retains source templates and explicit false overrides, exposes
native VBA state and renders actual query/value-list QHeaderViews. Query names/
aliases supply labels; Value Lists consume their first row as headings. ListCount
includes the separate visual heading per the documented rule. Actual ComboBox/
ListBox GUI/VBA/SQL tests cover header clicks, toggling, bound/dirty state, empty
counts, scalar/errors, data ordinals and report context. Initial targeted layout
test: 1/1 (9.01 s). Field captions/Field List headers/exact Windows extra-index/
visual/error context remain open; no new original business execution is claimed.
See COLUMN_HEADS_RUNTIME.md.

Final validation: host 22/22 (89.09 s), Debian 12 22/22 (73.71 s), Ubuntu
24.04 22/22 (65.85 s), GUI-disabled core 1/1 (0.89 s). Fresh host/distro DEBs
pass payload/backend/permissions/fixture-exclusion checks. Isolated installed
Debian/Ubuntu packages pass backend/counter execution and GUI startup. All four
original fixture hashes pass. Previous native ListBox commit b916960 completed
private CI successfully. Repository privacy remains unchanged.

## 2026-11-02 — Native ListBox multiple selection and indexed VBA assignment

Simple/Extended Qt selection now backs Selected/ItemsSelected, ItemData and
ListIndex getters. An explicit indexed setter protocol handles nested VBA/With/
bracket targets without getter-as-assignment effects. Actual mouse/keyboard,
collection/For Each/scalar/error/lifetime and validation tests pass. Multi Value/
OldValue stay Null while record loading, comparison, mirroring and persistence
exclude these non-value controls. Saving another or the same bound field leaves
selection intact; selection alone performs zero inserts on a new row. Extended
requery clears; selected Simple requery rejects without model/selection loss.
Targeted core/main-window/layout/original-heavy set: 4/4 (40.66 s). Full Windows
order/index/type/context, Simple preservation and ListIndex writer remain open;
no original business execution increment. See MULTISELECT_LISTBOX_RUNTIME.md.

Final validation: host 22/22 (86.50 s), Debian 12 22/22 (72.69 s), Ubuntu
24.04 22/22 (66.07 s), GUI-disabled core 1/1 (0.93 s). Fresh host/distro DEBs
pass payload/backend/permissions/fixture-exclusion checks. Isolated installed
Debian/Ubuntu packages pass backend/counter execution and GUI startup. All four
original fixture hashes pass. Previous column-heading commit 9b4fa12 completed
private CI successfully. Repository privacy remains unchanged.

## 2026-11-02 — Native ComboBox popup row cap

ListRows now retains templates/default 16 and drives the actual Qt maximum
visible data-row setting. Native VBA/GUI/SQL tests measure 3/5-row viewports,
scroll and persist an item beyond the initial viewport, and retain pending text,
dirty state and model identity. Integer/property range, scalar-once, invalid
metadata and report limits are checked. Initial targeted layout suite: 1/1
(9.70 s). Full Windows heading/theme/screen-edge/open-popup timing remains open;
no original business coverage increment. See COMBO_LISTROWS_RUNTIME.md.

Final validation: host 22/22 (91.26 s), Debian 12 22/22 (74.80 s), Ubuntu
24.04 22/22 (67.20 s), GUI-disabled core 1/1 (0.94 s). Fresh host/distro DEBs
pass payload/backend/permissions/fixture-exclusion checks. Isolated installed
Debian/Ubuntu packages pass backend/counter execution and GUI startup. All four
original fixture hashes pass. Previous multiple-selection commit fc7ba53
completed private CI successfully. Repository privacy remains unchanged.

## 2026-11-02 — Integer and Long module fields

Native module declarations now admit Integer/Long with zero defaults and rounded
16/32-bit assignment checks. Core and actual GUI/VBA/SQL tests verify overflow/
Null/type errors without stored-value corruption, visibility/shadow rules,
per-instance private state and shared standard-module Long state, database
isolation, close/reopen/project reset and SQL insertion. Targeted core/main-window/
layout/original-heavy set: 4/4 (42.34 s). Full subtype/intermediate arithmetic/
ByRef/array/class behavior remains open; no extra original business execution
is claimed. See NUMERIC_MODULE_FIELDS_RUNTIME.md.

Final validation: host 22/22 (75.17 s), Debian 12 22/22 (75.89 s), Ubuntu
24.04 22/22 (69.00 s), GUI-disabled core 1/1 (0.90 s). Fresh host/distro DEBs
pass payload/backend/permissions/fixture-exclusion checks. Isolated installed
Debian/Ubuntu packages pass backend/counter execution and GUI startup. All four
original fixture hashes pass. Previous popup-row commit 552f10f completed private
CI successfully. Repository privacy remains unchanged.

## 2026-11-02 — Checked Integer/Long procedure values

Parameter types now survive parsing into execution. Checked assignments cover
Integer/Long locals, parameters/results and integral For counters. Known ByRef
conflicts reject before callee effects; matching references and parenthesized
temporaries retain native copyback scope. Invalid numeric Set/For Each targets
reject before body effects. Core and native GUI/VBA/SQL tests cover rounding,
overflow/preserved values, SQL persistence and zero SQL effects for invalid
arguments/references. Targeted four-test set: 4/4 (43.64 s). Full subtype/promotion/
Variant-reference/loop/compile/error timing remains open; no new original business
execution is claimed. See NUMERIC_PROCEDURE_VALUES_RUNTIME.md.

Final validation: host 22/22 (79.98 s), Debian 12 22/22 (78.62 s), Ubuntu
24.04 22/22 (70.77 s), GUI-disabled core 1/1 (1.11 s). Fresh host/distro DEBs
pass payload/backend/permissions/fixture-exclusion checks. Isolated installed
Debian/Ubuntu packages pass backend/counter execution and GUI startup. All four
original fixture hashes pass. Previous numeric-module commit 581862c completed
private CI successfully. Repository privacy remains unchanged.

## 2026-11-02 — Checked Boolean and String procedure values

Procedure locals, parameters, results and known-reference metadata now preserve
Boolean/String conversion semantics. Core and native GUI/VBA/DAO tests cover
False text, typed result kinds, error preservation, ByVal/reference/temporary
behavior, Unicode/apostrophes and stored Boolean 0/-1 values. Invalid declared
scalar loop/object targets reject before body effects. A synthetic parameter
test exposed the separate unsupported BIT alias and was corrected to the already
supported BOOLEAN spelling; no BIT support is implied. Targeted four-test set:
4/4 (43.12 s). Full Variant/object-default/locale/compiler/reference semantics
remain open; no original business execution increment. See
SCALAR_PROCEDURE_VALUES_RUNTIME.md.

Full validation: host 22/22 (75.41 s), Debian 12 22/22 (77.21 s), Ubuntu
24.04 22/22 (69.27 s), GUI-disabled core 1/1 (1.25 s). Host and distro
DEBs pass payload/backend/permissions/fixture-exclusion checks; isolated distro
installations pass backend execution and GUI startup. All four original fixture
hashes pass. Previous ae2f65b private CI completed successfully. Remote repository
privacy remains enabled.

## 2026-11-02 — Boolean SQL parameter aliases and strict binding

Added Access BIT/LOGICAL/LOGICAL1 aliases to the existing BOOLEAN/YESNO parameter
parser. Invalid text no longer becomes true through Qt conversion. Native DAO
tests execute all aliases and check stored 0/-1/NULL and unchanged records after
error 13; direct binding rejects empty text and nonfinite numbers. See
SQL_BOOLEAN_PARAMETERS.md for the primary reference and bounded scope.

Validation: host 22/22 (80.58 s), Debian 12 22/22 (86.58 s), Ubuntu
24.04 22/22 (76.93 s), GUI-disabled core 1/1 (1.05 s). Host/distro DEBs
pass payload/backend/permissions/fixture-exclusion checks. Isolated Debian/Ubuntu
package installations pass backend execution and GUI startup. All four fixture
hashes pass; repository privacy remains enabled. Previous 4ffa700 private CI
is still running at this checkpoint; no CI success is assumed.

## 2026-11-02 — Native VBA CBool conversion

Added actual Boolean CBool results, one-time object scalar resolution and native
94/13/91/450 errors. Shared typed Boolean conversion now accepts numeric text
with surrounding spaces. Core and real VBA/DAO tests verify Boolean type identity,
Currency, Empty, text/numeric conversion, errors and stored SQL zero/true=-1.
Targeted tests 2/2 (1.18 s). Scope and remaining gaps: CBOOL_RUNTIME.md.

Full validation: host 22/22 (78.97 s), Debian 12 22/22 (76.80 s), Ubuntu
24.04 22/22 (67.83 s), GUI-disabled core 1/1 (1.57 s). Host/distro DEBs
pass payload/backend/permissions/fixture-exclusion checks; isolated Debian/Ubuntu
installations pass backend execution and GUI startup. All four original fixture
hashes pass; repository remains private. Previous 3add688 private CI is still
running at this checkpoint, not counted as success.

## 2026-11-02 — Integer/Long storage identity and native CInt

Checked Integer/Long module fields, locals, parameters/results and initial values
now retain numeric subtype tags. CInt performs checked 16-bit ties-to-even
conversion; CLng returns a tagged Long. VarType/TypeName and standard VbVarType
constants expose these identities; Variant copies retain tags. Reference/change
checks include numeric tags. Core and actual native VBA/DAO tests cover default
types, copies, stored rounded integer values and overflow preservation.
Storage-aware Len, literal/arithmetic type rules and other primitive types remain
open; no original business execution coverage increment. See INTEGER_VALUE_IDENTITY.md.

Targeted core/DAO/layout/original-heavy tests: 4/4 (41.34 s). Full validation:
host 22/22 (88.34 s), Debian 12 22/22 (77.28 s), Ubuntu 24.04 22/22
(68.59 s), GUI-disabled core 1/1 (1.38 s). Host/distro DEBs pass payload,
backend, permissions and fixture-exclusion checks. Isolated Debian/Ubuntu
installations pass backend execution and GUI startup. All four original fixture
hashes pass; repository remains private. Previous 472a23b private CI completed
successfully.

## 2026-11-02 — Shared native GUI build implementation

App, main-window integration tests and layout tests now link the same static
linaccess_forms library instead of compiling the form implementation separately
for each executable. This reduces repeated compilation while retaining native
execution and the same installed executable; no runtime shared library or QtTest
app dependency is introduced. The GUI-disabled build remains independent of Qt.
This build change adds no compatibility catalog point.

Validation: host 22/22 (75.02 s), Debian 12 22/22 (80.86 s), Ubuntu
24.04 22/22 (72.63 s), GUI-disabled core 1/1 (1.23 s). Host/distro DEBs
pass payload/backend/permissions/fixture-exclusion checks. Isolated Debian/Ubuntu
installations pass backend execution and GUI startup. Build rules contain one
form-preview compile target; the app has no QtTest runtime dependency. All four
original fixture hashes pass and the repository remains private. Previous
449f17a private CI was still running at this checkpoint; no success is assumed.

## 2026-11-02 — Checked Single/Double values and CSng

Added checked Single/Double module fields and procedure values, typed defaults/
results, native CSng and shared scalar conversion dispatch. Native Qt/VBA/DAO
regressions store distinct rounded Single/Double values, verify overflow/input
atomicity and bounded reference/loop behavior. Core/DAO/layout/original-heavy
tests pass 4/4 (42.76 s). Single underflow and For remain explicitly diagnosed;
full formatting/promotion/literal/Variant/locale/compiler parity remains open.
No original business execution increment. See FLOATING_VALUES_RUNTIME.md.

Full validation: host 22/22 (92.89 s), Debian 12 22/22 (77.18 s), Ubuntu
24.04 22/22 (70.00 s), GUI-disabled core 1/1 (1.90 s). Host/distro DEBs
pass payload/backend/permissions/fixture-exclusion checks. Isolated Debian/Ubuntu
installations pass backend execution and GUI startup. All four fixture hashes
pass; repository remains private. Previous 9bf6a90 and 449f17a private CI
completed successfully.

## 2026-11-02 — Checked Currency storage and exact DAO text transfer

Added declared Currency module/procedure storage and exact checked conversions/
defaults/results/references. The DAO value adapter preserves canonical Currency
text instead of rejecting the kind or converting to Double. Actual native
Qt/VBA/DAO tests persist 922337203685477.5807 through a TEXT parameter, verify SQL
storage class, overflow preservation and no SQL effects after invalid arguments
or references. Currency loops reject before effects. Targeted tests: 4/4
(42.72 s). Full Jet Currency SQL/columns and mixed promotion remain open.
No original business execution increment. See CURRENCY_STORAGE_RUNTIME.md.

Full validation: host 22/22 (77.08 s), Debian 12 22/22 (79.40 s), Ubuntu
24.04 22/22 (70.83 s), GUI-disabled core 1/1 (1.07 s). Host/distro DEBs
pass payload/backend/permissions/fixture-exclusion checks. Isolated Debian/Ubuntu
installations pass backend execution and GUI startup. All four original fixture
hashes pass; repository remains private. Previous a8db286 private CI was still
running at this checkpoint; no success is assumed.

## 2026-11-02 — SQL numeric parameter types and DAO Type metadata

Corrected SQL INTEGER from 16-bit to the documented 32-bit Long domain and added
numeric aliases with real Single binding precision/finite checks. Declared DAO
Parameter.Type returns documented Integer/DataTypeEnum values and is read-only
for the native Access workspace. Common unqualified DAO type constants are
available. Tests cover bounds, NULL, invalid/nonfinite values without SQL effects,
metadata arity/read-only/replacement/close behavior and actual VBA Type=dbLong
execution storing 40000. Five targeted tests pass (42.06 s); final native read-only
VBA case passes DAO-only verification (0.26 s). Inference, Size and full conversion
contracts remain open. See DAO_PARAMETER_TYPES.md. No business execution increment.

Full validation: host 22/22 (75.62 s), Debian 12 22/22 (83.00 s), Ubuntu
24.04 22/22 (73.46 s), GUI-disabled core 1/1 (1.02 s). Host/distro DEBs
pass payload/backend/permissions/fixture-exclusion checks. Isolated Debian/Ubuntu
installations pass backend execution and GUI startup. All four original fixture
hashes pass; repository remains private. Previous d8557a7 private CI completed
successfully.

## 2026-11-02 — Native integer division and corrected Mod

Added native backslash integer division and replaced floating Mod with checked
rounded integer remainder. Corrected multiplication/integer-division/Mod
precedence; signed, Null/Empty, range and zero-denominator handling are tested.
A new fixture-independent Qt/VBA/SQL regression verifies calculated controls,
pending edits, explicit save and zero-division/overflow preventing later SQL.
Targeted core/GUI tests pass 2/2 (0.96 s). GUI suite now has 23 CTests.
SQL rewriting and complete numeric promotion remain open; no original business
execution increment. See INTEGER_OPERATORS_RUNTIME.md.

Full validation: host 23/23 (84.08 s), Debian 12 23/23 (87.04 s), Ubuntu
24.04 23/23 (78.45 s), GUI-disabled core 1/1 (1.46 s). Host/distro DEBs
pass payload/backend/permissions/fixture-exclusion checks. Isolated Debian/Ubuntu
installations pass backend execution and GUI startup. All four original fixture
hashes pass; repository remains private. Previous 27b7f99 private CI completed
successfully.

## 2026-11-02 — Native VBA power operator

Added caret power with documented left-to-right chaining and corrected prefix
operand precedence for unary signs. Core and native Qt/VBA/SQL tests verify
calculated powers, actual audit writes, Double identity and domain/overflow
preventing later SQL. Targeted tests pass 2/2 (1.19 s). Zero-negative powers and
complete underflow explicitly diagnose 3251; SQL rewriting and full libm/type/
error identity remain open. No original business execution increment. See
POWER_OPERATOR_RUNTIME.md.

Full validation: host 23/23 (75.61 s), Debian 12 23/23 (81.19 s), Ubuntu
24.04 23/23 (70.64 s), GUI-disabled core 1/1 (1.08 s). Host/distro DEBs
pass payload/backend/permissions/fixture-exclusion checks. Isolated Debian/Ubuntu
installations pass backend execution and GUI startup. All four original fixture
hashes pass; repository remains private. Previous 653fd0e private CI completed
successfully.

## 2026-11-02 — Native logical exclusion, equivalence and implication

Added Xor/Eqv/Imp Boolean/null tables, checked numeric bitwise operations and
logical precedence while retaining arithmetic/power/prefix rules. Core and native
Qt/VBA/SQL regressions cover truth tables, bit patterns, calculated Boolean
controls, stored -1/0 values and range errors preventing writes. Targeted tests
pass 2/2 (0.97 s). SQL rewriting, full mixed-type promotion and existing And/Or
edge semantics remain open. No business execution increment. See
LOGICAL_OPERATORS_RUNTIME.md.

Full validation: host 23/23 (77.68 s), Debian 12 23/23 (80.48 s), Ubuntu
24.04 23/23 (72.18 s), GUI-disabled core 1/1 (1.23 s). Host/distro DEBs
pass payload/backend/permissions/fixture-exclusion checks. Isolated Debian/Ubuntu
installations pass backend execution and GUI startup. All four original fixture
hashes pass; repository remains private. Previous ddf7490 private CI was still
running at this checkpoint; no success is assumed.

## 2026-11-02 — Native Fix/Sgn and checked Abs/Int

Added Fix and Sgn; corrected documented Abs/Int Null propagation and known
numeric subtype preservation, including checked minimum-integer Abs overflow.
Currency Fix uses exact scaled units. Core and fixture-independent native
Qt/VBA/SQL tests verify negative rounding/sign values, Currency precision and
overflow preventing later SQL. Targeted tests pass 2/2 (1.02 s). Full Variant/
Byte/Date/locale/promotion/SQL function parity remains open. No business execution
increment. See NUMERIC_FUNCTIONS_RUNTIME.md.

Full validation: host 23/23 (74.01 s), Debian 12 23/23 (80.44 s), Ubuntu
24.04 23/23 (72.40 s), GUI-disabled core 1/1 (1.32 s). Host/distro DEBs
pass payload/backend/permissions/fixture-exclusion checks. Isolated Debian/Ubuntu
installations pass backend execution and GUI startup. All four original fixture
hashes pass; repository remains private. Previous cefe109 private CI completed
successfully.

## 2026-11-02 — Native previous-field copy shortcut

Added bounded Ctrl+apostrophe field-value replacement for writable native bound
TextBoxes through existing user buffers/events/validation/undo, preserving NULL/
empty and preventing implicit record save. A new fixture-independent Qt/VBA/SQL
CTest covers Unicode, KeyPreview, cancelled/retried validation event ordering,
read-only/capacity/binary guards, new-row/no-INSERT/undo and SendKeys delivery.
Targeted test passes 1/1 (0.26 s); GUI suite now has 24 CTests. Exact Access
edit-mode selection, other adapters and complete keyboard parity remain open.
No business execution increment. See COPY_PREVIOUS_FIELD_RUNTIME.md.

Full validation: host 24/24 (73.74 s), Debian 12 24/24 (78.93 s), Ubuntu
24.04 24/24 (71.67 s), GUI-disabled core 1/1 (0.96 s). Host/distro DEBs
pass payload/backend/permissions/fixture-exclusion checks. Isolated Debian/Ubuntu
installations pass backend execution and GUI startup. All four original fixture
hashes pass; repository remains private. Previous a374198 private CI completed
successfully.

## 2026-11-02 — Native new-record keys and navigation RunCommand

Added Ctrl+Plus/Ctrl+Shift+Equals and native SendKeys through existing record
navigation, including pre-save rejection of disallowed new-record actions.
Added documented AcCommand new/next/previous/first/last mappings and constants.
Fixture-independent Qt/VBA/SQL tests verify key/validation cancellation, pending
record preservation, permitted save/navigation, no implicit new-row INSERT and
correct command identities. Targeted tests pass 3/3 (1.58 s). Full keyboard/
layout/navigation/query semantics remain open. No business execution increment.
See NEW_RECORD_SHORTCUT_RUNTIME.md.

Full validation: host 24/24 (76.77 s), Debian 12 24/24 (81.56 s), Ubuntu
24.04 24/24 (71.67 s), GUI-disabled core 1/1 (1.29 s). Host/distro DEBs
pass payload/backend/permissions/fixture-exclusion checks. Isolated Debian/Ubuntu
installations pass backend execution and GUI startup. All four original fixture
hashes pass; repository remains private. Previous 1670fd1 private CI completed
successfully.

## 2026-11-02 — Native record deletion shortcut

Added Ctrl+Minus through the existing active-document deletion command. A new
fixture-independent Qt/VBA/SQL CTest verifies KeyPreview/policies/validation,
Delete and confirmation callback cancellation, default-No real dialogs, user
No/Yes, SendKeys and an attempted deletion rejected by a foreign key. Data and
position/after-status checks cover rollback and commit behavior. Targeted tests
pass 2/2 (0.90 s); GUI suite now has 25 CTests. Full batch/datasheet/query/Windows
keyboard identity remains open. No business execution increment. See DELETE_KEY_RUNTIME.md.

Full validation: host 25/25 (76.88 s), Debian 12 25/25 (81.66 s), Ubuntu
24.04 25/25 (72.69 s), GUI-disabled core 1/1 (1.31 s). Host/distro DEBs
pass payload/backend/permissions/fixture-exclusion checks. Isolated Debian/Ubuntu
installations pass backend execution and GUI startup. All four original fixture
hashes pass; repository remains private. Previous 07d8511 private CI completed
successfully.

## 2026-11-02 — Preserve long native TextBox values

Reproduced a native SQL round-trip failure when saving another field while a
long Memo was limited by QLineEdit's default capacity. Native TextBoxes now use
the int editor capacity. A new fixture-independent Qt/VBA/SQL regression preserves
over 40000 UTF-16 units with Unicode/newlines through unrelated-field update,
VBA append and Ctrl+apostrophe pending copy/save. Targeted tests pass 4/4
(0.99 s); GUI suite now has 26 CTests. Full Memo multiline/format/field-size
parity remains open. No business execution increment. See LONG_TEXT_VALUES_RUNTIME.md.

Full validation: host 26/26 (77.64 s), Debian 12 26/26 (83.70 s), Ubuntu
24.04 26/26 (74.38 s), GUI-disabled core 1/1 (1.17 s). Host/distro DEBs
pass payload/backend/permissions/fixture-exclusion checks. Isolated Debian/Ubuntu
installations pass backend execution and GUI startup. All four original fixture
hashes pass; repository remains private. Previous 8c97bcc private CI completed
successfully.

## 2026-11-02 — Native modeless literal form search

Added actual Ctrl+F/acCmdFind search UI and literal text matching/navigation over
complete native table forms, with field/match/case/advance modes. Native Qt/VBA/
SQL tests cover Unicode/long text, active-document calls, masked-prefill safety,
validation both before and during dialog use, unsaved current/no-match drafts,
SQL save on navigation and Unload/close lifetime. Targeted tests pass 4/4
(1.02 s); GUI suite now has 27 CTests. Synthetic UI was visually inspected in
a private screenshot. Full Find/Replace/wildcard/formatted/source/locale/MDI
identity remains open. No business execution increment. See FIND_FORM_RUNTIME.md.

Full validation: host 27/27 (109.53 s), Debian 12 27/27 (85.29 s), Ubuntu
24.04 27/27 (74.94 s), and first Debian 13/trixie full run 27/27 (71.14 s).
GUI-disabled core passes 1/1 (1.38 s). Host and all three distro DEBs pass
payload/backend/permissions/fixture-exclusion checks. Isolated installations on
all three distributions pass backend execution and GUI startup. All four
original fixture hashes pass; repository remains private. Previous 1b5fa5f
private CI completed successfully.

## 2026-11-02 — Repeat native form search after closing the dialog

Moved actual matching/navigation into a lifetime-guarded native method with
per-document search history. Shift+F4, documented acCmdFindNext (341) and SendKeys
reuse it without reopening the dialog. Real Qt/VBA/SQL regression verifies
closed-dialog advance/wrap, active-document routing from a hidden caller,
no-history diagnostics and validation cancellation preserving draft/SQL/position.
Targeted find_form_documents passes (0.13 s). Full/shared Access search-history
and Find/Replace semantics remain open. No business execution increment.
See FIND_FORM_RUNTIME.md.

Full validation: host 27/27 (99.04 s), Debian 12 27/27 (79.78 s), Ubuntu
24.04 27/27 (71.50 s), Debian 13 27/27 (73.47 s), GUI-disabled core 1/1
(1.64 s). Host/all three distro DEBs pass payload/backend/permissions/fixture-
exclusion checks. Isolated installations pass backend execution and GUI startup.
All four original fixture hashes pass; repository remains private. Previous
bcb99e8 private CI completed successfully.

## 2026-11-02 — Checked VBA Byte subtype and conversion

Closed an accepted-but-unchecked Byte declaration path. CByte/modules/locals/
parameters/results now share bounded conversion and explicit Byte type identity;
known ByRef mismatch and overflow fail without losing the previous value.
Native Qt/VBA/DAO/SQL tests exercise actual calculated controls, Byte query
parameters, stored integer values and explicit record save. Core/native targeted
regression passes 2/2 (0.96 s); GUI suite now has 28 CTests. Full promotion,
arrays/For counters/locales/Windows oracle remain open. No inspected original
module uses Byte/CByte, so no original business execution increment is claimed.
See BYTE_VALUES_RUNTIME.md.

Full validation: host 28/28 (103.95 s), Debian 12 28/28 (88.28 s), Ubuntu
24.04 28/28 (82.37 s), Debian 13 28/28 (79.53 s), GUI-disabled core 1/1
(2.89 s). Host/all three distro DEBs pass payload/backend/permissions/fixture-
exclusion checks. Isolated installations pass backend execution and GUI startup.
All four original fixture hashes pass; repository remains private. Previous
3bcb787 private CI completed successfully.

## 2026-11-02 — Typed native Not bit inversion

Reproduced wrong result type/Byte value in independent core and native form
regressions (0/2 before fix). Not now respects 8/16/32-bit Byte/Integer/Long,
Boolean/Null/Empty and one scalar default evaluation. Native Qt/VBA/SQL writes
actual inverted Byte 251; range errors prevent later SQL and explicit source
record-save behavior is preserved. Targeted tests pass 2/2 (1.08 s). Full mixed
promotion/text/locale/Jet SQL/Windows semantics remain open. Original business
execution coverage is unchanged. See TYPED_NOT_RUNTIME.md.

Full validation: host 28/28 (86.11 s), Debian 12 28/28 (84.74 s), Ubuntu
24.04 28/28 (78.45 s), Debian 13 28/28 (81.28 s), GUI-disabled core 1/1
(1.70 s). Host/all three distro DEBs pass payload/backend/permissions/fixture-
exclusion checks. Isolated installations pass backend execution and GUI startup.
All four original fixture hashes pass; repository remains private. Previous
7badc0f private CI is still running at this checkpoint; no CI success is claimed.

## 2026-11-02 — Integer return subtype for VarType

Reproduced core and actual calculated-form failures from returning supported
VarType codes as Double (0/2 before fix). Native VarType now returns Integer;
nested inspection, Null/Error/Byte-array codes, Variant copy, calculated display
and actual SQL capture pass 2/2 (1.14 s). Full object-default inference, arrays,
Date/Decimal and Windows oracle remain open. No original business execution
increment is claimed. See VARTYPE_RESULT_RUNTIME.md.

Full validation: host 28/28 (85.88 s), Debian 12 28/28 (84.76 s), Ubuntu
24.04 28/28 (75.57 s), Debian 13 28/28 (78.50 s), GUI-disabled core 1/1
(1.59 s). Host/all three distro DEBs pass payload/backend/permissions/fixture-
exclusion checks. Isolated installations pass backend execution and GUI startup.
All four original fixture hashes pass; repository remains private. 7badc0f
private CI has now completed successfully; 6d00c90 CI is still running.

## 2026-11-02 — Inspect native scalar default value types

Reproduced Object-only VarType results against independent core/native tests
(0/2). Advertised native scalar defaults now expose actual control/DAO/Err value
types; getter failures propagate and recursion is bounded. Err.Number retains
Long. New native Qt/VBA/SQL CTest stores actual observed codes and distinguishes
Null, EOF 3021 and closed DAO field/parameter 3420, with no data/dirty changes.
The regression exposed generic closed-recordset errors, now aligned with native
3420 guards. Targeted core/native/DAO passes 3/3 (1.38 s); GUI suite has 29 CTests.
Full class/default/type/storage/Windows semantics remain open. No original
business execution increment. See DEFAULT_TYPE_INSPECTION_RUNTIME.md.

Full validation: host 29/29 (106.37 s), Debian 12 29/29 (82.32 s), Ubuntu
24.04 29/29 (75.11 s), Debian 13 29/29 (75.91 s), GUI-disabled core 1/1
(3.10 s). Host/all three distro DEBs pass payload/backend/permissions/fixture-
exclusion checks. Isolated installations pass backend execution and GUI startup.
All four original fixture hashes pass; repository remains private. Previous
6d00c90 and 0fd9b43 private CI completed successfully.

## 2026-11-02 — Native character functions and original KeyPress execution

Added native ASCII Asc/Chr and supported UTF-16 AscW/ChrW with checked subtype/
scalar/range/Null/unsupported-codepage handling. Core and actual bound/calculated
Qt/VBA/DAO/SQL round-trip/pending-save tests pass 2/2 (1.15 s). New original
regression SHA-verifies unchanged module bytes and executes three KeyPress
procedures over 384 direct ASCII and 285 printable Qt cases on prepared native
controls. Actual field effects/parameter outcomes verify wiring. access_2003_layouts passes
(12.61 s). GUI suite now has 30 CTests. No full original form/business or Windows
oracle claim; ANSI/DBCS/full BSTR/SQL/function/keyboard limits remain explicit.
See CHARACTER_FUNCTIONS_RUNTIME.md.

Full validation: host 30/30 (97.61 s), Debian 12 30/30 (94.31 s), Ubuntu
24.04 30/30 (83.33 s), Debian 13 30/30 (82.17 s), GUI-disabled core 1/1
(1.58 s). Host/all three distro DEBs pass payload/backend/permissions/fixture-
exclusion checks. Isolated installations pass backend execution and GUI startup.
All four original fixture hashes pass; repository remains private. Previous
a357ed9 private CI completed successfully.

## 2026-11-02 — Declared native integer field types

Recovered Byte/Integer/Long identity from unambiguous imported integer metadata,
with checked scalar writes and bounded readonly Field.Type. Native Qt/VBA/DAO/
SQL tests cover range preservation, rounding, pending/save, generic 53-bit data
and invalid stored values. Targeted tests pass 3/3 (6.98 s). Strengthened original
MDB regression verifies all 127 INT columns and 471904 non-NULL Long field values
against independent SQL integer-row/NULL multisets; full CSV cell comparison
remains active (7.58 s). GUI suite has 31 CTests. MySQL FLOAT mixes Currency/
Single, so no blanket floating reclassification is made. Full raw SQL/Jet/type/
DAO/Windows identity remains open. See INTEGER_FIELD_TYPES_RUNTIME.md.

Full validation: host 31/31 (90.79 s), Debian 12 31/31 (94.73 s), Ubuntu
24.04 31/31 (83.44 s), Debian 13 31/31 (84.25 s), GUI-disabled core 1/1
(1.91 s). Host/all three distro DEBs pass payload/backend/permissions/fixture-
exclusion checks. Isolated installations pass backend execution and GUI startup.
All four original fixture hashes pass; repository remains private. Previous
94f6302 private CI completed successfully.

## 2026-11-02 — Shared native bankers Round

Reproduced incorrect Access-SQL half rounding via SQLite (0/1 before fix).
The SQL rewrite/callback and VBA now share native C++ rounding, with bounded
precision/subtypes/scalar/error behavior and exact explicit Currency quantization.
Core/actual Qt/VBA/QueryDef/SQL tests pass 2/2 (1.00 s), including Microsoft
examples, pending/save, exact large fractional Currency TEXT and failed mixed
UPDATE without partial changes. GUI suite has 32 CTests. Single/high precision,
full binary/type/locale/Jet money/Windows identity remain open. No additional
original business workflow is claimed. See ROUND_RUNTIME.md.

Full validation: host 32/32 (104.95 s), Debian 12 32/32 (97.80 s), Ubuntu
24.04 32/32 (86.23 s), Debian 13 32/32 (86.65 s), GUI-disabled core 1/1
(2.74 s). Host/all three distro DEBs pass payload/backend/permissions/fixture-
exclusion checks. Isolated installations pass backend execution and GUI startup.
All four original fixture hashes pass; repository remains private. Previous
446b0c5 private CI completed successfully.

## 2026-11-02 — Native form visibility and hidden-document context

Added explicit Form.Visible for actual activated Qt windows, preserving loaded
hidden forms and drafts on show. Explicit hide clears active context; formerly
activated hidden callers cannot receive default record/unnamed Close without an
active target. Named Close and hidden project-close Unload cancellation still
work. New actual Qt/VBA/SQL test verifies no implicit save/close/finished,
scalar-once/Null, pending-hide rejection, dirty show, hidden command/SendKeys
failure and self-close during Deactivate. Targeted visibility/mainwindow/original
layout/Kassa-DAO passes 4/4 (48.48 s); GUI suite has 33 CTests. Init/modal/dirty-hide/
embedded/report/full focus/MDI/Windows identity remains open. No complete original
business workflow increment. See FORM_VISIBILITY_RUNTIME.md.

Full validation: host 33/33 (115.68 s), Debian 12 33/33 (100.31 s), Ubuntu
24.04 33/33 (90.73 s), Debian 13 33/33 (90.20 s), GUI-disabled core 1/1
(3.14 s). Host/all three distro DEBs pass payload/backend/permissions/fixture-
exclusion checks. Isolated installations pass backend execution and GUI startup.
All four original fixture hashes pass; repository remains private. Previous
0eb9501 private CI completed successfully.

## 2026-11-02 — Native Ctrl+Page Down/Up record navigation

Reproduced missing native navigation/validation via a failing actual Qt/VBA/SQL
regression. The shortcuts now use existing next/previous record commands,
preserving focus, KeyDown cancellation, field/record validation, SQL save,
Current and lifetime guards. Tests cover text/ComboBox focus, SendKeys both ways,
no plain Page/Ctrl+Shift record shortcut, first/last/new boundaries, no empty
insert, disabled additions and self-close in Current. Targeted new/visibility/
original-layout/Kassa-DAO regressions pass 4/4 (49.62 s). GUI suite has 34 CTests.
Multi-page scrolling/other views/full Windows keyboard identity remain open.
No additional complete original business workflow. See RECORD_NAVIGATION_KEYS_RUNTIME.md.

Full validation: host 34/34 (93.76 s), Debian 12 34/34 (104.68 s), Ubuntu
24.04 34/34 (90.84 s), Debian 13 34/34 (91.94 s), GUI-disabled core 1/1
(1.55 s). Host/all three distro DEBs pass payload/backend/permissions/fixture-
exclusion checks. Isolated installations pass backend execution and GUI startup.
All four original fixture hashes pass; repository remains private. Previous
0c4ed45 private CI was still in progress at this checkpoint.

## 2026-11-02 — Native VBA Replace

Reproduced unavailable Replace via failing core regression (0/1). Added native
UTF-16 suffix/non-overlap/count/compare handling with optional omissions,
scalar-once inputs and checked errors. Core and actual Qt/VBA/DAO parameter/SQL
regressions pass 2/2 (1.13 s), covering pending/save/Undo, Unicode/module
comparison and failed arguments without SQL effects. Additional tests cover
numeric rounding/types and scalar-once optional arguments. GUI suite has 35
CTests. Full Windows collation/BSTR/error identity and Jet SQL Replace remain
open; no additional complete original business workflow. See REPLACE_RUNTIME.md.

Full validation: host 35/35 (93.88 s), Debian 12 35/35 (98.10 s), Ubuntu
24.04 35/35 (86.33 s), Debian 13 35/35 (87.27 s), GUI-disabled core 1/1
(1.55 s). Host/all three distro DEBs pass payload/backend/permissions/fixture-
exclusion checks. Isolated installations pass backend execution and GUI startup.
All four original fixture hashes pass; repository remains private. Visibility
commit 0c4ed45 private CI completed successfully; 100e57c was still in progress.

## 2026-11-02 — Native string-function result types

Reproduced Double-versus-Long mismatch in core regression (0/1). Supported Len
string measurements/InStr now return checked Long, StrComp Integer; Null remains
Null. Core and actual Qt/VBA/DAO parameter/SQL type-code regressions pass 2/2
(1.04 s), including Unicode positions, nested inspection, pending/save and Null
Undo without source changes. GUI suite has 36 CTests. Numeric/UDT Len storage,
complete locale/BSTR/Windows and SQL result subtype propagation remain open.
No additional complete original business workflow. See STRING_RESULT_TYPES_RUNTIME.md.

Full validation: host 36/36 (101.85 s), Debian 12 36/36 (101.77 s), Ubuntu
24.04 36/36 (87.96 s), Debian 13 36/36 (92.16 s), GUI-disabled core 1/1
(1.77 s). Host/all three distro DEBs pass payload/backend/permissions/fixture-
exclusion checks. Isolated installations pass backend execution and GUI startup.
All four original fixture hashes pass; repository remains private. Navigation
commit 100e57c private CI completed successfully; 3232150 was still in progress.

## 2026-11-02 — Native VBA IsObject

Reproduced unavailable intrinsic in core regression (0/1). Added Boolean object-
reference inspection and argument parser object context without scalar defaults.
Nothing/closed stored references remain True; explicit invalid member getters
still fail. Core and actual Qt/VBA/DAO parameter/SQL tests pass 2/2 (1.30 s),
including forms/controls/collection/Err/recordsets/fields/closed/Set Nothing,
3420 before SQL effects and pending text/source/Undo preservation. GUI suite has
37 CTests. Full classes/COM/ADO/object-model/Windows identity remain open. No new
complete original business workflow. See ISOBJECT_RUNTIME.md.

Full validation: host 37/37 (103.25 s), Debian 12 37/37 (105.47 s), Ubuntu
24.04 37/37 (90.26 s), Debian 13 37/37 (94.48 s), GUI-disabled core 1/1
(1.71 s). Host/all three distro DEBs pass payload/backend/permissions/fixture-
exclusion checks. Isolated installations pass backend execution and GUI startup.
All four original fixture hashes pass; repository remains private. Replace
commit 3232150 private CI completed successfully earlier in this segment;
7bf428a was still in progress at the last check.

## 2026-11-02 — Currency/Single/Double source metadata

Analysis of a hash-fixed private test copy established that all 43 original
MySQL FLOAT columns are Currency. Reproduced missing Currency metadata in the
original-table regression (0/1, 0/43 labels). Native import now reads MDBTools'
Access backend and checks exact Currency/Single/Double mappings against known columns;
source metadata is persisted without guessing ambiguous FLOAT as Single.
Synthetic parsing/persistence and original-table tests pass 2/2 (7.81 s), retaining
all CSV cell, Long/Boolean/date/counter checks. GUI suite remains 37 CTests.
SQLite REAL/CSV conversion/native numeric storage is not made exact by labels;
full Currency/Single/SQL/precision and existing-project migration remain open.
No new dependency or complete original business workflow. See JET_FLOATING_METADATA.md.

Distribution validation exposed older MDBTools also exporting 57 Double columns
as FLOAT (100 ambiguous columns total). Extended the checked Access-backend
mapping to Double; synthetic and original regressions pass 2/2 (7.81 s). Original
metadata assertions now require both 43 Currency and 57 Double columns. No source
data or precision is inferred from the MySQL FLOAT label.

Corrected full validation: host 37/37 (108.10 s), Debian 12 37/37 (118.21 s),
Ubuntu 24.04 37/37 (100.17 s), Debian 13 37/37 (99.85 s), GUI-disabled core
1/1 (2.47 s). Host/all three distro DEBs pass payload/backend/permissions/fixture-
exclusion checks. Isolated installations pass backend execution and GUI startup.
All four original fixture hashes pass; repository remains private. Previous
8dba067 private CI completed successfully. Initial distribution failures were
resolved by handling older MDBTools' additional Double-to-FLOAT mapping.

## 2026-11-02 — Bounded native Currency fields

Reproduced Double field values despite verified Currency metadata (0/1). Added
checked binary64-to-int64 Currency tick recovery, bounded strictly inside +/-2^39,
canonical round-trip validation, DAO dbCurrency and exact native text formatting.
Core decimal-reference/boundary/10000-value tests and original layout/Kassa-DAO/
Currency-column regressions pass (4/4 in the initial targeted run). All 121613
non-NULL original Currency field values match independent CSV decimal row/NULL
multisets across 43 columns. Corrected actual Qt/VBA/DAO/SQL test passes 1/1
(0.09 s), including fixed arithmetic/types, pending/save, range checks, DAO
Edit/Update, large display/OldValue/Undo and copy/new-row without SQL insertion.
GUI suite has 38 CTests. Full-range/SQL money/parameters/query subtype/locale/
Windows semantics remain open; no full original business workflow increment.
See CURRENCY_FIELDS_RUNTIME.md.

Full validation: host 38/38 (128.74 s), Debian 12 38/38 (103.69 s), Ubuntu
24.04 38/38 (87.32 s), Debian 13 38/38 (91.35 s), GUI-disabled core 1/1
(2.09 s). Host/all three distro DEBs pass payload/backend/permissions/fixture-
exclusion checks. Isolated installations pass backend execution and GUI startup.
All four original fixture hashes pass; repository remains private. Previous
03704c9 private CI completed successfully.

## 2026-11-02 — Bounded Currency/Money query parameters

Reproduced unavailable Currency parameter declarations (0/1). Added typed fixed
Currency Value/dbCurrency, scalar-once and checked failure retention, numeric
binding through the lossless +/-2^39 REAL interval and Money alias. Actual native
DAO/SQL and bound Qt/VBA UPDATE/requery regressions pass 3/3 (0.33 s), checking
large inserts/recovery, numeric-versus-text comparison, rounding/Null/missing/
readonly/lifetime and failed assignments without extra effects. GUI suite has
39 CTests. Full-range/SQL money arithmetic/formatting/result typing/LIKE/Windows
identity remain open. No full original business workflow increment. See
CURRENCY_PARAMETERS_RUNTIME.md.

Full validation: host 39/39 (106.15 s), Debian 12 39/39 (110.08 s), Ubuntu
24.04 39/39 (95.17 s), Debian 13 39/39 (98.20 s), GUI-disabled core 1/1
(1.87 s). Host/all three distro DEBs pass payload/backend/permissions/fixture-
exclusion checks. Isolated installations pass backend execution and GUI startup.
All four original fixture hashes pass; repository remains private. Previous
3d5606b private CI completed successfully.

## 2026-11-02 — Bounded native SQL CCur

Reproduced unavailable SQL converter (0/1). Added SQL rewrite/registration and
shared native Currency conversion with bounded REAL output and typed errors.
Actual converter/parameter/field regressions pass 3/3 (0.16 s), including decimal
rounding, booleans/numeric expressions/large values, literals/aliases, invalid
arguments, DAO multi-row action rollback and actual bound Qt/VBA UPDATE/requery.
GUI suite has 40 CTests. Full query Currency subtype propagation, full-range/
regional/SQL arithmetic/aggregate/Windows identity remain open. No additional
complete original business workflow. See CURRENCY_SQL_RUNTIME.md.

Full validation: host 40/40 (100.19 s), Debian 12 40/40 (108.66 s), Ubuntu
24.04 40/40 (93.62 s), Debian 13 40/40 (97.47 s), GUI-disabled core 1/1
(1.93 s). Host/all three distro DEBs pass payload/backend/permissions/fixture-
exclusion checks. Isolated installations pass backend execution and GUI startup.
All four original fixture hashes pass; repository remains private. Previous
d1e3ac1 private CI completed successfully.

## 2026-11-02 — Currency SQL references and proven query result types

Reproduced rejected Currency SQL object scalars and missing query Field.Type
(0/2). Added checked numeric object encoding and proven root CCur/Currency-
parameter projection descriptors, carried through results into DAO/Qt controls
and exact query-table display. Filters preserve descriptors; ambiguous/wildcard/
set operations keep unknown types. Native object/result/converter/parameter/
field regressions pass 5/5 (0.25 s); extended report/saved-query/TOP/empty tests
pass 2/2 (0.12 s). Checks include accepted pending Value/no save, ranges, numeric
comparison, Forms/Formulare/Reports/Berichte/.Value, Nullable/EOF, parentheses/
unary signs, DISTINCT, filters and actual readonly Qt display. GUI suite has
42 CTests. General SQL typing/object grammar/SQL Screen/locale/money arithmetic/
Windows identity remain open; no full original business workflow increment.
See CURRENCY_QUERY_TYPES_RUNTIME.md.

Full validation: host 42/42 (120.79 s), Debian 12 42/42 (103.74 s), Ubuntu
24.04 42/42 (89.02 s), Debian 13 42/42 (89.21 s), GUI-disabled core 1/1
(1.89 s). Host/all three distro DEBs pass payload/backend/permissions/fixture-
exclusion checks. Isolated installations pass backend execution and GUI startup.
All four original fixture hashes pass; repository remains private. Previous
ca35ddd private CI completed successfully.

## 2026-11-02 — Exact sums of proven Currency expressions

Reproduced residual/generic SUM result for explicit Currency input (0/1). Added
checked integer-unit aggregate and bounded projection rewrite/type propagation.
Sum/query/object regressions pass 3/3 (0.16 s), covering grouping/Null/empty,
Currency parameters, large cancellation, final range/int64 overflow and actual
Qt report display. All 43 original Currency-column sums match independent wide-
integer CSV decimal sums, with zero outside backing range (7.91 s); all 121613
individual value and existing CSV/Long/Boolean/date/counter checks remain active.
GUI suite has 43 CTests. Implicit field/view/complex aggregate inference and full
SQL money/parser/locale/Windows semantics remain open; no full original business
workflow increment. See CURRENCY_SUM_RUNTIME.md.

Full local validation: host 43/43 (104.29 s), Debian 12 43/43 (103.88 s),
Ubuntu 24.04 43/43 (90.98 s), Debian 13 43/43 (92.46 s), GUI-disabled core
1/1 (1.86 s). Host/all three distro DEBs pass payload/backend/permissions/fixture-
exclusion checks. Isolated installations pass backend execution and GUI startup.
All four original fixture hashes pass; repository remains private.
Remote e9d0f60 CI did not start any job steps: GitHub reported an account-level
billing/spending restriction. This is an external CI blocker, not a successful
remote test run; local distribution checks above completed independently.

## 2026-11-02 — Verified Currency columns in simple table queries

Reproduced missing implicit query type (0/1). Added bounded live physical-table
Currency resolution, qualified/aliased fields, exact implicit Sum(column), result
metadata through DAO/readonly Qt and filter propagation. SQL changes clear stale
hints; previously resolved source metadata refreshes before reuse. Initial
Currency tests pass 4/4 (0.21 s); expanded query/object/sum/original-layout/Kassa/
table run passes, with final full-suite rebuild below. All 43 original sums now
use plain Sum(column) and match independent CSV integers; all 121613 individual
Currency comparisons remain active. GUI suite has 44 CTests. General/view/join/
nested/HAVING/action/complex SQL typing and full money/Windows semantics remain
open; no complete original business workflow increment. See
CURRENCY_SCHEMA_QUERIES_RUNTIME.md.

Full local validation: host 44/44 (122.41 s), Debian 12 44/44 (108.77 s),
Ubuntu 24.04 44/44 (88.35 s), Debian 13 44/44 (94.54 s), GUI-disabled core
1/1 (2.28 s). Host/all three distro DEBs pass payload/backend/permissions/fixture-
exclusion checks. Isolated installations pass backend execution and GUI startup.
All four original fixture hashes pass; repository remains private. GitHub's
previously confirmed account-level CI restriction remains separate from these
successful local distribution checks; no remote test success is claimed.

## 2026-11-02 — Exact arithmetic of proven Currency query expressions

Reproduced generic Double evaluation for Currency operands (0/1). Added bounded
projection expression parsing and native fixed-point +,-,* callbacks using shared
checked Currency arithmetic. Arithmetic/sum/query/schema/object tests pass 5/5
(0.28 s), covering precedence/associativity/parentheses/signs, four-decimal
multiplication ties, Null/parameters, Sum(a+b), overflow/backing limits and actual
readonly Qt report output without source writes. GUI suite has 45 CTests.
Mixed/division/complex/view/join/action/full-range/Windows semantics remain open;
no complete original business workflow increment. See
CURRENCY_ARITHMETIC_SQL_RUNTIME.md.

Full local validation: host 45/45 (105.25 s), Debian 12 45/45 (108.67 s),
Ubuntu 24.04 45/45 (95.36 s), Debian 13 45/45 (95.63 s), GUI-disabled core
1/1 (1.72 s). Host/all three distro DEBs pass payload/backend/permissions/fixture-
exclusion checks. Isolated installations pass backend execution and GUI startup.
All four original fixture hashes pass; repository remains private. GitHub's
account-level CI restriction remains separate; no remote test success is claimed.

## 2026-11-02 — Checked Currency targets in simple UPDATE

Reproduced unchecked Double assignment in the new regression (0/1). Added
bounded physical-table UPDATE token handling, live Currency metadata, original
parameter bindings, checked nullable target conversion and reuse of proven
fixed-point RHS arithmetic. Initial regression passes; expanded tests verify
aliases/qualified names, simultaneous assignments, ordinary columns, ordered
parameters, Null, text rounding, range errors and statement-wide rollback.
A mixed Double half-value assertion was invalid because binary evaluation
precedes target conversion; replaced it with an unambiguous mixed-value check,
retaining exact text/Currency rounding checks. Final targeted UPDATE test passes
(1/1); the other 10 selected Currency/layout/original Kassa tests passed in the
expanded run. Native VBA/DAO/Qt update/requery and exact boundary display checked.
GUI suite has 46 CTests. Joined/append/complex/full Jet/Windows semantics remain
open; no complete original business workflow increment. See
CURRENCY_UPDATE_RUNTIME.md. Full local validation follows below.

Full local validation: host 46/46 (105.91 s), Debian 12 46/46 (111.22 s),
Ubuntu 24.04 46/46 (96.44 s), Debian 13 46/46 (97.80 s), GUI-disabled core
1/1 (1.32 s). Host/all three distro DEBs pass payload/backend/permissions/fixture-
exclusion checks. Isolated installations pass backend execution and GUI startup.
All four original fixture hashes pass; repository remains private. GitHub's
account-level CI restriction remains separate; no remote test success is claimed.

## 2026-11-02 — Checked Currency targets in single-row INSERT VALUES

New regression reproduces missing target conversion (0/1). Added bounded
physical-table INSERT VALUES conversion with explicit/positional/reordered
columns, original parameter bindings, omitted/Null values and exact proven
Currency expressions. INSERT/UPDATE regressions pass 2/2 (0.12 s). Invalid
second-column values insert no row; malformed incomplete VALUES remains a syntax
error. Real VBA/DAO insertion, form requery/navigation, Currency VarType/units and
Qt display pass. Test initially referenced an unavailable native CurrentRecord
member; corrected the assertion to the inserted field identity instead. That
separate API gap remains open. GUI suite has 47 CTests. Append/complex/default/
constraint/full Jet/Windows semantics remain open; no complete original business
workflow increment. See CURRENCY_INSERT_RUNTIME.md. Full validation follows.

Full local validation: host 47/47 (97.83 s), Debian 12 47/47 (103.93 s),
Ubuntu 24.04 47/47 (91.06 s), Debian 13 47/47 (91.90 s), GUI-disabled core
1/1 (1.31 s). Host/all three distro DEBs pass payload/backend/permissions/fixture-
exclusion checks. Isolated installations pass backend execution and GUI startup.
All four original fixture hashes pass; repository remains private. GitHub's
account-level CI restriction remains separate; no remote test success is claimed.

## 2026-11-02 — Bound Form.CurrentRecord and Currency append batch

Added read-only Long CurrentRecord for supported bound forms, typed native
Me/Forms/Screen/unqualified/calculated-control lookup, current-event positions,
filtered navigation/new-row ordinals and no-save/read-only/lifecycle safeguards.
Regression reproduced missing member (0/1), then passes 1/1 (0.14 s). Public
Microsoft property contract inspected; no Windows differential run available.
Unbound/design/report/unknown new-row ordinals explicitly remain outside scope.
See CURRENT_RECORD_RUNTIME.md.

Extended the existing Currency INSERT regression (no additional executable) to
multi-row physical-source INSERT SELECT, checked target conversion, frozen
positional source rows, exact self-append arithmetic, original parameter order
and atomic late-row failure. Reproduced unchecked append, then combined
CurrentRecord/INSERT tests pass 2/2 (0.13 s). Complex join/TOP/view/nested/HAVING/
set-operation/default/constraint/full Jet append remains open. See
CURRENCY_INSERT_RUNTIME.md. GUI suite has 48 CTests. User requested less repeated
verification overhead: feature work is now batched, with focused checks during
implementation and a shared full delivery validation before push. No full
original business workflow increment. Full local validation follows.

Shared delivery validation: host 48/48 (84.12 s), GUI-disabled core 1/1;
host DEB payload/backend/build-ID/permissions/fixture-exclusion checks pass.
All four original hashes pass; repository remains private. To reduce repeated
validation overhead as requested, the distribution matrix was not rerun for
this batch: last fully checked three-distribution state is 4225a17 (47/47 on
each platform plus packages/isolated installations). No current-batch distro
or remote CI success is claimed; the GitHub account restriction remains external.

## 2026-11-02 — Shared native Gregorian calendar function batch

Added shared C++20 ISO/Gregorian DateDiff, Year/Month/Day/Hour/Minute/Second and
Weekday, native SQL callbacks/compiler/alias integration, weekday/week constants
and optional defaults. Focused core/query tests pass 2/2 (1.04 s), including real
VBA procedures, Integer/Long VarType, scalar getter once, leap dates, boundary
counting/negative results, week starts, Null, errors and Long overflow. No new
test target: GUI suite remains 48 CTests. Public Microsoft contracts read;
Windows differential oracle remains unavailable. Date Variants, numeric/locale/
regional/Hijri/fractional-second/coercion/w/non-default ww-first-week/full calendar
semantics remain open. No complete original business workflow increment. See
CALENDAR_FUNCTIONS_RUNTIME.md. Shared validation follows below.

Shared delivery validation: host 48/48 (82.73 s), GUI-disabled core 1/1
(1.00 s), host DEB payload/backend/build-ID/permissions/fixture-exclusion checks
pass. Four original hashes and private repository visibility rechecked. The
three-distribution matrix remains at 4225a17; it was not repeated for this batch.
No current-batch distribution or remote CI success is claimed.

## 2026-11-02 — Native form ordering and active-field sort commands

Added OrderBy/OrderByOn with token-validated output-field lists, ASC/DESC,
verified direct-table qualification, imported initial flags, native table/query
sorting and rollback/validation safeguards. Physical table row identities survive
reordering; query sources without row identities preserve ordinal only. Existing
CurrentRecord/calculated-control/Current-event behavior stays connected. Added
unqualified VBA property routing and active-document DoCmd sort commands 163/164
(verified against Microsoft's AcCommand enum). Invalid/unbound/calculated/binary
focused-field commands are guarded; full sort/collation/report/design/Windows
semantics remain open. Existing form_current_record regression first reproduces
missing properties (0/1), then passes with imported/query/VBA/qualified/sort-command
coverage. No new test target: suite remains 48 CTests. No complete original
business workflow increment. See FORM_ORDER_RUNTIME.md. Shared validation follows.

Shared delivery validation: host 48/48 (81.15 s), GUI-disabled core 1/1
(1.06 s), host DEB payload/backend/build-ID/permissions/fixture-exclusion checks
pass. Four original hashes and private visibility rechecked. Three-distribution
matrix remains at 4225a17; no current-batch distro/remote CI success is claimed.

## 2026-11-02 — Actual app table/query editor batch

Replaced raw app table model with checked source-aware AccessTableModel and a
native text delegate preserving Currency precision/long Memo contents. Generated
native/managed AutoNumbers are read-only, hidden behind (AutoWert) while pending,
and omitted from inserts to allocate through existing defaults. Shared assignment
conversion checks source integer ranges, Currency, Boolean and dates; rejected
edits preserve cells and emit generic typed feedback. Binary edits remain blocked.
The custom SQL dialog now uses the actual Access compiler/model, typed declared
parameter prompts and DAO action execution/affected rows; action definitions can
be saved. Unsupported grammar remains diagnosed, not treated as full Jet support.

Existing field/main-window integration tests pass 2/2 (1.68 s): actual Currency
line editor/max value/rounding/rejection, 50000+1 Memo characters, date/Integer
edits, SQLite and managed counters, actual SQL dialog CCur/DateDiff execution,
DAO insertion, Currency parameter prompt and action save. No new target: GUI
suite remains 48 CTests. Full datasheet/designer/formatting/constraints/OLE/Jet/
Windows identity remains open. No original business workflow increment. See
APP_EDITORS_RUNTIME.md. Shared delivery and refreshed distribution validation follow.

Shared full validation before the final sort guard: host 48/48 (115.45 s),
Debian 12 48/48 (104.07 s), Ubuntu 24.04 48/48 (89.18 s), Debian 13 48/48
(92.37 s), GUI-disabled core 1/1 (1.49 s); packages and isolated installations pass.
Found/reproduced Qt's sort/select clearing pending manual edits (new regression
0/1). Added model/view sorting protection for cached changes and a live editor,
restoring the header indicator. Final focused editor tests pass on host 2/2
(1.65 s), Debian 12 2/2 (2.63 s), Ubuntu 24.04 2/2 (2.10 s), Debian 13 2/2
(1.85 s). All four final packages rebuilt/checked; final isolated distro installs
pass backend execution/GUI startup. Other 46 tests were not redundantly rerun
for this isolated guard. Four original fixture hashes pass; private visibility
rechecked. Original query counts remain 72 ordinary + 5 parameterized executed,
13 binding/Jet boundaries, 9 missing dependencies and 1 source-column mismatch;
no Access reference oracle or full business-workflow claim. Remote CI remains
blocked by the previously confirmed account restriction.

## 2026-11-02 — Native datasheet checkbox/keyboard completion batch

Added checked Boolean checkbox roles and Space toggling with -1/0 persistence,
plus table Ctrl+S invoking manual-save transactions only after checked live-cell
completion. Delegate persistent indices identify the edited cell; invalid values
remain in the editor, valid values reach the model before reset/save. Existing
native field/main-window paths pass; no new CTest target (suite remains 48).
A checkbox test assumed unsorted row order; corrected it to locate record identity.
The main-window shortcut fixture initially referenced a prior project table;
now creates its own table via the real SQL editor before opening it in explorer.
Actual Space/Ctrl+S, pending database value retention, final persistence and
invalid Currency editor retention are checked. Full Yes/No defaults/Null, per-row
save events/undo/keyboard/Windows identity remain open. See APP_EDITORS_RUNTIME.md.
No complete original business workflow increment. Delivery validation follows.

Delivery: native field regression passed (0.26 s), final main-window regression
passed 1/1 (1.43 s); current host app rebuilt and DEB payload/backend/build-ID/
permissions/fixture-exclusion checked. Unaffected suite/platform runs were not
repeated: the previous shared matrix is documented under 644e71d. Four original
hashes and private visibility rechecked. No new full-suite/distro/remote CI pass
is claimed for this focused GUI-only delta; account CI restriction remains open.

## 2026-11-02 — Pending datasheets in native document close preflight

Added an app-independent native project close participant and table-page guards,
covering tab/window/project switch and application/VBA Quit before form unloads.
Changed live editor/cached values remain untouched; unmodified editor values do
not falsely block. Reverting changes allows normal close. Existing main-window/
field regressions pass 2/2 (1.82 s), including actual native VBA Quit 2501 with
retained caller and no pending data loss. The first alternate-project fixture
used a nonexistent path, causing the file dialog watchdog; corrected it to an
existing private test copy. No new test target (suite remains 48). Full automatic
record saving/events/undo/crash/window/Windows identity remains open; no complete
original business workflow increment. See APP_EDITORS_RUNTIME.md. Shared host
validation follows because the common document close path changed.

Shared host delivery validation: 48/48 (83.72 s), rebuilt host DEB passes payload/
backend/build-ID/permissions/fixture-exclusion checks. All four original hashes
and private visibility pass. GUI-disabled core sources are unchanged from their
last checked state. The three-distribution matrix remains at 644e71d; no fresh
cross-distribution or remote CI success is claimed for this batch. GitHub's
account restriction remains separate from successful local checks.

## 2026-11-02 — Guarded datasheet close shortcuts

Added native page-scoped Ctrl+W/Ctrl+F4 through the existing close preflight and
page-wide Ctrl+S. Accepted tab close releases model queries and defers page/widget
deletion until the shortcut sender finishes. Actual main-window key regression
passes 1/1 (1.45 s): pending Ctrl+W refuses without loss, reverted Ctrl+F4 closes.
No new CTest target or broad suite repetition for this focused UI delta. Full
MDI/window/keyboard identity remains open. See APP_EDITORS_RUNTIME.md. Host
package validation follows; prior shared host suite is 6dc8a97 (48/48).

Focused delivery: current host app rebuilt and DEB payload/backend/build-ID/
permissions/fixture-exclusion checks pass. Four original hashes/private visibility
pass. No repeated unaffected suite/distro run; shared host suite remains 6dc8a97
and three-distribution baseline 644e71d. Remote CI account restriction persists.

## 2026-11-02 — Checked DAO count subtypes and schema action counts

Reproduced stale SQLite DDL RecordsAffected (0/1 after selecting the correct
existing dao_execution test). Added zero CREATE/ALTER/DROP counts while preserving
DML/joined-action and make-table counts. Database/QueryDef.RecordsAffected and
loaded Recordset.RecordCount now use checked Long; Fields.Count uses checked
Integer (public Microsoft contract read). Overflow is 6. Existing DAO/main-window
tests pass 2/2 (1.79 s), covering actual DDL/DML/SELECT INTO counts, subtypes and
SQL-editor zero-row feedback. No new CTest target (suite stays 48). Full cursor/
metadata/error/Windows identity remains open; no original business workflow
increment. See DAO_COUNTS_RUNTIME.md. Focused delivery validation follows.

Focused delivery: host app rebuilt; DEB payload/backend/build-ID/permissions/
fixture-exclusion checks pass. Four fixture hashes/private visibility pass.
Unchanged core and unaffected full-suite/distribution checks were not repeated;
shared host baseline is 6dc8a97, matrix baseline 644e71d. Remote CI restriction
remains external; no remote test success is claimed.

## 2026-11-02 — Checked source scalar targets in simple SQL writes

Reproduced unchecked typed targets (0/1). Generalized the existing bounded
Currency write parser to source Byte/Integer/Long/Boolean/date targets, retaining
Currency's separate exact converter and arithmetic. Added shared scalar callback,
Null handling, canonical Boolean/date storage and 0100–9999 date validation for
text/QDateTime read/conversion. Existing DAO/Currency insert/update tests pass
3/3 (0.42 s), including banker ties/range errors, invalid dates, late-row UPDATE/
append rollback, declared Double parameters and ordinary unmarked SQLite behavior.
No new target (suite remains 48). Joined/complex/full Jet expression/constraint/
Date Variant/locale/Windows semantics remain open; no original workflow increment.
See TYPED_SQL_WRITES_RUNTIME.md. Shared delivery/platform validation follows.

Shared delivery validation: host 48/48 (84.02 s), Debian 12 48/48 (99.21 s),
Ubuntu 24.04 48/48 (86.65 s), Debian 13 48/48 (87.95 s). Host/all three distro
DEBs pass payload/backend/build-ID/permissions/fixture-exclusion checks; isolated
distro installs pass backend execution/GUI startup. GUI-disabled core sources
are unchanged from their last checked state. All four original hashes/private
visibility pass. Remote CI remains blocked by the previously confirmed account
restriction; no remote success is claimed.

## Native Variant array runtime (2026-11-02)

Implemented one-dimensional Array values, Option Base 0/1, Long bounds,
checked indexed reads/writes, type inspection, immutable-backed value copies and
whole-array ByRef/ByVal execution. Actual VBA loops persist modified array values
through DAO into a synthetic SQLite database. Indexed ByRef elements fail 3251.
Final focused existing core/DAO tests passed (2/2, 1.37 s), including nested
reads and explicit rejection of unsupported indexed ByRef writes. All four
original fixture SHA-256 checks passed. Full GUI/package/distribution validation
was not repeated for this segment; the last full result belongs to e799b2f.
No new CTest target and no repeated distribution matrix for this segment.
Full Access array semantics remain open; see VARIANT_ARRAY_RUNTIME.md.

## Variant array resizing (2026-11-02)

Added one-dimensional Variant ReDim, ReDim Preserve and Erase. Actual procedures
exercise negative lower bounds, growth/shrinkage, Empty initialization, unchanged
copies, rejected lower-bound changes and release/reallocation. The existing core
and DAO targets pass (2/2, 1.27 s); no new target or distribution-matrix repetition.
Explicit allocation limit: 1,000,000 elements (3251 beyond it). Typed declarations,
multidimensional arrays, array enumeration and indexed ByRef binding remain open.

Combined array delivery validation: desktop app build and GUI-free core build
passed; GUI-free core test 1/1 (1.10 s). The new host .deb passed package-content,
permissions, reference Build-ID and pinned SQLite 3.54.1 backend checks. All four
original fixture hashes remain verified. Remote branch confirmed at 18b2bec and
repository privacy confirmed true. No full 48-test or distribution matrix rerun;
the last full matrix remains e799b2f.

## Native text-list and reverse-string programming block (2026-11-02)

Implemented native Split, Join and Filter with typed String arrays; bounded array
For Each observes later element writes while preserving independent value copies.
Actual VBA list filtering/enumeration persists through DAO. Added InStrRev and
StrReverse to VBA and compiled Access SQL through a shared native runtime adapter,
with projection-alias use and strict malformed Unicode/error handling. Full BSTR,
array declaration/rebinding/alias, locale and Jet projection typing remain open.
Final existing core/DAO/query-model targets passed 3/3 (1.55 s), including the
UTF-8 Join guard. Desktop and GUI-free builds passed; GUI-free core 1/1 (1.01 s).
The .deb passed content/permissions/Build-ID/pinned-backend checks and all four
original fixture hashes passed. Packaging used an installed staging tree to avoid
recompiling unrelated test binaries. No new CTest target or full distribution
matrix repetition.
See VARIANT_ARRAY_RUNTIME.md and REVERSE_STRING_RUNTIME.md for exact boundaries.

## Optional scalar parameter defaults (2026-11-02)

Implemented omitted trailing Optional parameters with explicit scalar-literal
defaults in native Sub/Function execution, fresh checked binding per call and
preserved supplied ByRef/ByVal behavior. Required arity still fails 450. Native
VBA function defaults also feed actual DAO insertions. Existing core/DAO targets
pass 2/2 (1.53 s); desktop build passed. No new test target or full-suite/matrix
rerun. Missing/IsMissing, positional holes, named arguments, Object/constant
expression defaults and complete Optional/ParamArray semantics remain open.
The previous combined package validation belongs to 7dd65e5; this incremental
parameter segment has not yet repeated package/headless/distribution validation.

## Native VBA Like and combined parameter delivery (2026-11-02)

Implemented VBA Like at comparison precedence: wildcard/digit/class/range/
negation matching, empty [] groups, Null, module compare mode and error 93.
The matcher uses two-row dynamic programming over UTF-16 units. Full Windows
locale/ligature/BSTR matching remains open; unsupported range/surrogate cases
fail 3251. Existing DAO execution now combines Split/Filter, array For Each,
Like conditions, an Optional defaulted function and real SQL persistence.

Allowed same-allocation array element writeback through a ByRef helper inside
array For Each; later element changes remain visible while copies stay unchanged.
Full cross-module rebinding/alias semantics remain open.

Final existing core/DAO targets pass 2/2 (1.42 s), GUI-free core 1/1 (1.08 s),
desktop and GUI-free builds pass. Combined package (including 7235313) passes
contents/permissions/Build-ID and pinned SQLite backend checks. No new test target
and no full 48-test/distribution matrix repetition. Original fixtures were never
modified; all four hashes remain checked. See VBA_LIKE_RUNTIME.md.

## Native ParamArray evaluation (2026-11-02)

Native Sub/Function calls now bind a final ParamArray to a zero-based Variant
array, independently of Option Base. Supports multiple/no extra arguments,
required preceding arguments, bounds/indexing/For Each/Join, typed results and
ByVal helper forwarding. Actual DAO writes combine this with Optional defaults.

The first scope is argument evaluation: slot replacement, ReDim, whole-array
assignment and ByRef forwarding fail explicitly with 3251. Optional/ParamArray
combinations, named/hole binding, expansion and full slot/reference semantics
remain open. Existing core/DAO targets pass 2/2 (1.44 s); desktop build passes.
No new CTest target and no additional package/headless/distribution pass for this
increment. Last combined package/headless validation belongs to e0afbcd.
See PARAMARRAY_RUNTIME.md for supported scope and remaining boundaries.

## Native Missing and named argument programming block (2026-11-02)

Implemented Optional without explicit defaults, a distinct Missing Variant,
IsMissing/Error metadata, positional holes and Missing forwarding that does not
write callee defaults into omitted caller slots. Added user-procedure named
binding with reordered values/references, duplicate/unknown/required checks,
and native DoCmd parameter signatures. Existing GUI execution performs real SQL,
opens a filtered child, reads its record-bound control and closes it by name.

Existing core/DAO/current-record targets pass 3/3 (1.94 s); desktop build passes.
No new test targets or full-suite/distribution repetition. Package/headless
validation is still the earlier e0afbcd baseline pending a shared later delivery.
Full builtin/DAO/COM named calls, SCODE/coercion and Windows semantics stay open.
See NAMED_ARGUMENT_RUNTIME.md and OPTIONAL_ARGUMENT_RUNTIME.md.

## Native form/report OpenArgs programming block (2026-11-02)

Native text/Null OpenArgs is stored before document events, with instance and
Forms/Reports reads. Existing GUI execution verifies Form_Open/Form_Load and
Report_Open SQL audit values, parent Null, close/reopen with new args, and a
current assignment rejection. Nontext values and replacing an open document's
args fail explicitly. Default data/window mode constants are accepted while
other modes remain unsupported. Core/DAO/native GUI targets pass 3/3 (1.78 s);
desktop build passes. No new test target or matrix/package repetition yet.
See DOCUMENT_OPENARGS_RUNTIME.md for remaining Windows/Variant/property limits.

## Native RunSQL UseTransaction (2026-11-02)

RunSQL evaluates positional/named/default UseTransaction. True uses a unique
native savepoint with checked release/rollback cleanup; False adds no wrapper.
Existing DAO workspace transactions remain active and own their outer rollback.
Native GUI execution verifies successful calls, late unique failure with original
3022, continued execution, outer DAO rollback and Null SQL rejection (94) through
both direct and named syntax. Existing core/DAO/query-model/GUI targets pass 4/4
(1.83 s), desktop build passes. No new target or full matrix/package repetition.
SQLite atomicity remains with False; complete Jet partial-failure semantics,
cross-connection execution and callback reentrancy remain open. See
DOCMD_RUNSQL_TRANSACTION_RUNTIME.md for the precise transaction boundary.

## Native hidden-form opening (2026-11-02)

New acHidden forms initialize and retain data/OpenArgs independently without
show/activation. Native GUI SQL audit verifies Open/Load/Current without Activate,
Visible=False, Forms access and named close. Existing core/DAO/query-model/GUI
targets pass 4/4 (1.94 s); desktop build passes. No new target or full-matrix run.
Re-hiding visible documents, report/icon/dialog modes and full Windows event/
window equivalence remain open; see HIDDEN_FORM_RUNTIME.md. Combined package/
headless verification will cover a later grouped delivery.

## Native named builtins and Nz correction (2026-11-02)

Selected existing intrinsic functions now bind named values before their native
implementation, with written evaluation order, defaults and checked duplicate/
unknown/required conditions. InStr retains full positional shape to distinguish
its overloads. Native DAO persistence consumes named Join on ParamArray values.

Corrected Nz(Null)'s VBA subtype to Empty using the Microsoft reference; query
expressions retain zero-length String. Existing arithmetic/concatenation now
matches the documented examples. Core/DAO/query-model/GUI targets pass 4/4
(1.76 s); desktop build passes. No new targets/full matrix repeat. Full builtin
metadata, Object/SCODE and Windows coercion remain open. See NAMED_BUILTIN_RUNTIME.md.

## Native dialog/window programming block (2026-11-02)

Implemented form acDialog with real Qt modality and caller continuation only
on close/hide, retaining cancelled Unload. Added imported/readable Modal/PopUp,
hidden first reveal/guarded re-hide, icon opening and active-document Minimize/
Restore/Maximize. Existing GUI SQL audits verify dialog event/caller ordering,
cancelled then accepted Unload, hidden events and Qt minimized/restored states
with a distinct caller. Core/DAO/query-model/GUI targets pass 4/4 (1.99 s);
desktop build passes. Full application-Quit propagation is the next follow-up;
other Window/desktop/data/view/property limits are in DIALOG_WINDOW_RUNTIME.md.

The earlier grouped adea3f8 snapshot passed GUI-free core build/test 1/1 (1.23 s)
and .deb contents/permissions/Build-ID/pinned-backend checks. This new dialog
increment has not repeated that grouped packaging/headless validation yet.
No new CTest targets or full distribution-matrix repetition.

## Quit propagation and grouped validation (2026-11-02)

Accepted native Quit now stops active procedure/expression/dialog frames, including
Open/Load construction initiated by the UI; cancelled Quit remains error 2501.
Typed array function returns retain their subtype. The converted-form opener in
mainwindow now uses independent native document opening. DoCmd.RunSQL SELECT
reports 2342 while DAO.Execute retains 3065. See QUIT_EXECUTION_RUNTIME.md.

The full existing suite passed 46/48; both failures shared an obsolete RunSQL
expectation. After correcting the DoCmd-specific error mapping and expectation,
the two affected targets passed (50.87 s). Thus all 48 existing targets are
covered by the full run plus focused rerun, not one final 48/48 invocation.
GUI-free build/test passed 1/1 (1.74 s). All four original fixture SHA-256 checks
passed. The .deb passed contents, permissions, matching Build-ID and pinned
SQLite backend checks. No new test targets or distro matrix rerun. Remote CI
remains blocked by the previously reported billing restriction.

## Named DAO method arguments (2026-11-02)

Existing native Database.OpenRecordset/Execute/temporary CreateQueryDef,
QueryDef.OpenRecordset/Execute, TableDef.OpenRecordset and Recordset Find methods
now advertise their implemented argument names to the expression binder. Written
argument evaluation order and duplicate/unknown/required checks use the existing
binder. Unsupported optional Options/LockEdit parameters are not advertised or
silently ignored. Existing native method/type/SQL restrictions still apply.
The existing DAO regression now uses reordered named recordset/action/query
arguments and named FindFirst criteria against real SQLite data. No new target.

Desktop build and the existing dao_execution target pass (0.30 s) for this
increment. Grouped packaging/full-suite evidence remains the preceding Quit
snapshot; neither was repeated for this argument-metadata-only increment.

## Native DAO relative/bookmark movement (2026-11-02)

Implemented Recordset.Move Rows, Start with native Long conversion and bookmark
validation, independent cursor positioning, boundary clamping and pending-edit
guards. Method signature supports named Rows/Start. Existing cursor regression
checks relative movement, bookmark offset, EOF return, stale-token rejection
and Long overflow. Desktop build and dao_execution pass (0.30 s); no new target,
full-suite/package rerun or claim of full DAO/Windows cursor equivalence.

## Native DAO edit metadata (2026-11-02)

Recordset.EditMode returns Integer dbEditNone/dbEditInProgress/dbEditAdd from
the actual per-cursor buffer state; Updatable returns the real table/snapshot
write capability. Both metadata properties reject assignment. Added native
VBA dbEdit constants. Existing clone/edit/add/update/snapshot regression now
checks these values. Desktop build and existing core/DAO targets pass 2/2
(1.64 s). No new target or repeated full-suite/package run. Concurrent server
permissions, locks and all Windows DAO capability rules remain outside scope.

## Native Application.Quit entry point (2026-11-02)

Application.Quit and named default save-all options now reach the native
DoCmd.Quit close/cancellation/frame-stop path; added AcQuitOption constants.
Prompt/save-none design modes still fail explicitly. Existing form_current_record
regression passes (0.50 s), including named Application.Quit inside acDialog and
no child/caller continuation. Desktop build passes. No new test target or
full-suite/package rerun; see QUIT_EXECUTION_RUNTIME.md for scope.

## Native default command buttons (2026-11-02)

SaveAsText Default metadata (including control templates) now configures native
CommandButton defaults. Enter in a form editor validates pending controls and
invokes the visible/enabled default button; a focused command button retains
priority. VBA reads/writes Default use the actual Qt default state and selecting
another default clears the previous one. Null assignment is rejected. Existing
GUI regression uses real key events and SQL writes to check imported defaults,
runtime switching and focused-button priority. Desktop build and existing
form_current_record target pass (0.49 s). No new target. Cancel-button/Escape,
multiline EnterKeyBehavior, subform and complete Windows keyboard parity remain
open; this increment does not claim those semantics.

Grouped delivery check for c63c584: GUI-free build/test 1/1 (1.25 s), matching
.deb Build-ID/permissions/contents/pinned SQLite backend, all four original
fixture SHA-256 checks and private repository status pass. No distro matrix
rerun; remote CI billing restriction unchanged.

## Attached-label keyboard access (2026-11-02)

The SaveAsText parser now retains an attached label's parent control index.
Native labels attached to text/combo/list editors use Qt buddy mnemonic focus,
including ampersand captions. Existing GUI regression checks structural binding
and real Alt+I focus transfer. Desktop build and form_current_record pass
(0.48 s). No new target. Detached-label heuristics, duplicate mnemonic cycling,
label mouse-event rules and complete Windows accessibility/printing equivalence
remain open. Grouped packaging/headless evidence above covers c63c584 only.

## Native form Edit/ReadOnly opening modes (2026-11-03)

OpenForm DataMode now reaches native construction before Open/Load and controls.
ReadOnly disables all three edit/add/delete policies; Edit enables them despite
imported disabled policies. Existing native read-only query restrictions remain.
Existing core/GUI targets pass 2/2 (1.80 s), including Open-event policy reads,
actual read-only editor state and real keyboard edit persisted to a test table.
Desktop build passes. No new target. See FORM_DATA_MODE_RUNTIME.md for remaining
DataEntry/reuse/locking/view/reference boundaries.

## Grouped native DAO/form verification (2026-11-03)

The bdc5466 desktop/all-target build passed. Full existing suite: 46/48 passed
(56.94 s); both failures shared an obsolete expectation that acFormEdit must
fail. The negative case now checks still-unsupported acFormAdd; positive Edit
behavior is covered by the native keyboard/persistence regression above.
The two affected original-layout/DAO targets then passed 2/2 (47.20 s). All 48
existing targets are thus covered across full run and focused rerun, not a
single final 48/48 invocation. No new test targets.

GUI-free build/test passed 1/1 (1.37 s). Matching .deb Build-ID, pinned SQLite
backend smoke, package permissions/contents, all four original fixture SHA-256
checks and repository privacy passed. No distro matrix repeat; remote CI remains
blocked by the previously reported billing restriction.

## Native Cancel/Escape button policy (2026-11-03)

Implemented imported/template Cancel, unique live VBA Default/Cancel switching,
checked Boolean policy conversion and real Escape Click dispatch after focused
TextBox undo. No implicit close replaces a missing handler. Also guarded native
Default reads when a control is rendered as an unsupported placeholder. Existing
GUI regression now uses VBA to switch policies and real Escape/Enter/Alt events
with persisted SQL effects. Desktop build and form_current_record pass (0.56 s).
No new target/full-suite/package repeat for this increment. The grouped evidence
above covers e5143eb; see COMMAND_BUTTON_KEYBOARD_RUNTIME.md for boundaries.

## Guarded control and mnemonic focus (2026-11-03)

Native Control.SetFocus now checks pending validation before moving focus,
returning 2110 on cancellation and rejecting forms closed during validation.
Attached-label shortcut dispatch also checks validation before native buddy
focus and contains GUI-side exceptions with deferred document disposal.
The existing GUI regression activates its actual window, checks cancelled
BeforeUpdate through both VBA/Alt focus paths, accepts corrected text and
persists it. Desktop build and form_current_record pass (0.54 s). No new target
or full-suite/package repeat. Full mouse/cross-document/Form.SetFocus and exact
Windows focus-event ordering remain open.

## Native DAO absolute cursor positions (2026-11-03)

Implemented Long Recordset.AbsolutePosition read/write using the actual
materialized cursor position, -1 without a current row, independent clone
navigation, checked bounds/conversion and pending-edit guards. Existing DAO
regression checks cursor isolation, EOF and invalid/pending moves; native VBA
reads/writes the property while consuming actual SQL results. Desktop build and
dao_execution pass (0.33 s). No new target/full-suite/package repeat. Deleted-row,
lazy cursor, external/forward-only and exact Windows error semantics remain open.
Reference: https://learn.microsoft.com/en-us/office/client-developer/access/desktop-database-reference/recordset-absoluteposition-property-dao

## Native DAO LastModified and key-change identity (2026-11-03)

Added cursor-local read-only LastModified bookmark after successful Update/AddNew,
with native token validation and explicit snapshot/unmodified/stale boundaries.
DAO Update consumes SQLite RETURNING rowid and remaps the shared cached identity
before refresh so active clones and existing bookmarks follow the same record
when its Integer primary key changes. Existing cursor regression changes/restores
a real primary key, follows both cursors/bookmarks, checks per-cursor ownership
and add/update recovery. Existing native VBA assigns Bookmark=LastModified.
Desktop build and dao_execution pass (0.30 s); no new target/full-suite/package
repeat. External SQL/triggers, form-save identity and exact Windows edge/error
semantics remain open.
Reference: https://learn.microsoft.com/en-us/office/client-developer/access/desktop-database-reference/recordset-lastmodified-property-dao

## Native form key-change identity (2026-11-03)

Native form update/save now consumes RETURNING rowid and follows the actual
committed key. Snapshot refresh accepts checked identity mappings so existing
Recordset.Clone cursors and old form byte-array bookmarks follow the same row
through a changed sort position. Old field values remain captured before refresh.
Existing GUI regression uses a real primary-key editor/save and checks current
record, old bookmark, clone identity and persisted SQL. Desktop build and existing
GUI/DAO targets pass 2/2 (0.90 s). No new target/full-suite/package repeat. See
FORM_KEY_IDENTITY_RUNTIME.md for external trigger/SQL/filter/concurrency limits.

## Grouped native identity/keyboard verification (2026-11-03)

Snapshot 992d082 passed the complete desktop/all-target build and all 48 existing
CTest targets in one run (83.53 s). Matching .deb Build-ID, pinned SQLite backend
smoke and package contents/permissions pass. All four original fixture SHA-256
checks and repository privacy pass. The GUI-free runtime sources are unchanged
since the prior successful 1/1 GUI-free check (1.37 s); no redundant repeat.
No new CTest targets, distro matrix repeat or remote CI success claim. Remote CI
remains blocked by the previously reported billing restriction. Complete Access
2003 parity is still not established.

## Native direct-table DataEntry (2026-11-03)

Implemented imported and live VBA DataEntry with SQL-bound own-entry rowids,
actual insertion persistence, key-change tracking, committed-deletion cleanup,
10000-row limit and AllowAdditions dependency. DataEntry=False saves and removes
filter/sort; a fresh entry transition clears the session. Grouped OR filters
prevent existing-row leakage. Native host/property routes support Me/unqualified
DataEntry; active non-table sources fail explicitly. acFormAdd is still open.

Existing native GUI regression enters two real rows, changes a key, compares
session versus physical counts, switches through VBA and toggles AllowAdditions.
Desktop build and form_current_record pass (0.60 s). No new test target or
full-suite/package repeat; previous grouped verification remains 992d082.
See DATA_ENTRY_RUNTIME.md for external SQL/rollback/query/reference boundaries.

## Native Add-mode opening (2026-11-03)

OpenForm acFormAdd now creates the real direct-table DataEntry session and
forces AllowAdditions before startup events, retaining imported edit/delete
policies. The existing native GUI regression opens over pre-existing records,
sees an empty entry recordset, enters/saves a real row and checks session versus
physical table counts. The negative mode case now checks unsupported value 99.
Core/GUI targets pass 2/2 (1.94 s), desktop build passes. After rebuilding the
original-layout adapter, existing GUI/original-layout/real DAO targets pass 3/3
(49.84 s). GUI-free build/test passes 1/1 (1.25 s). No new targets or full suite/
package/distro repeat for this increment. Full Windows edit/delete override
matrix, writable query entry and open-instance mode changes remain open; see
FORM_DATA_MODE_RUNTIME.md. Last full 48/48/package evidence remains 992d082.

## Native NewRecord subtype and availability (2026-11-03)

Corrected native Form.NewRecord from Boolean to its documented read-only Integer
-1/0 subtype. Qualified/unqualified/legacy reads now share the same native getter;
True requires an available direct-table new-row slot with additions enabled.
The existing GUI regression checks TypeName/value through actual VBA and an
empty source with additions disabled after real SQL deletion/requery. Desktop
build and form_current_record pass (0.64 s). No new target/full-suite/package
repeat; full Windows unbound/view/no-current edge cases remain open.
Reference: https://learn.microsoft.com/en-us/office/vba/api/access.form.newrecord

## Entry/NewRecord delivery verification (2026-11-03)

Snapshot 1490e24 passed rebuilt original-layout/real DAO targets 2/2 (50.11 s),
matching .deb Build-ID/pinned backend/contents/permissions and all four original
fixture SHA-256 checks. GUI-free runtime remains the same source verified 1/1
(1.25 s) in the Add-mode delivery. No full-suite/distro repeat. Last one-run full
48/48 evidence is 992d082; no remote CI success claim.

## Native DAO field ordinal metadata (2026-11-03)

Recordset Field.OrdinalPosition now returns checked Integer metadata from the
actual native field/projection order, with read-only and closed-owner guards.
Existing native VBA reads a calculated SQL field's position; the existing cursor
regression checks metadata writes and invalidation after Close. Desktop build
and dao_execution pass (0.31 s). No new target/full-suite/package repeat for
this increment. TableDef field reordering/persistent ordinal gaps and full DAO
metadata remain open.
Reference: https://learn.microsoft.com/en-us/office/client-developer/access/desktop-database-reference/field-ordinalposition-property-dao

## Preserve existing native String/Byte arrays (2026-11-03)

ReDim Preserve now grows/shrinks existing typed String/Byte arrays held by Variant
variables, retaining subtype/prefix and filling new cells with empty String or
zero Byte. Lower-bound checks, copy/ByVal independence, supported ByRef writeback,
iterator protection and the one-million-element limit remain explicit. Non-
Preserve typed resizing, typed array declarations, Byte Erase and multidimensional
arrays remain open. Existing core procedures verify native typed return data,
copy isolation, growth/shrink/defaults and a rejected Byte lower-bound change.
Desktop build/core test pass 1/1 (1.60 s), existing DAO target passes (0.45 s),
and GUI-free build/test passes 1/1 (1.33 s). No new target/full-suite/package repeat.

## Native qualified VBA.Array (2026-11-03)

Qualified VBA.Array now uses the native intrinsic with base zero regardless of
Option Base; ordinary Array retains the module setting. Direct For Each supports
the qualified result. Existing local/host VBA roots retain normal lookup; named
qualified Array and full reference/library binding remain explicit boundaries.
Core regression checks base 1 versus 0, empty bounds, actual direct enumeration
and local-root precedence. Desktop build/core target pass 1/1 (1.32 s); GUI-free
build/test passes 1/1 (1.39 s). No new target/full-suite/package repeat.

## Native optional Let assignments (2026-11-03)

The optional Let keyword now routes existing local/function-result, form/control,
array-element and object-property value assignments through their native
instructions. Non-assignment Let forms are rejected before execution rather
than becoming message/command calls. Existing core procedures verify actual
assignment/return values and invalid Let rejection; DAO VBA uses Let on a real
cursor property. Desktop build and core/DAO targets pass 2/2 (1.61 s), original
VBA inventory passes (0.91 s), and GUI-free build/test passes 1/1 (1.30 s). No new
target/full-suite/package repeat. Property Let declarations, unsupported lvalue
forms and the complete lexer/AST remain open.

## Checked native numeric And/Or (2026-11-03)

Non-Null numeric And/Or now coerce through checked 32-bit VBA Long operands
instead of accepting host-wide 64-bit bit masks. Two declared Integer operands
retain Integer results; other numeric cases currently return Long, matching the
existing bounded Xor/Eqv/Imp path. Boolean-only behavior remains native Boolean.
Core regression checks subtype, banker rounding and 32-bit overflow. Desktop
build and core/DAO targets pass 2/2 (1.66 s). No new target/full-suite/package
repeat. Complete Byte/mixed/Null promotion matrices and all Windows operator
coercion behavior remain open.

## Grouped native language verification (2026-11-03)

Snapshot 32f548e passed the complete desktop/all-target build. Full suite:
45/48 passed (89.16 s); the three failures directly read the obsolete Boolean
storage of NewRecord after its documented Integer correction. Existing state
assertions now use truthy() for both positive and negative checks; the dedicated
native VBA test still verifies Integer subtype. The three affected GUI targets
then passed 3/3 (1.48 s). All 48 existing targets are covered across full run and
focused rerun, not one final 48/48 invocation. No new targets.

GUI-free build/test passed 1/1 (1.44 s). Matching .deb Build-ID, pinned SQLite
backend smoke, package contents/permissions, all four original fixture SHA-256
checks and repository privacy passed. No distro matrix repeat; remote CI billing
restriction unchanged. Last single-run 48/48 evidence remains 992d082.

## Native based integer literals (2026-11-03)

Expression lexing now recognizes hexadecimal &H and octal &O integers with
checked 32-bit patterns, Integer/Long identity and explicit %/& suffixes. Signed
16/32-bit patterns reach actual native bitwise operations; based literals also
work in supported Optional literal defaults. Existing core checks cover masks,
Not, subtype, concatenation, octal values, range failures and a default Long
parameter. Desktop build and core/DAO/original syntax targets pass 3/3 (2.46 s);
GUI-free build/test passes 1/1 (1.87 s). No new target/full-suite/package repeat.
Complete decimal literal inference, all unary/suffix/leading-zero edge rules,
Const declarations and raw Jet SQL literal translation remain open.

## Native character SQL bridge and reverse-search Null (2026-11-03)

Asc/AscW/Chr/ChrW now execute through registered native SQL callbacks with bound
scalar arguments, no UI context and shared VBA conversion logic. Added native
Chr$/ChrW$ and named String/CharCode argument signatures. Existing ANSI ASCII and
isolated-surrogate boundaries remain explicit; full codepages/BSTR support and
all SQL result metadata remain open.

Reference correction: InStrRev returns Null for Null string arguments, while
StrReverse(Null) raises 94. Native core now follows that distinction and the SQL
scalar emitter preserves actual Null results. Existing SQL regression verifies
character conversion on real rows, Unicode roundtrip, typed errors and Count
excluding Null reverse-search results; core checks named/dollar conversions.
Desktop build and core/DAO/query-model targets pass 3/3 (1.62 s). No new target or
full-suite/package repeat. GUI-free follow-up is recorded with the next delivery.
References: Microsoft Learn InStrRev/StrReverse function return rules.


## Native repetition functions (2026-11-03)

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


## Native table Close and hidden editor lifetime (2026-11-03)

Explicit and default interpreted DoCmd.Close now close actual table pages after
checked field validation and transactional record submission. Failure retains
page/editor/drafts. Hidden pending editors and a reentrant live table popup are
protected against data loss and dangling lifetime. See TABLE_DOCUMENT_RUNTIME.md.
Existing targeted trio passes 3/3 (15.85 s). Full desktop suite passes 48/48
(91.15 s), GUI-free core 1/1 (2.01 s); Debian payload/permissions/Build-ID/
pinned SQLite checks and four original fixture hashes pass.
No new test targets or full Access parity claim.
Theoretical checklist: 292 checked / 70 open scoped points.


## Native table SelectRecord and Copy (2026-11-03)

The actual table dispatcher now handles row selection and checked clipboard Copy
from interpreted DoCmd/Application calls. Readonly rows copy, dirty drafts copy
without writing SQL, and invalid active fields refuse either command without
replacing clipboard text. Canonical details: TABLE_DOCUMENT_RUNTIME.md.
Theoretical checklist: 294 checked / 70 open scoped points.

Existing main-window regression passes 1/1 (1.97 s), including invalid numeric
field refusal with unchanged clipboard. Last full desktop/package baseline is
local 2dd59b2 / private API 32cb556, 48/48 (91.15 s), GUI-free core 1/1 (2.01 s).
This controller increment is focused-verified; no new targets.


## Table command Zoom and active-column retention (2026-11-03)

acCmdZoomBox now opens the actual native table cell, with checked typed cache
updates and normal Cancel/readonly Close completion. SelectRecord retains the
active column: an actual Qt first-column jump was reproduced through Zoom and
fixed without losing whole-row selection. See TABLE_DOCUMENT_RUNTIME.md.
Existing main-window/Currency targets 2/2 (2.47 s); full/package baseline remains
2dd59b2 / API 32cb556, 48/48 (91.15 s). No new targets.
Theoretical checklist: 295 checked / 70 open scoped points.


## Native table record navigation and checked autosave (2026-11-03)

Named/default GoToRecord and native First/Previous/Next/Last navigate actual
SQL-backed view order, including rows past Qt's initial batch. A sole changed
existing record with a primary key is saved transactionally before relative
navigation. Multiple drafts, inserts/deletes and unsafe current identity remain
explicit unsupported boundaries. SQL failure retains the current row/draft for
correction and retry. Canonical details: TABLE_DOCUMENT_RUNTIME.md.
Theoretical checklist: 297 checked / 71 open scoped points.

Full rebuilt desktop suite passes 48/48 (91.33 s), GUI-free core 1/1 (2.49 s),
Debian payload/permissions/Build-ID/pinned SQLite checks and all four original
fixture hashes pass. Existing focused trio passes 3/3 (16.79 s), exercising
300 actual SQL rows, named/default/relative targets, sole-record SQL persistence,
constraint failure/retry and untouched multirow drafts. Fixtures explicitly sort
ascending; production navigation follows the current view order. No new targets.


## Native table sort commands (2026-11-03)

Interpreted ascending/descending sort acts on the actual current field through
Qt/SQL, retaining the primary-key record and field after fetching as needed.
Readonly sorting works; a sole current edit is checked/saved, unrelated drafts
are explicitly refused. See TABLE_DOCUMENT_RUNTIME.md.

Existing main-window regression passes 1/1 (2.08 s): real ascending/descending
SQL row order, readonly sorting, primary-key/current-field retention past the
initial batch, checked sole-record persistence and refusal of unrelated drafts.
Last complete suite/package baseline remains e4e0900 / private API 714a96e,
48/48 (91.33 s), GUI-free core 1/1 (2.49 s). This controller change is
focused-verified, without adding test targets.
Theoretical checklist: 298 checked / 71 open scoped points.


## Native table Find UI, commands and lifetime (2026-11-03)

Actual modeless Find supports literal text scope/modes/case, Ctrl+F/Shift+F4,
native Edit menu and DoCmd/Application Find/FindNext. It fetches actual SQL rows,
reactivates its owning page, saves a sole current update through checked native
transactions, protects unrelated drafts and is disposed with its table. View
refusals return their actual native error to VBA; binary Copy remains 3251 with
unchanged clipboard/payload. Canonical scope: TABLE_FIND_RUNTIME.md.
Theoretical checklist: 301 checked / 72 open scoped points.

Full rebuilt desktop suite passes 48/48 (92.69 s), GUI-free core 1/1 (2.45 s),
Debian payload/permissions/Build-ID/pinned SQLite checks and all four original
fixture hashes pass. No additional test targets. This also covers the preceding
native table sort commands and the final weak-owner reentrancy guard.


## Bounded native DoCmd.FindRecord (2026-11-03)

Current-text-field SearchAll executes on actual active tables/forms with named
and omitted argument binding, distinct explicit Empty coercion, first/next origin
and normal no-match completion. Readonly fields work; form validation cancellation
raises 2501 before subsequent VBA statements. Explicit document context prevents
old-caller fallback from the database window. General formatted/numeric/all-field/
directional/wildcard behavior remains open. Canonical scope: FIND_RECORD_RUNTIME.md.
Existing trio 3/3 (15.83 s), GUI-free core 1/1 (1.99 s). No new targets.
Theoretical checklist: 303 checked / 72 open scoped points.

Full rebuilt desktop suite passes 48/48 (88.05 s). Debian payload/permissions/
Build-ID/pinned SQLite checks and all four original fixture hashes pass. The
GUI-free core baseline after binding/constant changes is 1/1 (1.99 s); subsequent
form extensions are Qt-only. No additional test targets.

## Native search direction extension (2026-11-03)

Table/form dialogs and DoCmd.FindRecord now support Up/Down without wrap and
SearchAll with wrap. FindNext retains the chosen direction. FindFirst starts at
the first record; otherwise the API starts after the current record in that
direction. Existing main-window/form-Find/layout trio passes 3/3 (16.70 s),
covering real opposite-boundary refusal, reversed matching, UI and interpreted
named direction arguments. Last complete suite/package baseline is a71acc9 /
private API f9367b0, 48/48 (88.05 s), GUI-free core 1/1 (1.99 s). This Qt-only
extension is focused-verified; no extra targets.
Theoretical checklist: 304 checked / 72 open scoped points.


## Typed native search in dialogs and FindRecord (2026-11-03)

Exact Number/Currency/Date/Boolean matching is shared by real table/form dialogs
and interpreted FindRecord. Adjacent Doubles, four-decimal units, native Date
canonicalization and NULL-versus-False remain distinct; readonly search and
invalid query retention are checked against actual SQL values. Canonical details:
TYPED_FIND_RUNTIME.md. Existing trio 3/3 (2.79 s); no new targets.
Theoretical checklist: 308 checked / 72 open scoped points.

Full rebuilt desktop suite passes 48/48 (97.33 s), GUI-free core 1/1 (1.86 s),
Debian payload/permissions/Build-ID/pinned SQLite checks and all four original
fixture hashes pass. No extra test targets. This also covers the preceding
search-direction extension and the typed native invocation/state changes.

## Scoped table keys and DoCmd.SendKeys (2026-11-03)

Native table cells and their owned Find/Zoom dialogs now resolve actual focus
per key. Global and named/omitted DoCmd.SendKeys share ordered queued delivery,
modal continuation and DoEvents draining. Project identity stays fixed, queued
focus changes cannot type into an old caller, and real Text/Value/SQL buffering
is retained. Canonical scope and resource limits: TABLE_KEYBOARD_RUNTIME.md.
Existing main-window/layout pair 2/2 (17.52 s), original Kassa event target
1/1 (38.68 s); no new targets. The original async probe now yields the full
queue instead of assuming one processEvents call; receiver/content checks stay.
Theoretical checklist: 310 checked / 72 open scoped points.

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

Final main-window storage check passes 1/1 (3.81 s), including saved-main-schema
lookup, temporary-shadow isolation, explicit no-catalog refusal and genuine
empty-project counts. GUI-free core passes 1/1 (1.71 s).

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


## Editable keyed single-table query documents (2026-11-03)

Default/acEdit OpenQuery uses the real typed SQL table model for the strict
unparameterized SELECT-star/single-main-table/primary-key subset. Query identity,
caller-independent lifetime, actual record Save, validated Close even acSaveNo,
and unchanged-definition Requery are integrated. Invalid records keep the open
cache; changed definitions and source metadata/shadows refuse without discarding
drafts. A second editable source datasheet refuses before focus changes.
OPEN_QUERY_RUNTIME.md documents projection/filter/join/provider/concurrency limits.
Full desktop 48/48 (90.33 s), rebuilt GUI-free core 1/1 (1.48 s), Debian package
and four original fixture hashes pass. No additional CTest targets. Scoped
checklist: 348 checked / 72 open; this is not total Access 2003 parity.


## Filtered editable queries and structural key identity (2026-11-03)

Bounded compiled WHERE predicates feed the native SQL datasheet model. A real
VBA save commits the keyed row and refreshes filtered membership; an excluded
control row remains untouched. Subquery/function/extra-clause and broad
projection/join/parameter boundaries remain explicit. Source primary keys must
be structurally non-null, including actual INTEGER rowid alias proof; nullable
and DESC indexed keys refuse 3251. OPEN_QUERY_RUNTIME.md defines these bounds.
Full suite 48/48 (89.83 s), freshly rebuilt GUI-free 1/1 (1.49 s), Debian payload/
permissions/Build-ID/pinned SQLite and four original fixture hashes pass. No new
targets. Scoped checklist 350 checked / 72 open, not total Access 2003 parity.


## Declared dynamic arrays and real ByRef array parameters (2026-11-03)

Local/module dynamic primitive/Object/Variant declarations retain type through
unallocated state, checked bounds, Preserve and Erase. Module values persist
across procedure calls. Required declared array parameters perform real matching-
type caller writeback; a() denotes the whole array in VBA expression/call/copy
paths. Current Microsoft VBA argument guidance informed the direct declared-
array ByVal guard; explicit Variant copies remain detached. Fixed/class/Optional/
COM and full enumeration/error precedence remain outside the bounded path.
VARIANT_ARRAY_RUNTIME.md documents scope and primary references.
Complete desktop 48/48 (93.65 s), fresh GUI-free core 1/1 (1.85 s), Debian payload/
permissions/Build-ID/pinned SQLite and all four original fixture hashes pass.
Existing core tests extended; no new CTest targets. Scoped checklist 354 checked /
72 open, not total Access 2003 parity.


## Real native VBA file size and timestamp functions (2026-11-04)

FileLen returns checked native byte sizes as VBA Long; FileDateTime converts
actual last-write timestamps to local typed Date values. Relative/absolute
UTF-8 paths, named/qualified calls, NUL/path failures and sparse-file overflow
are integrated without reading/modifying source file contents. Windows drive/
UNC mapping and VBA handle/snapshot/LOF semantics remain open; no general
Windows filesystem or SQL registration claim. FILE_METADATA_RUNTIME.md defines
scope and current primary references. Existing core target extended only.
Full desktop 48/48 (93.70 s), fresh GUI-free core 1/1 (1.65 s), Debian payload/
permissions/Build-ID/pinned SQLite and four original fixture hashes pass.
Scoped checklist 356 checked / 72 open, not total Access 2003 parity.


## Native filesystem metadata through SQL and DAO (2026-11-04)

Native FileLen/FileDateTime callbacks execute from actual parameterized SQL,
returning canonical INTEGER4/DATETIME metadata. Real DAO OpenRecordset fields
expose dbLong/dbDate and actual Long/Date values; proven DATE/DATETIME Field.Type
now returns dbDate. Re-execution reads changed file sizes, compilation performs
no I/O, and checked filesystem errors retain their codes through the sanitized
SQL adapter. No deterministic/pure rewrite declaration or broad filesystem/Jet
claim. FILE_METADATA_RUNTIME.md records scope and remaining boundaries.
Full desktop 48/48 (94.77 s), GUI-free recheck 1/1 (1.89 s; runtime unchanged),
rebuilt Debian payload/permissions/Build-ID/pinned SQLite and four original
fixture hashes pass. Existing query-model/DAO targets extended only. Scoped
checklist 358 checked / 72 open, not total Access 2003 parity.


## Real fixed array storage and retained ByRef binding (2026-11-04)

Fixed local/module arrays accept multidimensional decimal literal bounds and
Option Base. Typed Erase resets native values while preserving dimensions/type;
independent Variant snapshots retain data. Matching ByRef helpers preserve fixed
binding for actual element/reset writeback and guard ReDim/shape replacement.
Module arrays persist across procedures. Constant-expression, class/Static and
full alias/iteration/error behavior remain scoped in VARIANT_ARRAY_RUNTIME.md.
Full desktop 48/48 (91.05 s), freshly rebuilt GUI-free 1/1 (1.76 s), Debian payload/
permissions/Build-ID/pinned SQLite and four original fixture hashes pass. Existing
core target extended only. Scoped checklist 360 checked / 72 open, not total
Access 2003 parity.


## Pure scalar Const bindings and folded fixed-array bounds (2026-11-04)

Local/private/public module constants retain actual immutable checked values.
Pure expression folding uses literals, previous constants and native operators
without app/database/filesystem context. Fixed dimensions accept the resulting
integral expressions; ByRef scalar constant arguments use caller-independent
temporaries. Direct writes refuse before body execution. Additional native VBA
character constants and Access-2003-target CRLF vbNewLine values are implemented.
VBA_CONSTANTS_RUNTIME.md documents inference/linking/class/conditional/default
and compiler diagnostic limits and current primary references.
Full desktop 48/48 (95.27 s), freshly rebuilt GUI-free 1/1 (2.11 s), Debian payload/
permissions/Build-ID/pinned SQLite and four original fixture hashes pass. Existing
core target extended only. Scoped checklist 363 checked / 72 open, not total
Access 2003 parity.


## Immutable optional constant defaults (2026-11-04)

Optional parameter defaults fold pure expressions and prior same-module constants
without runtime host evaluation. Native values are copied/coerced per omitted
call; named overrides, explicit Empty, Null/Missing and Collection Nothing keep
prior meanings. Repeated ByRef default calls are independent. Runtime-dependent
expressions and bare clocks are refused. OPTIONAL_ARGUMENT_RUNTIME.md and
VBA_CONSTANTS_RUNTIME.md define remaining linking/object/compiler boundaries.
Full desktop 48/48 (92.51 s), rebuilt GUI-free 1/1 (1.84 s), Debian payload/
permissions/Build-ID/pinned SQLite and four original hashes pass. A final test-only
UTC constructor adjustment uses Qt-6.4-compatible API; query-model 1/1 (0.04 s)
passes on the existing Qt 6.8.2 environment, not a separate minimum-Qt matrix.
No new CTest targets. Scoped checklist 364 checked / 72 open, not overall parity.


## Bounded context-free native procedure plan reuse (2026-11-04)

Native translation plans cache successful exact-source/name/parser-context
results, including readonly scalar/default/fixed-bound metadata. Execution keeps
fresh host/module/argument binding; mutable objects/DAO/widgets/results are not
cached. LRU limits are 128 entries / 16 MiB accounted data per thread, with large
plans left uncached and optional admission failure preserving translated code.
Existing core checks source changes, case-normalized calls and independent hosts.
An external synthetic private-counter probe produced 200/200 calls correctly
before/after, measuring 784.665/37.9019 ms on this environment; no app-wide speed
claim. VBA_PROCEDURE_CACHE.md specifies scope and memory/performance boundaries.
Full desktop 48/48 (81.60 s), freshly rebuilt GUI-free core 1/1 (1.18 s), Debian
payload/permissions/Build-ID/pinned SQLite and four original hashes pass. No new
CTest targets. Scoped checklist 365 checked / 72 open, not overall Access parity.


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


## Error-resume resource ownership (2026-11-04)

Explicit Resume variants now release obsolete error snapshots; restored live
Loop/With frames retain resources until their real scope ends. Single Next also
uses source-boundary cleanup when skipped after a failure. Existing core and DAO
cases verify collection mutation guards/release and a real SQL clone transaction
guard/release after retried With access. The clone case explicitly positions its
cursor with MoveFirst. No guard was relaxed and no CTest target was added.
Targeted existing cases 2/2 (1.67 s), full desktop 48/48 (81.49 s), freshly rebuilt
GUI-free core 1/1 (1.62 s), Debian payload/permissions/Build-ID/pinned SQLite and
all four original fixture hashes pass.

Progress reporting: the growing count of finely split checklist entries is no
longer used as overall progress. Historical counts describe only their then
scoped inventory. The checklist is not a fixed Access 2003 feature denominator;
complete parity remains open and no total percentage is asserted.


## Fixed 456-point work catalog and structured native errors (2026-11-04)

FIXED_WORK_CATALOG.md freezes all existing requirements as F001–F456 in five
categories and 14 subject subcategories. New details map to existing IDs and
never increase the denominator. 384 scoped checks / 72 open requirements are
not an overall Access parity percentage; broad full-compatibility targets stay
open. README, checklist and binding handoff record this reporting rule.

Under existing F051, native Err.Raise accepts signed Long codes and five
positional/named arguments, including omitted optionals. Writable Number,
Source, Description, HelpFile and HelpContext use checked native coercion.
Structured exceptions/results preserve the actual supplied metadata through
nested calls and native event functions; Clear and explicit Resume reset it.
vbObjectError is a native Long constant. Existing core execution covers actual
custom-error propagation, handler/resume reset, event failure, assigned metadata
and omitted-source Raise behavior. No new CTest target or checklist ID.
Localized standard descriptions, default project source, full HRESULT/handler/
debugger equivalence and Windows Help remain open.

Targeted core 1/1 (1.25 s), full desktop 48/48 (77.97 s), freshly rebuilt
GUI-free core 1/1 (1.39 s), Debian payload/permissions/Build-ID/pinned SQLite and
all four original fixture hashes pass.


## Native Enum constants, Error statement and unary Null (2026-11-04)

Existing F042 now includes real module Enum Long constants: automatic/explicit
integral numbering, aliases, negatives, overflow checks, Public/Private and
qualified/unqualified names. Ordered declarations feed both pure compile-time
defaults/bounds and actual native module state. A local Collection root retains
lookup priority; qualified constant writes are refused before body effects.
Enum-typed variables/arrays/parameters and complete module/class/type-library
rules remain open. Under F051, Error expression uses a native source-mapped
instruction independent of a shadowing Err local; native ErrObject identity is
exposed. Under F041, unary signs resolve scalar defaults and propagate Null.
No new checklist IDs or CTest targets; the catalog remains fixed at 456.

Existing targeted core/DAO 2/2 (1.64 s), freshly rebuilt GUI-free core 1/1
(1.31 s), full desktop 48/48 (82.80 s), Debian payload/permissions/Build-ID/
pinned SQLite and all four original fixture hashes pass.


## Native GoSub/Return and handler retry (2026-11-04)

Under existing F040/F051, native procedure-local GoSub/Return uses a bounded
LIFO stack and preserves actual caller conditional/loop/With scopes. Nested
subroutines share procedure locals, GoTo retains the caller floor, Return
continues after the latest GoSub and unmatched Return raises 3. Label handlers
snapshot the return context, can run their own GoSubs and restore or abandon
the fault context with explicit Resume variants, including Resume 0. Obsolete
snapshots release ownership. Direct handler Return into the suspended fault
GoSub remains refused (3251); numeric and block-label boundaries are unchanged.

Existing core runs nested collection/numeric/inline-If scopes, held mutation
guards and post-loop release, plus a failing callee retried after the handler's
own lease-inspection GoSub. Targeted core/DAO 2/2 (1.80 s), freshly rebuilt
GUI-free core 1/1 (1.29 s), full desktop 48/48 (80.12 s), Debian payload/
permissions/Build-ID/pinned SQLite and four original fixture hashes pass.
No new CTest targets or catalog IDs; the fixed 456-point denominator remains.


## Scoped GoTo and numeric line labels (2026-11-04)

Under existing F040, native GoTo preserves same/ancestor destination scopes and
unwinds only exited conditional, enumeration/numeric-loop and With frames.
Inactive/sibling entry is refused before body effects. Inline labels preserve
their colon during wrapper expansion and are rebound/deduplicated in the merged
procedure. The transient scope-parent table costs one integer per instruction;
cached plans retain source-only label strings and existing integer unwind data.
Subroutine caller floors stay protected; GoSub/error/Resume label destinations
remain top-level. Numeric labels 0–65529 support colon/prefix forms, canonical
leading zeroes, GoTo/GoSub/error/Resume targets and preserved zero special cases.
Physical diagnostics remain separate; Erl/debugger equivalence is open.

Existing core verifies Continue-style loops, exited With/If scopes, inline labels,
released enumeration and refused uninitialized-loop entry. The existing real SQL
clone case verifies a GoTo out of With releases its actual transaction guard
after handled/retried field access. Targeted core/DAO 2/2 (1.98 s), additional
DAO 1/1 (0.37 s), freshly rebuilt GUI-free core 1/1 (1.37 s), full desktop
48/48 (76.53 s), Debian payload/permissions/Build-ID/pinned SQLite and all four
original fixture hashes pass. No new CTest targets or catalog IDs: fixed 456.


## Live native ByRef array-element binding (2026-11-04)

Under existing F043, named array roots bind direct elements through captured
checked offsets and opaque allocation tokens, without reevaluating subscripts
or retaining array contents in metadata. Native Sub/function/named calls, typed
multidimensional/Byte cells, same-cell aliases and scalar forwarding write
immediately to actual sources. Later-argument mutation is observed by ByRef;
ByVal keeps its earlier value. Source replacement is guarded while a cell is
active. Compound/ParamArray addresses and roots reached through existing cell
aliases remain refused (3251); broader property/COM/class identity is open.

Only real aliases publish caller changes; ByVal values no longer copy back old
state. Internal transport copies and unused fresh slots release objects after
binding; public C++ argument vectors retain their contract. Other active alias
frames refresh after publication, releasing stale module/static/object snapshots
while the writer completes its own publications before refreshing itself.
Existing core verifies one-time indices, later-argument order, aliases/forwarding,
immediate module reads, named order, copy isolation, Byte and allocation guards.
Existing DAO clears a clone through a forwarded module-array Object slot plus
a ByVal copy, then runs a real transaction before returning.

Targeted core/DAO 2/2 (2.21 s), freshly rebuilt GUI-free core 1/1 (2.05 s), full
desktop 48/48 (76.34 s), Debian payload/permissions/Build-ID/pinned SQLite and
all four original fixture hashes pass. No new CTest targets or catalog IDs;
fixed 456 remains, with full F043/Access parity still open.


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


## Native TextBox mouse events (2026-11-04)

Under existing F058/F148, imported and owned native form TextBoxes now execute
bound Click handlers after native editor mouse release, preserving handler-driven
cursor/selection changes without committing edits merely because of a click.
A normal left click dispatches once; right clicks do not dispatch Click. The
existing cancelable DblClick binding remains before native word selection, and
its second release does not generate an additional TextBox Click.

SaveAsText OnMouseDown/OnMouseUp metadata is retained. Actual TextBox mouse
handlers receive Integer Button (left=1/right=2/middle=4), Integer Shift
(Shift=1/Ctrl=2/Alt=4) and Single X/Y in the renderer's 96-dpi Twip coordinate
space. Unmapped extra buttons return unsupported 3251. Bound event functions
continue through the existing expression runtime. Failed procedures stop the
remaining click dispatch and retain native event diagnostics. Closed documents
are guarded through native execution/disposal scope.

The existing GUI case verifies MouseDown/MouseUp/Click order, actual typed
coordinates/modifiers, a VBA-driven cursor reset, unchanged record dirty state,
right-click exclusion and cancelable double click. Targeted GUI/layout 2/2
(11.43 s) passed before the final extra-button guard. Full validation follows.
No new CTest targets or fixed-catalog IDs. MouseMove, other control/form mouse
events, complete focus sequencing, custom scale modes and Windows Access event/
rendering oracles remain open; this is not complete mouse compatibility.

Public references: [TextBox.DblClick](https://learn.microsoft.com/en-us/office/vba/api/access.textbox.dblclick),
[TextBox.MouseDown](https://learn.microsoft.com/en-us/office/vba/api/access.textbox.mousedown)
and [event ordering](https://support.microsoft.com/en-us/office/order-of-events-for-database-objects-e76fbbfe-6180-4a52-8787-ce86553682f9).

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


## Report section offsets (2026-11-04)

The layout importer now retains PageHeader/PageFooter heights, exposes section
height/offset metadata, and positions report header, page header, detail, report
footer and page footer without dropping the page bands. Form positioning retains
its existing header/detail/footer order. Invalid heights and coordinate overflow
are rejected. The existing report-event test checks native Qt positions and an
actual SQL-bound detail value, including report-footer-before-page-footer order.

Validation: existing access_2003_layouts 1/1 (9.74 s), full desktop 48/48
(80.53 s), GUI-free core 1/1 (1.71 s), Debian payload/permissions/Build-ID/pinned
SQLite and all four original fixture hashes pass. This remains a single-detail
layout/PDF path: repeated detail bands, pagination, CanGrow, grouping and section
Format/Print/Page scheduling remain open. Fixed catalog 456 (384 checked bounded
entries), 48 CTest targets unchanged; no full Access parity percentage is claimed.


## Fixed-band multi-record report PDF (2026-11-04)

A separate native Bericht als PDF action now emits the actual loaded SQL rows
through repeated fixed detail bands, report header/footer and repeated page
header/footer. It keeps details whole, uses physical TWIP-derived dimensions on
the QPrinter page and restores the displayed record. Per-section QPicture vector
recording avoids rasterized widget output, duplicate bands and double DPI scaling.
The existing current-layout one-page actions keep their previous behavior.

This bounded path refuses truncated rows, unsupported controls/calculated field
bindings, declared group/growing/shrinking/multicolumn/running-sum/page-event rules
and geometry that would silently clip data. Failed output cannot overwrite an
existing destination. Access page setup, complete formatting/event scheduling,
groups, aggregates, CanGrow, subreports and direct-print view remain open (F085).

The existing report-event case checks two SQL rows on two PDF pages, restoration,
and destination preservation. Available pdftotext independently confirms each
value occurs exactly once in order, normalizing Unicode ligatures. Initial raster,
DPI and clipping defects were found and corrected before this validation. No new
CTest target or catalog item was introduced.

Validation: access_2003_layouts 1/1 (9.73 s); full desktop 48/48 (81.28 s);
GUI-free core 1/1 (2.23 s); Debian payload/permissions/Build-ID/pinned SQLite;
all four original fixture SHA-256 checks pass. Catalog stays 456/384/72 and the
broader full Access-2003 parity requirement stays open.


## Native paginated report preview and print dialog (2026-11-04)

Separate Berichtsvorschau and Bericht drucken actions now invoke the fixed-band
multi-record renderer through actual QPrintPreviewDialog/QPrinter/QPrintDialog.
Unsupported preview rendering closes the dialog and reports the failure rather
than leaving an apparently complete blank preview. The direct print dialog
explicitly disables selection/page-range options; it emits the complete report.
Existing current-layout actions and DoCmd direct-print-view limitations remain.

The existing report-event case opens the modal native preview and independently
checks QPrintPreviewWidget.pageCount()==2, then closes it and checks restored SQL
field contents. access_2003_layouts 1/1 (9.29 s) passes; desktop build, Debian
payload/permissions/Build-ID/pinned SQLite and all four fixture hashes pass after
the final dialog-option adjustment. No new target, no repeated full suite for
this toolbar-only change: preceding renderer full48/core1 validation is recorded
above. Physical printer output remains unvalidated. Fixed 456/384/72 unchanged.


## Calculated fields in paginated reports (2026-11-04)

The fixed-band renderer now accepts calculated controls supported by the existing
native Access expression engine, including direct-field Sum/Count. It preflights
all loaded rows before opening the output device and refuses calculation errors.
The existing PDF/real-preview case checks per-row arithmetic (10/20), report-wide
Sum(id)=3, restored controls and refusal of a missing-field expression without
replacing the existing PDF. No new expression parser or aggregate type claim.

Analysis used temporary copies of the original Bridge ZIPs: ten report layouts
contain 124 calculated controls; 87 controls mention Sum. Across those controls,
99 Sum calls have simple field arguments and six have compound expressions. No
source expressions, original data or VBA were printed or published in artifacts.
Compound aggregate arguments and complete typing/group scope remain open (F085).

Validation: existing access_2003_layouts 1/1 (9.46 s), desktop build, Debian
payload/permissions/Build-ID/pinned SQLite and four original hashes pass. This
renderer-only integration reuses the preceding full48/core1 evidence; no repeated
full suite or new CTest target. Fixed catalog stays 456/384/72.


## Compound Sum/Count arguments and render-scoped results (2026-11-04)

The Access expression parser now preserves source token offsets and passes the
exact balanced Sum/Count argument to a native per-row aggregate context. Arithmetic,
IIf and NULL arguments read the actual source row without moving the displayed
record. Bare Null/True/False/Empty are expressions rather than field names;
Count([*]) no longer acquires Count(*) semantics. Nested aggregates, missing source
context, incomplete rows, and compound arguments on empty sources fail explicitly.
Compound Sum rejects text and checks numeric overflow; full Access aggregate typing,
Currency precision and group scope remain open. Simple-field behavior is retained.

Paginated rendering caches at most 512 Number/Null/Error aggregate results within
the render scope, clears them on Requery/exit and retains no DAO/widget/object
references. Existing report tests check Sum(IIf(...))+Count(IIf(...))+Count(Null)
=31, a real SQL insertion/native Requery, then a second PDF and total52. Existing
original corpus independently compares three available report aggregate expressions
with SQL. The attempted compound-only original selector found no executable case;
it was replaced with an explicit available-expression check. Six original compound
calls still lack report-source context here; no original compound execution claim.

Validation: targeted core/layout checks and original DAO/report 1/1 (33.33 s)
before the render-cache addition, then final core/layout 2/2 (11.11 s), full desktop
48/48 (83.17 s), freshly rebuilt GUI-free core 1/1 (1.76 s), Debian payload/
permissions/Build-ID/pinned SQLite and four original SHA-256 checks pass. No new
CTest target or catalog ID. Fixed catalog 456/384/72; F085 remains open.

Reference: Microsoft Sum documentation (linked in PRINT_PREVIEW_ADAPTER.md),
including expression arguments and omission of NULL rows; native SQL is an
independent calculation, not a complete Access output oracle.


## Report page context and prepared print bands (2026-11-04)

Fixed-band output now plans each section with its record, destination position
and page, determines total pages, validates that section's own calculations and
records all its native vector commands before opening the output device. A
report-footer formula valid only on the final record is no longer refused due
to an error on an earlier unprinted row. The saved record/page context is restored.

Page/Pages expressions use planned page numbers; bounded native report getters
return Long/Integer according to Microsoft documentation. Outside actual paginated
rendering they fail explicitly. Setter semantics and Page-event execution remain
unsupported. PDF checks independently read 1/2 and 2/2 exactly once, and the
existing test exports a footer valid only on the second/final record. No complete
Access page/formatting/property/event model or output oracle is claimed.

Validation: desktop build and existing mainwindow/layout/original-DAO targets
3/3 (50.31 s), Debian payload/permissions/Build-ID/pinned SQLite and all four
original hashes pass. This GUI-only report change does not alter the headless
runtime; preceding full48/core1 validation remains recorded above rather than
being repeated. Fixed 456/384/72 and 48 CTest targets unchanged (F085 remains open).


## Actual report pages inside the opened document (2026-11-04)

Supported imported reports now open a real QPrintPreviewWidget directly in the
report document after initial events. Native page navigation/zoom/refresh and
Requery use the same prepared fixed-band renderer. The existing controls remain
as runtime backing in an explicit current-layout tab. A failed/unsupported render
disables the pages tab, selects the layout inspector and shows an incomplete-output
notice. Page/Pages reads use valid native preview metadata; Requery invalidates
old metadata before loading. The QPrinter outlives its preview widget.

Existing report cases check two document pages, native navigation and Page/Pages,
three pages after SQL insertion/Requery, and CanGrow fallback. The real modal
preview/PDF cases remain checked. This does not complete Access-2003 chrome,
window/print-event semantics, arbitrary runtime mutation invalidation, grouping,
CanGrow, direct-print view or original compound-report context. F085 stays open.

Validation: access_2003_layouts 1/1 (9.98 s), full desktop 48/48 (81.95 s),
Debian payload/permissions/Build-ID/pinned SQLite and four original hashes pass.
Headless runtime sources did not change; the preceding fresh core1 result remains
applicable without rebuilding it again. Fixed 456/384/72 and 48 targets unchanged.


## Native report paper and section backgrounds (2026-11-04)

Section BackColor Longs are now retained and painted by the actual native canvas
across the report content width. The same paint commands feed the layout view,
PDF and prepared paginated bands. Reports use a white paper base; forms retain
the existing gray base. RGB/system colors reuse the native OLE adapter. Runtime
Section objects/property mutation and full Windows color/profile/pixel parity
remain open. No original color-layout oracle is claimed.

The existing report-event case samples actual green/blue/red canvas pixels and
an unset white detail area. Desktop build, layout target1/1 (9.16 s), mainwindow/
original-DAO2/2 (36.88 s), freshly rebuilt GUI-free core1/1 (1.94 s), Debian
payload/permissions/Build-ID/pinned SQLite and four original hashes pass. The
preceding document integration full48 remains recorded above; no repeated full
suite for this color-only rendering change, no new target or catalog ID.
Fixed 456/384/72 unchanged.


## Native classic interface preference (2026-11-04)

The GUI defaults to the portable Qt Windows style when available, with eight-point
application font and installed Tahoma when present (otherwise existing family).
Ansicht/Klassische Oberfläche restores the original style/palette/font and saves
the preference through per-user QSettings. The main title now uses the product
name. Unset form surfaces inherit the live Window palette; explicit imported
colors and white report paper remain. No new dependency or Microsoft font bundle,
no full Office/Access-2003 chrome, MDI, pixel/Windows oracle or shortcut claim.

Existing transparent-surface checks were adapted from fixed Fusion gray to the
selected Window palette; explicit color checks remain. Initial classic-style GUI
run exposed that old palette assumption; no product color rule was suppressed.
Validation: desktop build; existing mainwindow/layout2/2 (11.78 s); final full48/48
under Windows style (81.16 s); actual application startup with isolated XDG
configuration/data until timed shutdown; Debian payload/permissions/Build-ID/
pinned SQLite and four original hashes pass. Headless sources unchanged from
preceding fresh core1. No new CTest target; fixed 456/384/72 unchanged.


## Owned reports use the native Report engine (2026-11-04)

Owned report definitions now prepare SELECT columns without stepping records,
generate a bounded automatic fixed-band layout, and open the existing native
Report/CodeBehind/page/PDF runtime. Explorer double-click runs the report; the
object menu edits SQL/title/VBA. Designer preview is a real nested native report
window that returns to retained edits. The old SQL/HTML/QTextDocument path was
removed. Raw multiline SQL is passed separately into the layout model, preserving
RecordSource and stored SQL/code. Extra definition metadata is copied on save.

Open/Load/NoData/Close procedures are auto-wired; unsupported initialization results
refuse complete rendering. Stored report lookup and main-window object reads/writes
use main._la_objects and reject imported-name ambiguity. DoCmd opens saved reports
with independent project lifetime. Automatic layout is limited to 1-32 columns,
may exceed paper width, and does not complete visual design, grouping, CanGrow or
print events. Existing ranked-TOP filtering remains 3251: the initial combined
TOP/Where test hit that known guard, so the evidence separates TOP execution from
WhereCondition on an ordinary SELECT. The guard was not weakened.

The existing mainwindow case creates the report through the actual editor, checks
exact multiline SQL/code, native TOP2 over three rows, typed Report/controls/module
state/PDF and TEMP-shadow immunity. Actual VBA opens an ordinary SELECT report
with a filter, closes its caller without closing the report, then edits/previews
with four balanced Open/Close SQL effects. No new CTest target or catalog point.

Validation: final mainwindow/layout2/2 (11.99 s), full desktop48/48 under Windows
style (82.41 s), Debian payload/permissions/Build-ID/pinned SQLite and four original
hashes pass. Headless runtime sources unchanged from preceding fresh core1.
Fixed 456/384/72 retained; full report and Access parity remain open.


## Native multiline report textboxes (2026-11-04)

Report textboxes keep the QLineEdit-derived native value/mask/control interface
but paint plain paragraphs and word wrapping with Qt text layout inside their
styled content area. Font, palette and horizontal alignment feed the same canvas
and vector PDF; displayText is authoritative, so this adaptation cannot paint raw
password text. Form editing behavior is unchanged. This remains fixed-height:
CanGrow/CanShrink, flowing controls, long-band page splits and Windows text/selection
metrics remain open.

The existing report case binds real multiline SQL data, preserves its raw value,
independently extracts both paragraphs and an automatic line break from PDF,
checks masked PDF lacks the synthetic clear text, then restores test data.
Validation: layout1/1 (9.46 s), final mainwindow/layout/original-DAO3/3 (46.27 s),
desktop build, Debian payload/permissions/Build-ID/pinned SQLite and four original
hashes pass. Headless sources unchanged; preceding full48/core1 retained without
another full run for this report-only painter. No new target or catalog item;
fixed 456/384/72 unchanged.


## Bounded detail CanGrow and native design option (2026-11-04)

TextBox/template and section CanGrow metadata are retained; native control reads
return the flag and runtime writes fail explicitly as design-only. Owned report
design saves a CanGrow option and emits it into native fields. The renderer uses
the actual Qt text layout height for one actually growing direct/unbound textbox
per detail record, shifts lower controls, grows the section across its width and
plans pages from variable detail heights. Scoped band visibility/background and
geometry/canvas/record restoration prevent later sections leaking into grown bands.

The long probe exposed Qt single-line display normalization: 44 paragraph breaks
in raw text became zero in displayText. Normal unmasked fields now preserve their
original text; password/other input masks still use their display string. Earlier
short multiline evidence had happened to wrap at that location; it did not prove
paragraph preservation. The corrected long probe does. Masked PDF checks remain.

Existing cases verify one fixed page versus two complete grown pages, final text
markers plus values moved below them, raw data/geometries/canvas restoration,
source flag reads through actual Report VBA, and refusal of an oversized detail
without replacing a valid PDF. Multiple simultaneously growing fields, overlap,
calculated growers, growing header/footer sections, CanShrink and splitting long
details remain open. No original CanGrow/Windows geometry oracle was executed.

Validation: mainwindow/layout2/2 (11.62 s), final layout1/1 (9.29 s), full desktop
48/48 under Windows style (78.58 s), freshly rebuilt GUI-free core1/1 (1.68 s),
Debian payload/permissions/Build-ID/pinned SQLite and four original hashes pass.
No new target or work-catalog point; fixed 456/384/72 unchanged. F085 stays open.
Reference and exact bounds: REPORT_CANGROW.md.


## Multiple detail CanGrow groups (2026-11-04)

Aligned growing fields share the greatest height delta; vertically disjoint
groups contribute cumulative movement of lower controls. Partial staggered
spans and actual control overlap remain refused. Existing SQL/PDF cases cover
parallel and stacked growth, complete end markers on both expected pages and
restored control geometry. This extends the preceding single-grow subset;
calculated growers, long-band splitting and full Access layout remain open.

Validation: desktop build, mainwindow/layout2/2 (11.71 s), final layout/original
DAO2/2 (42.54 s), Debian payload/Build-ID/pinned SQLite and four original hashes
pass. Core unchanged; preceding full48/core1 evidence retained. No new CTest
target or catalog ID; fixed 456/384/72 retained, F085 remains open.


## Native long-detail page splitting (2026-11-04)

Supported tall detail bands now split between actual Qt glyph/decorative-line
bounds, retaining their SQL row and repeating page headers/footers. Details
fitting one page stay whole. Fixed controls remain indivisible; no safe boundary
or the existing 3500-pixel canvas bound causes explicit refusal before output.
This does not complete KeepTogether/group scheduling, CanShrink, growing other
sections or a Windows print oracle. F085 remains open.

The native plain-text painter shares its shaped line geometry with pagination.
Each fragment emits only its visible lines and clips its native frame/text;
merely clipping a full QTextDocument had retained invisible duplicate PDF text.
Existing report cases now verify all 125 tokens exactly once in order for both
paragraph and automatically wrapped underlined SQL text, multiple native pages,
real Page/Pages, restored field geometry/canvas and preserved prior PDF on refusal.

Validation: final targeted layout1/1 (9.73 s), full desktop48/48 under Windows
style (79.54 s), freshly built GUI-free core1/1 (1.68 s), Debian payload/Build-ID/
pinned SQLite and four original hashes pass. Original DAO case passes (33.64 s).
No new CTest target or catalog point; fixed 456/384/72 retained.


## Detail ForceNewPage and owned report design (2026-11-04)

Native imported details retain and paginate all four ForceNewPage choices.
Before/after boundaries preserve row context and repeated page sections; the
report footer honors pending breaks without generating an empty trailing page.
The owned report editor saves/restores the selection and emits validated actual
metadata into the shared native runtime. Non-detail nonzero settings remain
explicitly unsupported. Full section VBA APIs and Windows page parity stay open.

Existing SQL/PDF cases check all three nonzero choices and row placement; actual
mainwindow design/save/reopen/nested-preview checks the persisted choice.
Validation: initial mainwindow/layout2/2 (12.78 s), final mainwindow/layout/original
DAO3/3 (47.29 s), desktop build. Previous full48 (79.54 s) and fresh GUI-free core1
(1.68 s) precede this isolated detail/layout change. Fixed 456/384/72; no new target
or point. F085 remains open. See REPORT_DETAIL_BREAKS.md.

Detail-break checkpoint: Debian payload/Build-ID/pinned SQLite and all four
original fixture SHA-256 checks also pass.


## Native Section identities and runtime detail page breaks (2026-11-04)

Form/Report.Section resolves existing indices 0–4 or retained imported names.
Weakly cached non-owning wrappers preserve name/index/Parent identity and guard
closed documents. Section/Access.Section participate in native typed slots and
TypeOf; the five standard section constants are available. Name, source Height,
CanGrow and detail ForceNewPage reads are bounded; Byte-validated detail writes
rebuild an existing real report preview. Other mutations and group/event/layout
scheduling remain unsupported. See NATIVE_SECTION_OBJECT.md for exact bounds.

The actual Report_Open procedure holds a typed Section and sets ForceNewPage;
the native preview has four pages. A later retained-object write resets it to one
page. Existing cases check metadata, identity, failed-write state preservation
and closed guards. The first run exposed Me.Section being routed as a control
lookup (2465); the parser now dispatches this indexed property explicitly without
changing Me! control lookup. The same actual VBA regression then passed.

Validation: final mainwindow/layout2/2 (12.41 s), full desktop48/48 under Windows
style (84.06 s), freshly rebuilt GUI-free core1/1 (1.76 s), Debian payload/Build-ID/
pinned SQLite and four original hashes pass. No new CTest target or catalog ID;
fixed 456/384/72 retained. Full F050/F085 remain open.


## Live native Section.BackColor (2026-11-04)

The weak native Section wrapper now reads/writes BackColor as an OLE Long.
Supported encodings are validated before mutation; actual section painting and
prepared print bands share the updated source value. Native canvas overrides
paint the full content width and existing report pages refresh. Unset reads expose
the actual native base RGB, not a guessed original system-color token. Full live
palette/profile equivalence and other Section mutations remain open.

The existing actual VBA Open/retained-procedure case now changes red to green,
checks native canvas pixels and source values, and verifies unsupported encoding
leaves the previous value intact. The PDF failure message also no longer claims
all growing fields are unsupported. Validation: desktop build and final mainwindow/
layout/original DAO3/3 (46.90 s). Core unchanged from preceding full48 (84.06 s)
and fresh GUI-free core1 (1.76 s); no redundant full run for this color adapter.
Fixed 456/384/72 retained; no new target or catalog ID.

Section-color checkpoint: Debian payload/Build-ID/pinned SQLite and all four
original fixture SHA-256 checks pass.


## Native report preview P/Ctrl+P/C (2026-11-04)

Ctrl+P opens the existing native full-report print dialog. Plain P and C are
scoped to the paginated preview/navigation pane; layout control text input stays
outside that scope. P respects print-action availability; C uses normal closing.
Existing report cases send actual keys to the native graphics surface, reject
both print dialogs without a job, preserve current data and close through C
before checking retained native-object guards. See PRINT_PREVIEW_ADAPTER.md.

Validation: desktop build and final mainwindow/layout2/2 (13.05 s). Earlier
Ctrl+P-only2/2 passed (12.54 s). The first plain-letter probe lost active surface
focus after a modal dialog; the corrected actual-surface regression passes.
No new CTest target or catalog ID; fixed 456/384/72. Core unchanged from prior
full48 (84.06 s) / GUI-free core1 (1.76 s). Full keyboard parity remains open.

Report-shortcut checkpoint: Debian payload/Build-ID/pinned SQLite and all four
original fixture SHA-256 checks pass.


## Report-header/footer ForceNewPage and title-page design (2026-11-04)

Imported header/detail/footer flags now feed the same native page plan. A header
after-break creates a title page; a footer before-break creates a summary page;
concurrent detail/before/after breaks coalesce on fresh pages. Native Section
reads/writes support all three report body sections and refresh current pages.
The owned report editor also persists header page-break choices, validated by
the adapter and read through actual Report VBA. Page sections, group scheduling,
nonzero zero-height header/footer rules and a Windows blank-page oracle remain
explicitly unsupported/open. See REPORT_DETAIL_BREAKS.md.

Existing real SQL/PDF cases check title/summary pages independently and together
with exact row/final-marker placement. Typed VBA sets/resets the three section
flags; mainwindow save/reopen/preview retains both selectors. Validation: desktop
build and mainwindow/layout/original DAO3/3 (47.84 s). Runtime core unchanged from
previous full48 (84.06 s) / fresh GUI-free core1 (1.76 s). Fixed 456/384/72; no new
target or catalog ID. F050/F085 remain open.

Report body-break checkpoint: Debian payload/Build-ID/pinned SQLite and all four
original fixture SHA-256 checks pass.


## Native report paper/orientation/margins and S (2026-11-04)

A functional native Qt dialog edits actual QPageLayout geometry for the open
report. Qt standard/custom paper, orientation and validated mm margins feed
existing pagination; preview/PDF/print copy the same layout. Cancel preserves it.
The initial QPageSetupDialog path failed on the PDF printer with no installed
physical printer. It was replaced by native controls backed by real page geometry,
not a fake printer. Original Access setup decoding/persistence and physical output
remain open; exact bounds are in REPORT_PAGE_SETUP.md.

Existing actual S/dialog cases accept landscape and a 25-mm left margin. PDF
MediaBox and independent text bounding coordinates verify physical output; a
cancelled orientation change preserves landscape in another PDF. Validation:
desktop build and final mainwindow/layout/original DAO3/3 (47.29 s). Runtime core
unchanged from earlier full48 (84.06 s) / GUI-free core1 (1.76 s). Fixed 456/384/72;
no new CTest target or catalog ID. Full report/page setup parity remains open.

Page-setup checkpoint: Debian payload/Build-ID/pinned SQLite and all four original
fixture SHA-256 checks pass.


## Native report F5/Z and fit-page navigation (2026-11-04)

F5 selects the real page-number editor and typed navigation waits for Enter.
Z toggles FitInView/native100% zoom. PageDown/Down and PageUp/Up navigate actual
pages only in FitInView, scoped to the graphics surface; other zoom modes retain
native scrolling. Existing actual Qt-key cases verify focus, deferred commit,
page changes and both zoom directions. Exact Windows zoom anchors and full
keyboard/window parity remain open. No new CTest target or catalog ID.

Validation: desktop build and mainwindow/layout2/2 (13.38 s). Fixed 456/384/72
retained. The full current suite and package are checked for this report milestone.

Navigation milestone: full current desktop48/48 under Windows style passes
(88.77 s), Debian payload/Build-ID/pinned SQLite and all four original fixture
SHA-256 checks pass. GUI-free runtime unchanged from fresh core1 (1.76 s).


## Saved native report page geometry and repair (2026-11-04)

Page setup optionally stores versioned geometry in an existing owned report
definition. Main-qualified, case-aware lookup and a conditional atomic update
preserve SQL/VBA/archive/extra fields and reject changed definitions without
committing caller transactions. Native report editing reloads latest extras so
nested preview page changes survive later Save. Unsaved/imported layouts retain
session setup; original Access printer/designer binary round trips remain open.

Stored geometry is validated before output. Invalid metadata selects explicit
layout fallback and blocks report PDF/print; actual page-setup repair clears it.
The existing mainwindow cycle saves/reopens geometry, verifies SQL/code and TEMP
shadow immunity, changes it in nested preview, then saves the editor and retains
the newer margin. A malformed orientation preserves a prior PDF and is repaired
through real controls. Five balanced Open/Close effects cover all document cycles.
Validation: desktop build, initial mainwindow/layout2/2 (12.76 s), final mainwindow
1/1 including invalid-state repair (2.81 s). Fixed 456/384/72, no new target/ID.
Full current suite/package are checked for this persistence milestone.

Stored-page-geometry milestone: full current desktop48/48 under Windows style
passes (83.50 s), Debian payload/Build-ID/pinned SQLite and all four original
fixture SHA-256 checks pass. GUI-free runtime unchanged from fresh core1 (1.76 s).


## Native numeric report Avg/Min/Max (2026-11-04)

Access expressions now aggregate finite native numeric fields/row expressions
with Avg/Min/Max. An online long-double mean avoids unnecessary sum overflow;
results remain bounded native Double/numeric values. A scalar-only accumulator
also preserves Sum/Count behavior, including direct Count without field coercion,
NULL omission, overflow/error propagation and render/Requery cache invalidation.
Text/Date/Boolean/Currency, full typing/rounding/group scope and Windows results
remain open, explicitly refused where unsupported. REPORT_AGGREGATES.md details
these bounds; no original Avg/Min/Max execution oracle is asserted.

Existing actual SQL/control/PDF cases compare the three aggregates and compound
Avg against independent SQL, update after insertion/Requery, handle all-NULL
values and preserve a valid numeric PDF on unsupported text minimum. Validation:
desktop build, mainwindow/layout2/2 (13.58 s), freshly rebuilt GUI-free core1/1
(1.99 s wall total). Full current suite/package are checked for this parser change.
Fixed 456/384/72 retained; no new target or ID.

Numeric-aggregate milestone: full current desktop48/48 under Windows style
passes (86.46 s), fresh GUI-free core1/1 (1.99 s wall total), Debian payload/
Build-ID/pinned SQLite and all four original fixture SHA-256 checks pass.


## Native report First/Last scalar selectors (2026-11-04)

Access report expressions now select the first/last actual loaded scalar value,
including NULL endpoints, with the existing row-expression binding. Supported
scalar values are copied; objects/arrays/binary values remain unsupported. Source
truncation/missing fields, nested aggregates and empty compound binding retain
their guards. Ordering is the actual loaded source order, not an invented Access
physical/insertion order. Full group/Jet type/order equivalence remains open.

The render cache additionally permits immutable typed scalars and text up to
4096 UTF-8 bytes, still at most512 entries (at most2MiB retained text), cleared on
Requery/render exit. No object, cursor or widget can be retained. Existing actual
SQL/control/PDF cases check ascending/descending endpoints, compound selectors,
updated text after insertion/Requery and actual VBA-control NULL values at both
ends. Validation: initial mainwindow/layout2/2 (13.08 s), fresh GUI-free core1/1
(1.60 s). Full current suite/package are checked for this parser/cache change.
Fixed 456/384/72; no new target/ID. Exact scope: REPORT_AGGREGATES.md.

First/Last milestone: full current desktop48/48 under Windows style passes
(87.00 s), fresh GUI-free core1/1 (1.60 s), Debian payload/Build-ID/pinned SQLite
and all four original fixture SHA-256 checks pass.


## Native global/Application field DLookup (2026-11-04)

DLookup now binds an actual field in a table/installed saved query, performs a
real bounded read-only TOP1 query and uses proven result-column types. Global
and Application calls accept positional/named arguments. Empty strings, NULL
matching values and no matches remain distinct; failed/unavailable sources do
not masquerade as empty lookups. Numeric/date/Boolean/Currency values use the
existing typed converter. Calculated Expr, explicit NULL criteria, unproven types
and complete domain/reference/collation/error parity remain open.

Existing actual VBA/calculated-control cases cover a real nullable typed table,
saved-query alias, apostrophe criteria, named calls, and dirty versus saved form
values. Currency units/type and Date/Boolean identities are checked. No lookup
cache or cursor/host retention is introduced. Initial desktop build and mainwindow/
layout2/2 pass (12.38 s); full current suite/package and fresh GUI-free runtime
are checked for this parser/Application change. Fixed 456/384/72; no new target/ID.
Exact scope and reference: NATIVE_DLOOKUP.md.

DLookup milestone: full desktop48/48 passes (84.24 s), fresh GUI-free core1/1
(1.55 s), Debian executable/Build-ID/pinned SQLite checks and all four original
fixture SHA-256 checks pass. No additional test target or catalog item.


## Typed report Min/Max (2026-11-04)

The native scalar accumulator now selects homogeneous Date and Currency extrema
without losing the selected subtype. Currency comparison uses signed scaled
integer units rather than Double. Mixed kinds, text and Boolean remain guarded.
Actual SQL-row CCur/CDate expressions retain Currency units and Date identity
through native control Value access and real PDF output. Desktop build and the
existing layout target pass (1/1, 10.86 s); no additional target or catalog ID.
Fixed 456/384/72; broad F085 remains open. See REPORT_AGGREGATES.md for limits.
Typed-extrema Debian payload/Build-ID/pinned SQLite and four original fixture
hashes also pass. The last full suite remains DLookup48/48 (84.24 s); this bounded
report-accumulator change used the existing targeted layout regression.

## Native field DMin/DMax/DAvg (2026-11-04)

Global/Application calls now share typed field/domain/criteria binding with
DLookup. Schema-only preparation exposes proven per-position types before user
rows are fetched. DMin/DMax preserve homogeneous Number/Date/Currency; DAvg is
Number-only. NULL observations are omitted; empty/all-NULL selections yield NULL.
The actual query executor enforces its existing row/budget limits without partial
results. Computed Expr, text ordering, Date/Currency averages and full Access
binding/error/result precision equivalence remain open. No original MDB writes.

Desktop build and existing query/model/SQL-value/layout3/3 pass (10.19 s); actual
VBA checks global/Application/named calls, predicates, NULL and Date/Currency
identity/value. Fresh GUI-free core1/1 passes (2.67 s). Full suite/package follow
for the shared parser/schema API change. Fixed 456/384/72; no new target or ID.
See NATIVE_DLOOKUP.md for exact scope and public references.
Domain aggregate milestone: full desktop48/48 passes (84.14 s), fresh GUI-free
core1/1 (2.67 s), Debian payload/Build-ID/pinned SQLite and all four original
fixture SHA-256 checks pass. No additional target or catalog item.

## Exact native report Currency Sum (2026-11-04)

Report Sum now accumulates homogeneous Currency units as signed integers and
retains Currency result identity. Overflow raises error6 before arithmetic;
mixed Currency/non-Currency inputs remain explicitly guarded. Existing Number
and Sum text-conversion behavior is preserved. SQL REAL storage bounds are not
weakened. Actual SQL decimal-string rows plus CCur reach INT64_MAX exactly;
another 0.0001 and Requery expose overflow6 and preserve the existing PDF.
Desktop build and existing layout1/1 pass (10.49 s total). No new target/catalog
item. Fixed 456/384/72; broad report/typing/oracle requirements remain open.
Exact report Currency Sum: Debian executable/Build-ID/pinned SQLite and all four
original fixture hashes pass. Last full suite remains domain aggregates48/48
(84.14 s), GUI-free core1/1 (2.67 s); bounded report change used layout regression.

## Shared Application DSum/DCount and preserved SQL result types (2026-11-04)

Application now exposes DSum/DCount with positional/named arguments over the same
SQL-expression executor as global calls. The scalar handoff retains proven SQL
column types, fixing lost Currency identity in DSum. Existing SQL precision,
expression and NULL-criteria behavior is retained. Actual VBA covers Currency
identity/value, non-NULL calculated CCur input, Count(*)/nullable field differences,
predicates, named arguments and all-NULL sum. The initial synthetic CCur(NULL)
expectation was corrected to explicitly filter NULL input; runtime error94 remains.
Desktop build and existing mainwindow/Currency-SQL/layout3/3 pass (15.79 s).
No new target/catalog item. Fixed 456/384/72; full Access coercion/reference/error
and original output-oracle requirements remain open. Scope: NATIVE_DLOOKUP.md.
Application Sum/Count: Debian payload/Build-ID/pinned SQLite and four original
fixture hashes also pass. Last full suite remains domain aggregates48/48 (84.14 s),
GUI-free1/1 (2.67 s); the shared native host path used existing targeted3/3.

## Native field DFirst/DLast (2026-11-04)

Global/Application field selectors use real query endpoints and proven native
scalar conversion, including NULL endpoints. DFirst uses TOP1; DLast requires a
complete result within the existing row/budget guards. An explicitly sorted saved
query supplies deterministic native evidence; underlying Access physical scan
order is not claimed. Existing actual VBA cases cover named calls, predicates,
empty selections, endpoint NULL and Currency identity. Computed Expr and full
Access binding/type/error/order-oracle parity remain open.
Desktop build plus core/layout2/2 pass (12.59 s); fresh GUI-free core1/1 (1.96 s).
Fixed 456/384/72; no new target/ID. See NATIVE_DLOOKUP.md for exact scope.
DFirst/DLast Debian executable/Build-ID/pinned SQLite and four original fixture
hashes also pass. Last full suite remains domain aggregates48/48 (84.14 s);
this bounded selector addition used core/layout2/2 and fresh GUI-free core1/1.

## Native sample report/domain statistics (2026-11-04)

Report Var/StDev and global/Application field DVar/DStDev now execute on actual
complete SQL rows through a shared Welford sample accumulator. NULL observations
are skipped; fewer than two values return NULL; n-1 is the sample divisor. Only
finite Number values are accepted. Long-double mean/central-moment updates and
Double result checks avoid naive squared-total cancellation. Currency/Date/text/
Boolean coercions, population functions and full Access error/rounding/grouping/
original-output oracle parity remain open.
Existing SQL-row report checks compare independent scalar SQL moments across
insertion/Requery, one observation and all-NULL input. Actual VBA domain calls
check sample values, named Application arguments and cardinality. Desktop build
and existing core/layout2/2 pass (13.19 s), fresh GUI-free core1/1 (1.79 s).
Full current suite/package follow for the accumulated domain/report changes.
Fixed 456/384/72; no new target/ID. Exact scope in REPORT_AGGREGATES.md and
NATIVE_DLOOKUP.md.
Sample-statistics milestone: full current desktop48/48 passes (86.40 s), fresh
GUI-free core1/1 (1.79 s), Debian payload/Build-ID/pinned SQLite and all four
original fixture hashes pass. No added targets or fixed-catalog IDs.

## Owned native report footer design (2026-11-04)

The report editor now stores an optional calculated footer expression and footer
ForceNewPage choice in the existing owned report definition. Native generation
emits an actual FormFooter and collision-safe calculated textbox using the real
report aggregate/pagination path. Metadata types, bounds and missing-footer breaks
are validated; empty expression/break0 retains the old layout. Existing SQL/VBA/
page geometry are preserved. This remains a fixed-height single calculated footer;
full report design/grouping, footer growth and original MDB round trips are open.
Existing actual main-window UI saves/opens a three-row Sum report, checks Value6,
two real pages with a footer-before break, exports PDF and reopens saved properties.
Desktop build and existing mainwindow/layout2/2 pass (12.93 s). Fixed456/384/72;
no new target/ID. Scope: NATIVE_REPORT_FOOTER_DESIGN.md.
Owned footer milestone: Debian payload/Build-ID/pinned SQLite and four original
fixture hashes also pass. Last full current suite remains statistics48/48 (86.40 s),
GUI-free1/1 (1.79 s); localized designer/generator change used mainwindow/layout2/2.

## Owned form query/SELECT data sources (2026-11-04)

The form editor accepts saved query names and one-line non-parameter SELECTs,
preparing actual field metadata without record evaluation. Query forms emit
read-only flags and omit automatic New/Save buttons; physical-table behavior is
unchanged. The common query-source reader now resolves current owned definitions
in main._la_objects as well as imported QueryDefs, with ambiguity/conflict/invalid
metadata guards. Saved query names remain RecordSource and Requery reads current
SQL. General query updateability/parameters/dependency-view installation and full
Access/original-layout parity remain open.
Existing actual UI checks SELECT metadata, saving/opening an owned COUNT-query
form, read-only value, source-table change plus Requery and reopening persisted
source/fields. Initial integration exposed the missing owned-query lookup; this
was implemented in the common reader. Desktop build and existing mainwindow/DAO/
query-model/layout4/4 pass (14.48 s). Full suite/package follow for the shared
query-source change. Fixed456/384/72, no new target/ID. Scope: NATIVE_QUERY_FORM_DESIGN.md.
README's obsolete single-page-only report sentence now reflects bounded native
multirow output and links the current report limitations.
Query-form milestone: full current desktop48/48 passes (88.67 s), Debian payload/
Build-ID/pinned SQLite and all four original fixture hashes pass. GUI-free core
last passed1/1 (1.79 s); this change affects the Qt data/GUI layer, not that runtime.
No added target or fixed-catalog ID.

## Owned SELECT dependency views and live revision refresh (2026-11-04)

The common query-view installer now includes owned main._la_objects definitions
in its revision and ordinary/ranked dependency handling. Forward references use
the existing resolver; parameter/action/unsupported SQL still do not become
executable views. Ambiguous owned/imported names receive no view. Non-table
form/report loads refresh the revision before reading, and a failed/busy refresh
is an explicit source error. Designer field preparation also refuses a busy
refresh instead of silently using old metadata.
The existing actual UI query form now traverses two owned queries with reverse
alphabetical dependency order. Application.Eval/DLookup reads the real base row;
changing the base SQL while open and calling Requery updates the dependent count
to zero. Desktop build and existing UI/DAO/model/layout/original-query5/5 pass
(23.20 s). Full suite/package follow for shared view/record-source behavior.
Fixed456/384/72; no new target/ID. Full Jet/parameters/original-order/result oracle
remain open. Exact scope: NATIVE_QUERY_FORM_DESIGN.md.
Owned dependency milestone: full current desktop48/48 passes (89.59 s), Debian
payload/Build-ID/pinned SQLite and four original fixture hashes pass. GUI-free
runtime unchanged, last core1/1 (1.79 s). No new target or fixed-catalog ID.

## Native DAO Byte() field values (2026-11-04)

DAO Field.Value now reads actual SQLite BLOBs as copied Byte() values, including
empty BLOB versus NULL. Native physical BLOB Field.Type is11. Edit/AddNew stage
Byte() only for declared binary fields, using a non-NULL empty QByteArray for an
empty binary. Unsupported scalar/binary coercions fail before mutation; existing
scalar conversion and closed/transaction/cursor guards remain. No pointer retention.
Actual VBA reads/modifies/saves exact binary bytes, reads pending Value, stores an
empty binary, cancels an edit and checks held Field after Close3420. Independent
SQL verifies hex, blob type, empty length and NULL. An initial property-array
indexing probe was replaced with an explicit Byte() variable read; no API was
widened to bypass property argument checks.
Desktop build and final existing DAO1/1 pass (0.53 s); existing mainwindow/layout
also passed on this native code before the test-only indexing correction. No added
target/ID; fixed456/384/72. OLE extraction/rendering, chunks, binary parameters and
original-format round trips remain open. Exact scope: DAO_BINARY_FIELDS.md.
DAO binary-value Debian payload/Build-ID/pinned SQLite and four original fixture
hashes pass. Last full suite remains owned dependencies48/48 (89.59 s), GUI-free
runtime unchanged, last1/1 (1.79 s). Binary changes used targeted existing cases.

## Stored native binary FieldSize (2026-11-04)

DAO Field.FieldSize exposes Long byte length for the current non-NULL binary
observation in browse mode, using the stored QByteArray size without a copied
Byte() allocation. Empty BLOB size0, Long identity and read-only383 execute in the
existing actual VBA binary case; SQL bytes/type/length checks remain. NULL, pending
Edit/AddNew size semantics, Memo/codepage sizes and original Jet/OLE storage sizes
remain explicitly unsupported. Desktop build and existing DAO1/1 pass (0.57 s).
Fixed456/384/72; no added target/ID. Reference and scope: DAO_BINARY_FIELDS.md.
Stored FieldSize Debian payload/Build-ID/pinned SQLite and four original fixture
hashes pass. Last full suite remains owned dependencies48/48 (89.59 s); GUI-free
runtime unchanged, last1/1 (1.79 s). This property addition used existing DAO1/1.

## Native Access FileDialog pickers and explicit Office identities (2026-11-04)

Application.FileDialog3/4 now uses actual modal Qt file/folder selection, native
Long Show-1/0, configured title/initial path/button/filter/index/multi-file mode,
Filters Clear/Add/Count/Item and read-only SelectedItems with one-based/default
items and For Each. Access-unsupported Open/SaveAs remain refused. Per-document
picker state persists; paths are actual selected values, not mock results. Weak
context guards protect retained objects after closure. Four explicit Office class
identities join the closed registry; typed slots and TypeOf use those identities.
The actual VBA case selects a real temporary filtered text file, enumerates its
path, cancels, selects a real directory and checks retained collection closure2450.
Initial core/layout2/2 pass (13.28 s), fresh GUI-free core1/1 (1.50 s). After completing
Filters read-only metadata, full current desktop48/48 passes (91.42 s), Debian
payload/Build-ID/pinned SQLite and all four original fixture hashes pass.
No added target/ID; fixed456/384/72. Full Office/COM/Windows Access2003 singleton,
default/case/error/CWD/namespace/lifetime parity remains open. Scope/references:
NATIVE_FILE_DIALOG.md. This full run also includes native binary Value/FieldSize.

## Owned named QueryDef SQL persistence (2026-11-04)

Named QueryDefs retain canonical query name/origin. Own SQL changes compile before
an atomic conditional main-catalog update, preserve other JSON metadata and respect
user transactions. Fresh owned members see saved SQL; SQL changes invalidate held
parameter generations. Existing imported per-instance SQL and temporary QueryDef
behavior remain distinct and do not overwrite archived sources. Workspace guards
now protect QueryDefs/parameters, and non-table DAO recordsets refresh dependency
views before loading. Rename/full QueryDefs collection/imported overlay and same-
name drop/recreate identity remain open.
Actual VBA/SQL checks canonical Name, parameter execution, persistent SQL visible
to another instance, stale parameter3420, external parameter-schema change, live
base-query revision, real SQL-metadata rollback and closed getter3420. Independent
SQL verifies unrelated JSON archive marker and unchanged source row count. Desktop
build and existing UI/DAO/model/layout4/4 pass (13.38 s). Full suite/package follow
for shared query-source/DAO workspace behavior. Fixed456/384/72; no new target/ID.
Exact scope: OWNED_QUERYDEF_SQL.md.
Owned QueryDef persistence milestone: full current desktop48/48 passes (82.63 s),
Debian payload/Build-ID/pinned SQLite and four original fixture hashes pass.
GUI-free runtime unchanged, last core1/1 (1.50 s). No added target or catalog ID.

## Native DAO QueryDefs catalog and owned lifecycle (2026-11-04)

QueryDefs Count/Item/Refresh/For Each reads real owned/imported catalogs, exposes
explicit DAO.QueryDefs identity and native zero-based case-folded name order (no
Windows-order claim). CreateQueryDef supports real named creation, optional initial
SQL and temporary empty-name instances. Own Delete/Create use unique savepoints
and monotonic workspace namespace generations; old references cannot retarget after
native deletion/recreation. Imported deletion/overlay, rename/append/attributes,
external metadata identity and full Access oracle remain open.
A real regression exposed stale managed views blocking same-name recreation.
Creation now distinguishes proven managed temporary views from namespace collisions.
Workspace transactions snapshot/restore query-view registry metadata with SQLite
schema rollback, avoiding restored views retaining obsolete SQL. Actual VBA/SQL
covers catalog count/name/index/enumeration/identity, named creation, duplicate3012,
creation rollback, optional SQL execution, deletion/recreation3420 and a dependency
view refreshed after delete/execute/rollback/edit. Final existing UI/DAO/layout3/3
pass (13.78 s); fresh GUI-free core1/1 passed (1.66 s) for the new closed DAO type,
which the subsequent Qt data lifecycle changes did not modify. Full suite/package
follow for shared workspace/schema behavior. Fixed456/384/72; no new target/ID.
Scope: DAO_QUERYDEFS_CATALOG.md and OWNED_QUERYDEF_SQL.md.
DAO catalog milestone: full current desktop48/48 passes (83.20 s), fresh GUI-free
runtime core1/1 (1.66 s), Debian payload/Build-ID/pinned SQLite and four original
fixture hashes pass. No added target or fixed-catalog ID.

## Native designer Ctrl+S and classic menu order (2026-11-04)

Owned query/form/report Ctrl+S invokes the real save validator/persistence path
without closing the editor, shows saved-name feedback and tracks the saved native
context. Explicit save buttons keep save-and-close behavior. Subsequent query saves
preserve other definition keys; report saves use the saved-name context to retain
latest nested-preview page geometry. Runtime form/data saving is unchanged.
Menus order Datei/Bearbeiten/Ansicht/Einfügen/Fenster, with real creation actions in
Einfügen and selected-object form/report editing in Ansicht. Main-window New/Open
have actual window-scoped standard shortcuts; existing preview/modal/project guards
still execute. No placeholder commands or full Access2003 visual/keyboard claim.
Existing UI checks activate real modal editors before Ctrl+S, read actual saved
SQL/VBA definitions while the editors remain open and verify repeated query saves
preserve an externally added metadata marker. Execute returns actual SQL value43.
Final current mainwindow1/1 passes (2.63 s); existing layout target passed during
this localized GUI change. Last full backend suite remains catalog48/48 (83.20 s),
GUI-free runtime unchanged, last core1/1 (1.66 s). No new target/ID; fixed456/384/72.
Scope: DESIGNER_SAVE_KEYS.md.
Designer keyboard/menu milestone: Debian payload/Build-ID/pinned SQLite and four
original fixture hashes pass. Final UI1/1 (2.63 s); last full catalog48/48 (83.20 s)
and GUI-free core1/1 (1.66 s) remain the broader runtime baseline.

Begrenztes natives VBA-/SQL-Partition mit benannten Argumenten und tatsächlicher
SQL-Gruppierung ist implementiert: [Umfang und Grenzen](NATIVE_PARTITION.md).
Vollstand danach48/48 (84.06 s), GUI-freier Core1/1, Paket und vier Originalfixture-
Hashes bestanden; fester Arbeitskatalog unverändert456/384/72.

Typed make-table milestone (2026-11-05): full current desktop48/48 passed
(87.98 s); Debian payload/build-ID/pinned SQLite3.54.1 and all four original
fixture hashes passed. GUI-free core runtime unchanged from fresh1/1 (2.36 s).
Original-query coverage remains78/100 without Access result oracle. No new
CTest target or catalog ID; fixed456/384/72. Native SQL scope: NATIVE_VBA_SQL.md.

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


## Native datasheet output and new source comparison (2026-11-06)

Tables and readonly/editable saved-query documents now print actual checked
snapshots and export paginated PDFs atomically. Visible header order, filter/sort,
complete bounded fetching, dirty/binary/Unicode refusal, modal/project/model
lifetime guards and existing-file preservation are implemented. A Qt percentage-
width table initially protruded by its grid padding; explicit point width now
reserves both edges and a live frame-boundary guard prevents actual clipping.
Scope: docs/DATASHEET_PRINT_RUNTIME.md. No physical printer job was accepted.

Final current desktop48/48 (92.62 s), fresh GUI-free core1/1 (1.90 s), current
Debian payload/Build-ID/pinned SQLite3.54.1 and four original fixture hashes pass.
Existing mainwindow target additionally prints both actual recipe tables/39 rows
from the supplied identical MDB, checking text/Null and unchanged original hash.
No new CTest target or fixed ID; F080/F148/F087 remain open,456/384/72 fixed.

The newly supplied Kassa MDB is a separate source variant, not a replacement for
fixtures/raw: SHA-256 d9ff0f9d002b3ab59b3336a34e6c09d9c058d5b9b944b754259a78e277e77102,
31227904 bytes. Same exported schema;28/29 tables have identical CSV bytes.
The sole change is tblAnmeldung: upload0 rows, previous fixture1 row. No login
state or credentials are manufactured. Its isolated copy passes the existing
full table/cell regression and original query inventory:78/100, still without
Access execution/printing oracle. Source comparison logs/copies stay private
under /workspace/work/reference-20261106. Keks MDB and supplied Bridge ZIP are
byte-identical to the existing fixtures; originals remain unchanged.

Remote CI is currently blocked before job execution by GitHub account payment/
spending limits (c0ed5d1, run42221275422); no fresh remote matrix pass is claimed.
Codex Call was read from the actual feat/ask-via-call fd2cea7 checkout. The single
reference-access request received a call ID but ended timeout with no API answer
or confirmed decision; no delivery/conversation claim. User subsequently supplied
source databases in chat. No completion call: the overall parity task is open.


## Datasheet preview and VBA print commands (2026-11-06)

The same checked snapshot now drives a genuine modal QPrintPreviewWidget page
view with actual page count, page navigation and zoom. Its dialog owns the widget
and is destroyed before its printer. It offers no uncontrolled print-to-file
shortcut bypassing the atomic output controller. Failed rendering disables the
page view and preserves an explicit failure; source/project closure cannot be
reported as successful output. The original project/reentry/draft guards remain.

DoCmd.RunCommand acCmdPrint and Application.RunCommand acCmdPrintPreview address
the active table/query through the existing dispatcher, preserving caller Me.
Microsoft's actual AcCommand enumeration was read: Print340, PrintPreview54.
Unsupported form/report RunCommand coverage remains unchanged. The renderer now
uses a 72-dpi measurement device for true 9-point document metrics, independent
of screen DPI; an extracted PDF glyph-height assertion checks the actual result.

Existing core/mainwindow/layout3/3 (18.86 s), fresh GUI-free core1/1 (1.77 s),
current Debian payload/Build-ID/SQLite3.54.1 and four original hashes pass.
Actual VBA Print writes a private PDF, then continues in its caller; named
Application preview navigates real pages and continues with its own Tag intact.
The last single full desktop48/48 (92.62 s) is the preceding40c6b13 output commit,
not a new full-suite claim for this increment. No new target or ID;456/384/72 fixed.


## Native Recordset.Type and local platform validation (2026-11-06)

Recordset.Type exposes the selected native dbOpenDynaset2/dbOpenSnapshot4 kind
as a readonly Integer. SQL-result readonly policy is separate from opening kind;
Database/QueryDef factories now carry the actual mode instead of reinterpreting
readonly as Snapshot. Clones retain that mode. Frozen form cursor adapters lack
proven DAO type and continue to refuse3251. Invalid arguments450, writes383,
closed handles3420; a real field named Type remains available through rs!Type.
Existing DAO execution1/1 (1.02 s), rebuilt current app/all-target relink, current
Debian payload/Build-ID/SQLite3.54.1 and four original hashes pass. GUI-free runtime
is unchanged from1/1 (1.77 s). Full table-type/default/locking/ADO/Access-oracle
semantics remain open under F054/F055. Original32-module inventory has zero dot-
Type spellings; no additional original business execution is claimed.
Reference: Microsoft Recordset.Type (DAO), readonly Integer, and
QueryDef.OpenRecordset (DAO) query/dynaset default; current documentation is not
an Access2003 execution oracle.

Independent immutable e2d64c9 source copies now pass full local Debian12/Qt6.4
48/48 (101.29 s) and Ubuntu24.04/Qt6.4 48/48 (93.65 s), including native preview/
VBA print and actual recipe PDF proof. Both .deb packages pass payload/Build-ID/
SQLite3.54.1 checks, actual installation/backend smoke and four original hashes.
These are local container results, not GitHub CI successes; account payment/
spending-limit rejection remains. The newer Recordset.Type increment has targeted
host evidence, not yet a new full platform-matrix claim.456/384/72 unchanged.


## Checked OpenRecordset mode arguments and latest verification (2026-11-06)

Database/QueryDef/TableDef.OpenRecordset now share checked native Long mode
conversion: round-to-even before selection, native omitted/Missing mode2,
explicit Empty0 refused, Null94, invalid text13 and unsupported modes3251.
Database/QueryDef arity errors are450. This removes Database's truncating cast,
QueryDef's inconsistent exact-Double comparison and TableDef's narrower Integer
conversion. Existing native supported modes remain2/4; table-type1, full Access
workspace defaults, options/locking/ADO and complete error priority remain open.
Existing DAO target1/1 (0.97 s), current app/DAO build, current .deb payload/
Build-ID/pinned SQLite3.54.1 and four original hashes pass. All three supplied
upload hashes also remain unchanged. No new target/checkbox/ID; F055/F122 open.

The preceding418719d runtime passed a fresh Debian13 full48/48 (94.15 s).
Debian12 and Ubuntu24.04 immutable418719d follow-ups rebuilt the actual app/DAO,
passed DAO1/1 (1.55 s /1.06 s), regenerated and installed their matching packages,
verified SQLite3.54.1 and all four original hashes. Their preceding e2d64c9 full
suites remain48/48 each; the final bounded mode-conversion increment above has
host-targeted evidence rather than an invented repeat of those full suites.
GitHub account CI remains blocked; local container results are distinguished.

TableDef.RecordCount: DAO/originales Kassa-Ziel2/2 (53,71 s) bestanden, einschließlich unabhängiger Zählung aller29 Originaltabellen. Aktuelle App/Paket/Build-ID/SQLite3.54.1 und vier Originalfixture-Hashes bestanden. Bestehende unveränderte Syntaxinventur separat187/187 akzeptiert (1,12 s), keine Ausführungsaussage. Kein neuer Vollsuite-/Matrixlauf dieses Inkrements.456/384/72 unverändert.

Binär-GetChunk: bestehendes DAO-Ziel1/1 (1,23 s), aktuelle App/Paket/Build-ID/SQLite3.54.1 und vier Originalfixture-Hashes bestanden. Nur gespeicherte positive native Binärbereiche; AppendChunk/OLE/Memo/Null/Randsemantik bleiben offen. Letzte Vollsuite a00eb29:48/48 (98,29 s), nicht diesem Inkrement zuschreiben.456/384/72 unverändert.

Binär-AppendChunk-Nachweis: finales bestehendes DAO-Ziel1/1 (1,38 s), danach vollständiger aktueller Debian13-Desktop48/48 (98,95 s), aktuelles Paket/Build-ID/SQLite3.54.1 und vier Originalhashes bestanden. Byte-Value-Zuweisung bleibt separat geprüft; Mischverwendungsfehler bewahren den Puffer. Kein neuer Test/Katalogpunkt.
Unveränderlicher629b3c9-Folgecheck: Debian12 DAO1/1 (1,49 s), Ubuntu24.04 DAO1/1 (1,27 s), jeweils aktuelle App/DAO neu gebaut, passende Pakete erzeugt/tatsächlich installiert, SQLite3.54.1 und vier Originalhashes bestanden. Diese629-Nachweise enthalten noch nicht die anschließenden TableDef.RecordCount-/Binärchunk-Inkremente; keine neue Vollmatrix/GitHub-CI behaupten. Letzte früheren vollständigen lokalen Distributionssuiten bleiben e2d64c9 mit48/48 je Plattform.456/384/72 unverändert; kein Gesamtabschluss.

Recordset.Fields-Fortsetzung: benannte/geprüfte Indizes, kanonische Ergebnisnamen, gehaltene Collection-/Field-Identität auch bei Enumeration, Null94/fehlend3265/Arity450 und explizite Doppelspalten-Verweigerung3251. Bestehendes DAO-Ziel1/1 (1,01 s), aktuelles Paket/Build-ID/SQLite3.54.1/vier Originalhashes bestanden; Details docs/DAO_RUNTIME.md. Kein neuer Vollsuite-Lauf dieses Inkrements.
Unveränderliche41b193c-Folgeläufe bestanden Debian12 DAO1/1 (1,61 s) und Ubuntu24.04 DAO1/1 (1,21 s), jeweils aktuelle App/Paket/Installation/SQLite3.54.1/vier Originalhashes. Diese Belege enthalten TableDef.RecordCount/GetChunk/AppendChunk, nicht die danach folgenden Field.Name-/Recordset.Fields-Inkremente. Letzte komplette Host-Suite41b193c:48/48 (98,95 s).456/384/72 unverändert; GitHub-CI weiter gesperrt.

Datenblatt-Seiteneinrichtung: funktionierender gemeinsamer Qt-Geometriedialog, Sitzungslayout und explizite versionierte Projektpersistenz, native RunCommand PageSetup32, Papier/Ränder für Vorschau/Druck/PDF. Conditional Savepoint-Speicherung bewahrt konkurrierende Änderungen und Aufrufertransaktionen; ungültige gespeicherte Geometrie blockiert PDF bis zur expliziten Reparatur. Details docs/DATASHEET_PRINT_RUNTIME.md. Ziel1/1 (3,64 s), vollständiger aktueller Desktop48/48 (89,86 s) inklusive Originalrezepte2 Tabellen/39 Zeilen, Paket/Build-ID/SQLite3.54.1/vier Originalhashes bestanden. Keine neue Distributionsmatrix/GitHub-CI;456/384/72 unverändert, F080/F087/F148 bleiben offen.

Druckexemplare: reale Datenblatt-PDFs mit1..100 Kopien, sortierter/unsortierter und vorwärts/rückwärts Seitenfolge; native Qt-Treiberkopien getrennt von manuellen Kopien. Vorhandenes Ziel1/1 (3,99 s), aktuelle App/Paket/Build-ID/SQLite/vier Originalhashes bestanden. Letzte Vollsuitee8375e8:48/48 (89,86 s), dessen GUI-freier Kern1/1 (2,20 s). Unveränderlichee837-Folgeprüfungen Debian12/Ubuntu24.04:2/2 (6,65 s /5,65 s), passende Paketinstallation/Backend/vier Hashes; noch ohne Kopien-Inkrement. Details docs/DATASHEET_PRINT_RUNTIME.md.456/384/72 unverändert, keine physische Drucker-/Windows-Parität behaupten.

Auswahldruck: vollständige markierte sichtbare Datensätze als eingefrorener eigener Text-Snapshot, reale Tabellen-/Readonly-Abfrage-PDFs mit unabhängig geprüfter Ein-/Ausschlussmenge. Leere/teilweise Zellmarkierungen verweigern3251 ohne Ganzdatenblatt-Fallback. Ziel1/1 (4,15 s), aktuelle App/Paket/Build-ID/SQLite3.54.1/vier Originalhashes und drei Uploadhashes bestanden. Details docs/DATASHEET_PRINT_RUNTIME.md. Kein neuer Vollsuite-/Matrixlauf dieses GUI-Inkrements;456/384/72 unverändert.

DAO-DDL-Identität: beobachtete eigene Database/QueryDef.Execute CREATE/DROP und SELECT INTO führen monotone Main-Tabellengenerationen; alte TableDefs/Fields/physische Recordsets/Clones verweigern3420 statt neu erstellte gleichnamige Daten zu ändern. Gleiches SQL/Rootpage/Rowid, alter Schreibpuffer, frische Wrapper, No-op/Fehl-DDL, ALTER und Workspace-Rollback unabhängig geprüft. Externes/unbeobachtetes SQL, TEMP-only und volle Jet-Locks/Lebensdauer bleiben offen. Vollständiger aktueller Desktop48/48 (89,70 s), App/Paket/Build-ID/SQLite/vier Originalhashes bestanden; schließt Kopien/Auswahldruck und Originalrezepte ein. Details docs/DAO_RUNTIME.md.456/384/72 unverändert.

Jet-Rohbinärimport: ausdrücklich gewählter MDBTools-Hexcodec (kein Raten aus CSV), exakte BLOB-/NULL-Werte und Byteanzahl-Vorschau im realen Teil-/Gesamtimport. Laufend begrenzte Child-Ausgabe20MiB/Diagnostik64KiB,30s; großer Export bricht atomar ohne Quelldateiänderung ab. Technische echte Jet4-MDBs mit Jackcess5.0.4 unabhängig geschrieben/zurückgelesen, keine Geschäftsdaten und keine neue App-Java-Abhängigkeit. Decoder-/GUI-/DAO-Bytes/Typ/Chunks/Hash/Rollback geprüft. Ziel2/2 (12,26 s), vollständiger aktueller Desktop48/48 (91,47 s), neue Kassa-Kopie Import/Zellwerte und78/100 Abfragen, App/Paket/Build-ID/SQLite/vier Originalhashes bestanden. Details docs/JET_BINARY_IMPORT.md; keine OLE-Anzeige/Anhänge/Originalformat-Schreib- oder Windows-Parität behaupten.456/384/72 unverändert; kein neuer CTest.

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

App und Hauptfensterintegration teilen jetzt dieselbe statische Desktopbibliothek
(MainWindow/Inspector/BridgeImporter und Qt-Metaobjekte), ohne zusätzliche Runtime-
DSO. Hauptfenster wird im gemeinsamen Build nur einmal kompiliert. GUI1/1 (5,25s),
GUI-freier Kern1/1 (1,93s), App mit SQLite3.54.1/JOIN-/Null-/Counter-Backendcheck
bestanden; kein frischer Vollsuite-/Distributionslauf.456/384/72 unverändert.
Fortsetzung: ISO-#Datum#-Defaults und führendes= bei begrenzten Literaldefaults
im Tabellen-/Feldentwurf und DAO; danach weitere offene Anforderungen fortlaufend.

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
