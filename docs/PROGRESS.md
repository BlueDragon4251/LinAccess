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
| Parsed controls / visible command buttons | 879 / 86 |
| Original query artifacts classified / executed without Access result oracle | 100 / 77 (72 ordinary + 5 parameterized) |
| Original date-filter fields exercised through actual calendar selection | 2 |
| Original generated record-search procedures / native selections / missing control-reference failures | 2 / 2 / 1 |
| Original public Eval function executions on prepared native numeric/Null states | 20 |
| Current local regression tests | 20 passed |

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
