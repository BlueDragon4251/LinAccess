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
