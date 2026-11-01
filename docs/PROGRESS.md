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
| Current local regression tests | 18 passed |

Current completed work package: positional Sub/Function calls, ByVal/ByRef aliases,
function expressions, top-level label error handling and original keypad regression.

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

Latest local validation: 18/18 regressions pass, the Qt-free runtime build/test
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
