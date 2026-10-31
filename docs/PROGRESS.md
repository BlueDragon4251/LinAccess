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
| Original Form_Load control initializer / verified property assignments | 1 / 6 |
| Original report sources classified / currently bound | 9 / 6 |
| Original report non-NULL textbox values compared with prepared reader | 54 |
| Original form layouts scanned | 47 |
| Parsed controls / visible command buttons | 878 / 86 |
| Original query artifacts classified / executed without Access result oracle | 100 / 49 |
| Current local regression tests | 17 passed |

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
The complex original Form_Load is classified as blocked at embedded-subform
Requery (438), not counted as successful execution.
