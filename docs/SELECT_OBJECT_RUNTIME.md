# Native DoCmd.SelectObject

SelectObject now dispatches through the native VBA object/command path. Supported
ObjectType values are Table (0), Query (1), Form (2) and Report (3); the matching
acTable/acQuery constants are available alongside existing form/report constants.
Macro/module constants are recognized but those targets explicitly reject 3251.
Arguments resolve scalar defaults once, including optional empty object-name
positions. Null is 94, wrong arity/missing required name 450, scalar type mismatch
13 and numeric argument overflow 6. Exact Windows error precedence remains open.

With InDatabaseWindow=True, the actual project explorer preflights the object by
stored kind/name metadata, then validates pending control input and focuses/
selects/scrolls the real item. Names are case-insensitive; omitted names select
the requested kind group. Imported query/form/report artifacts now appear in
their corresponding groups, retaining source status and inspector actions.
Other source artifacts remain separately inspectable. Display text is not
parsed to infer identity. Missing targets report native 7874, ambiguous source/
user-defined names 5, modal blocking 2046 and cancelled control validation 2501.
No arbitrary duplicate or renamed alias is chosen.

With the default False, already-open native Forms/Reports can be selected, shown
if hidden, restored if minimized and activated. This does not open a missing
object. Missing forms/reports report 2450/2451. Table/query window activation
remains unsupported (3251) pending the broader native table/query window model.
Modal blocking and pending-control validation retain explicit failures. These
paths do not call form-record save; ordinary validation/activation handlers can
perform their own actions.

## Evidence

Actual production-GUI integration on temporary SQLite projects executes VBA
SelectObject calls to tables, a saved query artifact and imported form nodes,
including omitted-name group selection, mixed-case names and hidden-form
restoration. It switches between two open native Form/Report windows without
saving the dirty form, verifies cancellation retains editor input, and checks
missing/type/arity/modal/unsupported-target errors. A scalar-name object is read
once across UI preflight/activation. An intentionally colliding user-defined
form/source artifact name rejects 5 without guessing. Core/GUI/layout/original-
heavy targeted tests pass 4/4 (41.01 s). Full validation is in PROGRESS.md.

The original 32-module lexical inventory has zero SelectObject references; no
new original business execution is claimed. This is a bounded runtime adapter,
not complete database-window/MDI/schema-object editing or Windows focus parity.

Public contract:
https://learn.microsoft.com/en-us/office/vba/api/access.docmd.selectobject

The current documentation calls the Boolean InNavigationPane while describing
selection in the Database window; the native target remains the Access 2003
database-window operation.
