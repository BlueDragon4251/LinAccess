# Native DAO execution subset

The C++ runtime now carries object references and dispatches members to a
SQLite-backed Database/Recordset/Field implementation. This is actual execution
against the open LinAccess project, not procedure recognition. No original MDB
is written and no additional per-statement confirmation is introduced.

Supported and tested:

- Positional Sub/Function calls with and without `Call`, value-returning functions
  inside expressions, ByVal and ByRef parameters (including aliased arguments).
- `Set db = CurrentDb()` and `CurrentDb.OpenRecordset(...)`, database references,
  `Nothing`, `Is`, multiple Dim declarations and DAO-qualified variable types.
- Table dynasets (`dbOpenDynaset`) and read-only SELECT/snapshot sources
  (`dbOpenSnapshot`). Results are materialized, limited to 100,000 rows.
- `EOF`, `BOF`, `RecordCount`, `MoveFirst/Last/Next/Previous`, `Requery`, `Close`.
- `rs!Field`, `rs.Field`, `rs.Fields(name/index).Value`, Field.Name, NULL values.
- `Edit`, `AddNew`, `Update`, `CancelUpdate` and `Delete`, using bound SQL values.
  Table updates identify the actual SQLite rowid; insert buffers contain only
  explicitly assigned fields, so available SQLite defaults remain effective.
- `Do While ... Loop` and `For Each ... Next` over native enumerable objects,
  including nested numeric/enumeration loops and `Exit For`. Empty enumerations
  retain the prior control-variable value. Error-handler Resume restores loop state.
- DAO `Fields`, `Fields.Count`, `Fields.Item(name/index)` and enumeration of actual
  Field objects. Field.Value remains bound to the recordset edit buffer.
- Object method arguments without parentheses, including `Err.Raise 11`.
- `DoCmd.Quit` validates native Access documents and honors Unload cancellation
  before Qt application termination (DOCUMENT_LIFETIME_RUNTIME.md).
- Variant Empty versus Null, Null-aware Boolean logic, numeric bitwise And/Or/Not,
  comparison-before-Not precedence, CLng tie-to-even rounding and Long overflow,
  Null propagation for Len and explicit invalid-Null errors for CStr.
- Procedure-local `On Error Resume Next`, `On Error GoTo label/0`, `Resume`,
  `Resume Next`, `Resume label`, `Err.Number`,
  `Err.Description`, `Err.Source`, `Err.HelpFile`, `Err.HelpContext`, `Err.Clear`
  and `Err.Raise(Number, Source, Description, HelpFile, HelpContext)` with named
  and omitted optional arguments. Writable Err fields use native Long/String
  coercion; signed Long custom codes and the Long vbObjectError constant are
  preserved through nested native calls and event-function failures. Explicit
  Resume and Clear reset these fields. Omitted optional Raise fields retain
  their current Err values; omitted/empty descriptions use a generic diagnostic
  when no stored description exists. Help metadata does not launch Windows Help.
  Full Access project-name/default localized message catalogs, HRESULT mapping,
  handler-state equivalence and debugger behavior remain open under F051.
  Division-by-zero, Long overflow, invalid Null and unset object errors carry
  their explicit codes; other host errors currently use generic error 5.

## Real Kassa regression

`real_kassa_dao_event` uses the original `Tagesende oder Ende` form and original
`Form_Tagesende oder Ende` module from the hash-verified 192-object Bridge corpus.
A real Qt click on the control `SF Ende Kassa` invokes the complete unchanged
`SF_Ende_Kassa_Click` handler, opens its actual table in a separate SQLite copy,
adds/updates a record, closes the recordset and requests application exit. The
test verifies the new row and its assigned operator/time fields. Inputs are
synthetic test inputs; original source/fixture bytes remain unchanged.

Hidden controls are instantiated for VBA lookup but remain invisible. Spaces in
button control names map to underscores in VBA event procedure names, as required
by this original form. This is evidence for one actual handler, not for all 187
procedures or for complete operation of the Kassa application.

## Remaining limits

No complete VBA grammar, class system, scope/type enforcement, arrays,
module lifetimes, Optional/ParamArray arguments or full Variant coercion is provided.
Typed Dim defaults are initialized, but assignment does not yet enforce the
declared type. Date/Currency/Decimal and 64-bit integer fidelity need dedicated
representations. Numbers currently pass through doubles.

Top-level error labels, Resume/retry and nested error-number propagation are tested.
Jump targets inside structured blocks, complete error codes and debugger support
remain unsupported. SQL is still SQLite SQL; no
Jet/ACE compatibility is inferred from opening a recordset. Persistent QueryDef creation, general query parameters/dependencies (see QUERY_ADAPTER.md),
Multiple-database/custom Workspace transactions (see DAO_TRANSACTIONS_RUNTIME.md),
unrestricted join updates and general bookmark/clone semantics,
Filter/Sort, pessimistic locking, concurrency, ADO and default Field-value coercion
without `.Value` remain unsupported. Delete currently refreshes the recordset;
full DAO deleted-record/navigation semantics remain to be implemented.

Forms/Reports collections, full form event ordering, embedded subforms and report
execution are still incomplete. Windows COM/ActiveX/WinAPI are not implemented by
this adapter. Unsupported members throw rather than report success.

## Additional original keypad regression

Eleven original register-form buttons now invoke the unchanged KeyCapt(C) helper
with their original arguments and update the actual Gegeben Qt control. The form
parser recognizes typed binary-property headers (`UnknownProp = {…} Begin`), which
previously closed the Form stack early. All 47 forms now yield 879 controls and
86 command buttons (81 declared visible after the NotDefault correction), rather than the earlier 685 controls/43 buttons.

The syntax inventory checks every artifact hash and classifies all 187 procedures
from 32 VBA modules. Its JSON explicitly marks each entry executed=false and omits
raw diagnostics/literals. Syntax acceptance is never an execution or compatibility
claim. The inventory stays in the ignored build directory and is not packaged.

The For Each regression also executes a counting function against every one of
the 29 original Kassa tables and compares the result with imported schema columns.
This test uses synthetic VBA against real fixture structures, not an original
application procedure. Original For Each procedures still require their specific
Access collections and other unsupported syntax. Native enumeration materializes
an object-reference snapshot; COM enumerator mutation/lifetime behavior is not
claimed. Default members and a VBA Collection class remain open.

## QueryDef parameter collections

`QueryDef.Parameters` returns an actual native collection. Count, zero-based
Item (name or ordinal), bang lookup and For Each use the declared parameter
order. Legacy `Parameters(name)` and collection lookups share cached Parameter
objects; Name/Value access and binding execute against the actual QueryDef.
Missing entries report 3265 and Null indices report 94.

The collection reflects replacement SQL. Parameter handles from an earlier SQL
generation become invalid (3420), including when the new declaration reuses the
same name. Closing the QueryDef invalidates parameter Name/Value operations and
collection access. Tests verify identity, native VBA iteration, real SELECT
results, SQL replacement, stale references, empty collections and close errors.
The six original parameter declarations exercise collection count/order/identity;
five also execute positive filtered recordsets (QUERY_ORIGINS.md).

Parameter Type/Attributes metadata, implicit parameter inference, Refresh,
ADO collections and every default-member/coercion combination remain incomplete.

Cursor/clone/bookmark operations and the original generated form-search path
are implemented in a bounded native subset. Exact scope, typed tokens, validation,
original successes and the missing original control reference are recorded in
[DAO_CURSOR_RUNTIME.md](DAO_CURSOR_RUNTIME.md).

Project-scoped DBEngine/Workspace transactions, typed Workspace locals, native
SQL savepoints and rollback/cursor-generation evidence are documented in
DAO_TRANSACTIONS_RUNTIME.md. CurrentDb.Execute now uses the Database adapter
including positional dbFailOnError; unsupported transaction combinations fail
before pretending to provide complete Jet semantics.

## Native matrix batches

Recordset.GetRows with explicit positive NumRows now returns actual zero-based
field/record Variant matrices, checked typed values and Byte-array BLOB cells.
Normal VBA multidimensional indexing/bounds/ReDim and array-copy semantics apply.
The optional default, nonpositive counts and concurrent/deleted-row partial
batches remain explicit boundaries. Scope and execution evidence:
[DAO_GETROWS_RUNTIME.md](DAO_GETROWS_RUNTIME.md).
