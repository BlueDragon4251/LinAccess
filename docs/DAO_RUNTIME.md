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


Native `Error expression` validates one expression before procedure effects and
uses the same actual Err propagation path and source line as Raise. Its native
instruction does not depend on a local variable named Err. Err exposes
its native ErrObject identity. Existing core cases cover handled signed custom
codes, a parenthesized number and refusal of a comma-list Error statement.
Unary plus/minus resolve supported scalar defaults and propagate Null instead
of converting it into numeric zero; exact wider numeric promotion remains open.
These changes refine existing F051/F041, without adding catalog points.


Native module Enum members now bind as immutable Long constants, with zero-based
automatic numbering, explicit supported integral constant expressions, aliases,
negative values and checked 32-bit overflow. Private/Public visibility is stored
in the actual native module state (omitted visibility currently uses Public).
The same ordered declarations feed compile-time defaults/bounds and runtime
module bindings; malformed/unclosed declarations and duplicate member names
are refused before procedure effects. Existing scalar Const behavior remains.

Unqualified and Enum-qualified constants resolve without a runtime Enum object;
existing local/instance roots retain lookup precedence. Qualified writes are
refused before body effects when they name a constant. Under F042, variables/
arrays/parameters declared As an Enum type, imported type libraries, complete
standard/class-module scope rules and Access validation remain open. Enum types
are not fabricated as runtime objects. No new catalog ID or CTest target.


Native GoSub/Return uses procedure-local identifier labels and a bounded LIFO
return stack (48 frames). It shares actual procedure locals; caller conditional,
For/For Each and With scopes stay alive. Return discards callee scopes and
continues immediately after the latest GoSub. GoTo inside a subroutine discards
only scopes above that caller's floor, retaining actual collection leases and
array-source guards. Return without GoSub raises 3; exhausted return stack
raises 28. Native calls do not enter or leave other procedures through GoSub.

Under F040/F051, label handlers snapshot the return stack with the real scopes.
A handler can make its own GoSub calls; Resume/Resume 0 and Resume Next restore
fault context, and Resume label abandons it. Obsolete snapshots are released.
Direct Return from a handler into the suspended fault GoSub is explicitly refused
with 3251 pending reference semantics; use Resume. On Error Resume Next remains
available. Scoped GoTo and numeric labels use the checked paths described below. Existing core execution covers nested GoSub in a live Collection
loop, inline If, numeric loops, GoTo/Return, retry through a label handler's own
GoSub, held mutation guards and actual lease release afterward.
Reference: https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/gosubreturn-statement


Native scoped GoTo retains active destination scopes and unwinds only exited
If/Select, For/For Each and With state. Labels may be defined inside blocks;
a destination must be in the same block or a lexical ancestor. Entering an
inactive/sibling block is refused before body effects. Subroutine caller floors
remain protected. GoSub/error/Resume targets remain top-level. Inline-wrapper labels are rebuilt from the merged plan, so
partners, duplicate detection and source mapping refer to the real procedure.
The transient scope-parent table uses one integer per instruction; cached plans
retain only existing source-only integer unwind fields and actual label strings.

Existing core execution verifies Continue-style For/For Each jumps, leaving
With/If, inline label binding, leaving enumeration and real mutation afterward.
The existing jump-into-uninitialized-nested-For rejection remains mandatory.
This refines existing F040; no new checklist or CTest entry.


Numeric line labels (0–65529), with or without a colon before a following
statement, use canonical numeric keys and the same scope checks as named labels.
GoTo, GoSub, label handlers and Resume targets accept them; leading zeroes do
not create different labels. On Error GoTo 0 still disables the handler and
Resume 0 still retries the failed instruction. Physical source diagnostics stay
separate from numeric labels; Erl/debugger equivalence is not claimed. The range
is documented in archived Microsoft KB Q111869 (Line Numbers Greater Than 65529
Not Supported): https://www.betaarchive.com/wiki/index.php/Microsoft_KB_Archive/111869

The existing real SQL clone/error-resume case also leaves With through GoTo:
the live clone blocks transactions during With and is released by the checked
branch unwind before the next transaction. Original data is not used.
