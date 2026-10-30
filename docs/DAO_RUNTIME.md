# Native DAO execution subset

The C++ runtime now carries object references and dispatches members to a
SQLite-backed Database/Recordset/Field implementation. This is actual execution
against the open LinAccess project, not procedure recognition. No original MDB
is written and no additional per-statement confirmation is introduced.

Supported and tested:

- `Set db = CurrentDb()` and `CurrentDb.OpenRecordset(...)`, database references,
  `Nothing`, `Is`, multiple Dim declarations and DAO-qualified variable types.
- Table dynasets (`dbOpenDynaset`) and read-only SELECT/snapshot sources
  (`dbOpenSnapshot`). Results are materialized, limited to 100,000 rows.
- `EOF`, `BOF`, `RecordCount`, `MoveFirst/Last/Next/Previous`, `Requery`, `Close`.
- `rs!Field`, `rs.Field`, `rs.Fields(name/index).Value`, Field.Name, NULL values.
- `Edit`, `AddNew`, `Update`, `CancelUpdate` and `Delete`, using bound SQL values.
  Table updates identify the actual SQLite rowid; insert buffers contain only
  explicitly assigned fields, so available SQLite defaults remain effective.
- `Do While ... Loop` with the existing execution budget.
- `DoCmd.Quit` requests Qt application termination.
- Variant Empty versus Null, Null-aware Boolean logic, numeric bitwise And/Or/Not,
  comparison-before-Not precedence, CLng tie-to-even rounding and Long overflow,
  Null propagation for Len and explicit invalid-Null errors for CStr.
- Procedure-local `On Error Resume Next`, `On Error GoTo 0`, `Err.Number`,
  `Err.Description`, `Err.Clear` and single-argument `Err.Raise(number)`.
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

No complete VBA grammar, class system, scope/type enforcement, arrays, ByRef,
module lifetimes, general parameter passing or full Variant coercion is provided.
Typed Dim defaults are initialized, but assignment does not yet enforce the
declared type. Date/Currency/Decimal and 64-bit integer fidelity need dedicated
representations. Numbers currently pass through doubles.

Label-based error handlers, Resume/retry, exact nested-handler semantics, complete
error codes and debugger support remain missing. SQL is still SQLite SQL; no
Jet/ACE compatibility is inferred from opening a recordset. QueryDefs, parameters,
Workspace transactions, SQL-join updates, bookmarks, recordset clones,
Filter/Sort, pessimistic locking, concurrency, ADO and default Field-value coercion
without `.Value` remain unsupported. Delete currently refreshes the recordset;
full DAO deleted-record/navigation semantics remain to be implemented.

Forms/Reports collections, full form event ordering, embedded subforms and report
execution are still incomplete. Windows COM/ActiveX/WinAPI are not implemented by
this adapter. Unsupported members throw rather than report success.
