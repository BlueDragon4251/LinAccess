# Native Recordset.CopyQueryDef

A cursor opened through an actual QueryDef retains its original Access SQL at
opening time, separately from compiled SQLite SQL. CopyQueryDef returns a fresh,
unnamed, independently editable QueryDef. Each invocation creates a separate
object; the source definition may subsequently change or close. Clones preserve
this source provenance. No persistent catalog entry is created. Existing database
connection, transaction and cursor-close guards apply.

The implemented scope is parameter-free SELECT QueryDefs. Parameterized sources
refuse with3251 until parameter-copy semantics have an Access oracle. Physical
table cursors, raw Database.OpenRecordset SQL and frozen form adapters have no
proven QueryDef provenance and refuse3251. Arguments refuse450; closed cursors
refuse3420. These bounded errors do not claim exact Windows error identity.
Sort/Filter transformation, forward-only modes, parameter values/properties,
pass-through/provider properties and full DAO/Access2003 parity remain open.

Public API reference:
https://learn.microsoft.com/en-us/office/client-developer/access/desktop-database-reference/recordset-copyquerydef-method-dao
This documents copying the definition, including editing its SQL independently;
it is not a Windows execution oracle for this implementation.

App and existing DAO target built. One bundled dao_execution check passed3.23s:
actual SQL rows/order, source edit/close, independent repeated copies, clone
provenance, cursor-close errors and actual typed VBA CopyQueryDef invocation.
No new test target or catalog ID. F054/F055 remain open;456/386/70 unchanged.
