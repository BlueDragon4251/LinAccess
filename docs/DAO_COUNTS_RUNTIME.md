# Native DAO counter values and subtypes

Database.RecordsAffected and QueryDef.RecordsAffected now return checked VBA Long
values. Recordset.RecordCount returns a checked Long for the existing loaded
snapshot count. Fields.Count returns checked VBA Integer, as specified by its
public DAO contract. Values outside the respective signed range report overflow
6 instead of silently exposing a generic Double.

CREATE/ALTER/DROP action statements return zero affected records. SQLite retains
its previous sqlite3_changes value across schema statements; that stale value is
not a DDL record count. INSERT/UPDATE/DELETE and existing joined/materialized
rewrites retain their actual affected count. SELECT INTO/make-table still counts
the newly materialized rows and overrides the schema-statement zero.

Existing DAO/main-window tests reproduce stale DDL counts, then check CREATE,
ALTER, DROP, INSERT/UPDATE, make-table counts, Long/Integer numeric subtypes and
actual zero-row DDL feedback in the SQL editor. No new CTest target is added.

This does not complete cursor-dependent/lazy RecordCount semantics, all DAO
metadata subtypes, transaction/error recovery or Windows differential validation.
Public references:

- https://learn.microsoft.com/en-us/office/client-developer/access/desktop-database-reference/database-recordsaffected-property-dao
- https://learn.microsoft.com/en-us/office/client-developer/access/desktop-database-reference/querydef-recordsaffected-property-dao
- https://learn.microsoft.com/en-us/office/client-developer/access/desktop-database-reference/recordset-recordcount-property-dao
- https://learn.microsoft.com/en-us/office/client-developer/access/desktop-database-reference/fields-count-property-dao


## Physical TableDef.RecordCount (2026-11-06)

A live local TableDef now returns checked readonly Long RecordCount from an
actual bounded SELECT COUNT(*) over main.<quoted table>. No rows are loaded into
a Recordset and a same-named TEMP table cannot redirect the count. Empty local
tables return0; current transaction writes/rollback are observed on the actual
connection. Extra arguments raise450 and assignment383; workspace/thread/
closed-table guards apply before SQL. Overflow is6. Linked-table count -1 is not
implemented because this native catalog exposes local physical tables, not a
linked-table adapter. Unobserved drop/recreate identity remains open under F121.

Existing DAO cases cover typed VBA/Long identity, empty/populated counts,
readonly/arity, same-connection insert/rollback and TEMP shadow separation.
The original Kassa DAO case compares TableDef counts for all29 imported physical
tables with independent SQL counts and real snapshot RecordCount. Original
business code and fixture bytes are unchanged. Reference (API contract, not an
Access2003 execution oracle):
https://learn.microsoft.com/en-us/office/client-developer/access/desktop-database-reference/tabledef-recordcount-property-dao
No new test target, catalog ID or checkbox;456/384/72 unchanged.
