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
