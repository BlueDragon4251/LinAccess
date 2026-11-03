# Native DAO GetRows and Variant matrices

Recordset.GetRows with an explicit positive NumRows copies actual materialized
record values into a two-dimensional Variant array. The first index is the
field, the second the record; both lower bounds are zero independently of
Option Base. Requests beyond the remaining rows return the available rows.
The cursor advances to the next unread record, including EOF after the last.
Reading the array does not rerun SQL or modify the source table.

Values use existing checked DAO field conversion: Currency scaled units, Date,
Boolean, text, numbers and Null retain their native Variant types. Native SQLite
BLOB values become Byte-array cells, including embedded zero/high bytes and
empty BLOBs. This does not add MDB/OLE extraction or complete DAO binary Field
properties. Array element replacement changes detached Variant data.

The entire batch is converted before the cursor moves. Unsupported exact integer
ranges and other conversion failures preserve the previous cursor position.
There is an explicit one-million-cell output bound (3251). Pending edits reject
3219; no current record rejects 3021; closed/stale recordsets retain 3420 and Null
counts reject 94. General Access error precedence is not established.

Omitted/Missing and nonpositive NumRows currently reject 3251: the optional DAO
default is not guessed from ADO GetRows behavior. Deleted-row partial batches,
external concurrent cursor changes, provider-specific semantics, full memory/
error behavior and complete DAO compatibility remain open. No Access 2003 result
oracle was available for this segment.

Existing DAO execution tests use a real SQLite table and a complete VBA handler
with named NumRows, Option Base 1, numeric/Null/text/Currency/Date/Boolean/binary
cells, bounds, cursor advance, clipped final batch and EOF. They check that array
writes leave SQL unchanged, errors leave the cursor unchanged, and an unsupported
exact 64-bit value never silently rounds. The core target checks normal matrix
indexing/copies/enumeration and multidimensional ReDim Preserve. No new targets.

Full desktop suite passes 48/48 (92.36 s), rebuilt GUI-free core 1/1 (1.50 s),
Debian package and four original fixture SHA-256 checks pass.

Contract references (current documentation, not an Access 2003 oracle):
[Recordset.GetRows](https://learn.microsoft.com/en-us/office/client-developer/access/desktop-database-reference/recordset-getrows-method-dao),
[ReDim](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/redim-statement).
