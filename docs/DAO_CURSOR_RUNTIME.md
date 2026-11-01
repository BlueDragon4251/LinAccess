# DAO cursors and native form search

Table-backed DAO dynaset/snapshot and read-only query recordsets now support Clone, Bookmark,
Bookmarkable, FindFirst/FindLast/FindNext/FindPrevious and NoMatch. Clones share
materialized data and have independent cursor positions and edit buffers. A new
clone has no current record until navigation, search or a bookmark assignment.
Reading a missing current record fails with 3021. Closing one recordset leaves
its clones usable. Updates, additions and deletions execute actual SQLite writes;
surviving clone cursors remain associated with their row identities.

Bookmarks are opaque byte-array Variants: IsArray=True, VarType=vbArray+vbByte,
TypeName="Byte()". String/numeric coercion fails with 13. Tokens are valid only
within their recordset/clone family; malformed, foreign, deleted and invalidated
tokens fail with 3159. Null fails with 94. A deleted token does not become valid
when SQLite reuses its rowid. Requery replaces bookmarks. General VBA array
syntax, byte-array indexing, BLOB/OLE storage and external bookmark interchange
are not implemented by this opaque-token support.

Find criteria run through the Jet adapter against the recordset's current
materialized rows, with typed SQL bindings. They do not rerun the underlying
query against a different dataset. Unknown fields fail with 3070; pending edits
prevent find/clone/bookmark movement (3219). A failed search sets NoMatch and
moves beyond the searched boundary. Strict identifier preparation prevents an
unknown quoted field silently becoming a SQLite string literal; the connection's
previous quote policy is restored on success and failure. Native SQL execution
and total search time are bounded. This is not complete Jet criteria/function,
collation, locking or multi-user dynaset compatibility. Exact provider-specific
BOF/EOF states and concurrency behavior still need an Access 2003 oracle.

The form Recordset view exposes Clone, RecordCount, Bookmarkable and a Bookmark
getter. Direct mutation/replacement/navigation through that restricted form view
fails explicitly with 3251. Its clones are read-only searchable snapshots.
Form.Bookmark reads the current record token and selects another native record
through ordinary validation/save/Current handling. Cancellation raises 2501.
Successful native table saves refresh shared clone values while preserving row
bookmarks. Explicit form Requery/filter reload invalidates the old form bookmark
family. A closed form view fails with 2450; an independent clone can outlive the
form. This does not supply a complete editable Access Form.Recordset interface.

Str/Str$ supplies the leading blank for non-negative numbers, negative signs,
period decimal separators, omitted leading fractional zero, uppercase exponent
and bounded Currency conversion used in the original generated search code.
Null/type/argument errors are explicit. Exhaustive floating-point formatting,
locale and historical VBA version equivalence remain unverified.

## Original-source regression

Module SHA-256:
`dbd3623b9404e490a814b70de6318d2c0877969b7feef081696d3e235ae3b427`.
The full original layout and unchanged module run against a prepared synthetic
table, without copying original rows. The original Open/Load routines complete.
Two original search procedures each execute positive, no-match and Null-input
cases (six calls); real combo mouse selections execute their original bound
AfterUpdate procedures and select the expected native record. Lookup SQL,
column display and visibility/enabled state are prepared through actual runtime
properties; this is not verification of the original complete lookup dataset.

The third original procedure references a control absent from the exported
layout, even after case/space normalization. Both direct execution and its actual
bound combo event fail with 2465 without moving the record. Its procedure binding
exists; no renamed control alias is guessed. This source inconsistency requires
comparison with the original Access 2003 runtime, not a compatibility-success
label. Full original sales/login/report workflows remain open.

Independent regressions cover cursor isolation, typed/foreign/stale bookmarks,
NULL/LIKE criteria, unknown-field errors, shared writes/deletes/rowid reuse,
ordered query snapshots, native validation cancellation and saved clone values.

Contract reference (current DAO documentation; no installed Access 2003 oracle):
https://learn.microsoft.com/en-us/office/client-developer/access/desktop-database-reference/recordset-clone-method-dao
