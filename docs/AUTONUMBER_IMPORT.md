# Imported AutoNumber metadata and bounded generation

MDBTools' MySQL schema exposes auto_increment modifiers that its SQLite schema
can omit. Encoded structural comments preserve the separate AutoNumber flag;
quoted keyword text is not a counter and unknown/noninteger targets are rejected.
Metadata records auto_number/auto_number_native. Original imported values remain
unchanged. These rules apply to new imports; existing tables are not rebuilt or
assigned guessed flags/seeds.

The original Kassa schema has 21 counters. Sixteen single INTEGER primary keys
use SQLite AUTOINCREMENT. Five without a primary key now use a managed DEFAULT
function and persistent _la_counter_sequences state. Allocation belongs to the
inserting SQL statement and rolls back on failure. Seed triggers advance after
explicit integer inserts/updates. All 21 new-import counters report generation;
this is bounded sequential allocation, not full Jet semantics.

Managed defaults require LinAccess SQL functions on each connection; application/
DAO initialization registers them, also after reopen. Form and DAO inserts omit
unset generated fields, including NULL new-record buffers. PasteAppend uses live
column/default metadata to omit generated values; manual keys are retained.
Direct SQL explicitly supplying NULL to a managed counter currently fails NOT
NULL instead of allocating. Counter values differ from SQLite rowid/bookmarks.

Tests cover explicit/generated/no-reuse/higher-key inserts, malformed metadata,
quoted names, failed-import rollback, missing/exhausted sequences, native form
save/Copy/PasteAppend, a counter-only form, DAO AddNew/NULL buffers and reopen
persistence. Two simultaneous connections allocate 40 distinct values. The real
29-table copy checks all 21 flags/generators and separate savepoint-isolated
insert/delete/insert probes on copied single-PK and non-PK tables. All source
cells remain equal; the original MDB is read-only and hash-checked. The installed
backend self-check executes fixture-free managed-counter allocation, constraint
rollback and deleted-highest-value no-reuse checks.

Open: stored source seeds above remaining data, Random/Replication IDs, exact Jet
signed-Long overflow, gaps after failed/rolled-back allocations, assignment timing
before form save, complete counter read-only/DAO Attributes behavior, composite
edge cases, table rename/drop metadata cleanup, direct-SQL explicit NULL and an
independent Access 2003 runtime oracle. Native generation is no blanket AutoNumber
compatibility claim.


## Actual Access 2003 reference and stored source state (2026-11-08)

The user executed the standalone technical probes under Access11.0 / DAO3.6.
Only the technical counter observations are retained in
`tests/access2003_reference/counter-observations.json`, with collector/archive
hashes. No original schema, business rows, original query SQL or VBA is included.
The reference shows allocation at AddNew before Update, consumed values after
CancelUpdate, failed Required validation and transaction rollback; generated
recordset Fields.Attributes=49, DataUpdatable=True, assignment error3164 and
SQL update error3113. COUNTER(100,7) returns100/107/114. A counter advanced by an
explicit2147483646 append generates2147483647, -2147483648, -2147483647.
These findings are a reference, not a statement that LinAccess implements them all.

The native readonly metadata helper now reads the Jet3/Jet4 table-definition
last-allocated Long (offset16/20 respectively), binds its result to before/after
source SHA256, and returns counter-column identities. New native MDB imports
restore both sqlite_sequence and managed counter state instead of deriving the
last allocation from remaining rows. Empty tables and deleted highest IDs retain
the source state. Negative state, UUID generators, multiple Long generators and
exposed NewValues/Seed/Increment settings are refused by this bounded decoder;
source random/custom-step identification and complete property decoding remain
open. In particular, absent DAO/MDBTools properties alone are not proof of all
possible counter modes. No full increment/mode claim is made.

Restoration checks active imported values against the declared source state and
runs in the import savepoint/whole-database transaction. Hash/identity/state
contradictions roll back; caller-owned CSV mode does not inspect its diagnostic
source filename or pretend to recover a stored source state. Existing projects
are not assigned reconstructed metadata.

Native-source counters also receive an UPDATE guard carrying the observed SQL
error3113. DAO editing an existing generated field raises3164. For the proven
writable direct-table Long-generator context, recordset Field.Attributes and
DataUpdatable expose49/True; unsupported contexts still fail explicitly.

Still open in F171: allocation before AddNew returns, durable consumed numbers
across rollback/failure and concurrent applications, signed-Long wrap, custom
increments, random/replication, full property/lifecycle semantics and schema
rename/drop cleanup. The old SQLite64-bit transactional allocator is not claimed
to match the newly collected Jet observations. The fixed456/384/72 catalog is
unchanged.
