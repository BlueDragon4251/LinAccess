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
