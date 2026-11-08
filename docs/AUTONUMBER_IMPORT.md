# AutoNumber generation and original counter state

LinAccess is a general database application. Counter identities and state are
bound to actual tables/columns, not to Kassa names or workflows. The original
Kassa file is one regression corpus; technical Jet references and newly authored
native tables exercise the same implementation.

## Current implementation

New native MDB imports preserve the separate AutoNumber flag and original values.
The readonly helper reads the Jet3/4 table-definition last-allocated Long at
offset16/20, checks source SHA256 before/after, and reports exact counter-column
identities. Restoration uses the verified source state, not MAX(remaining rows).
Empty tables and deleted highest IDs retain their last allocation. Source state,
active values and hashes must agree; an import failure rolls back atomically.
Caller-owned CSV does not inspect its diagnostic source filename or infer seeds.

New source-backed counters use INT storage with a managed DEFAULT, including
single-column primary keys. INT avoids SQLite's INTEGER PRIMARY KEY rowid alias.
`_la_counter32_sequences` stores signed32 state. SQL defaults and DAO reservations
advance2147483647 to -2147483648 and then -2147483647. The21 original counters
(16 single-PK, five without a PK) use this same mechanism. The original primary
keys/index directions and relationships remain independently checked. Counter
values and opaque rowid/bookmark identities are separate.

The native table designer creates the same managed32-bit counter. Newly created
native generators through DAO DDL are also configured. Explicit SQLite-style
INTEGER PRIMARY KEY AUTOINCREMENT tables retain their physical rowid alias;
DAO uses the32-bit state there, but raw SQLite automatic SQL cannot wrap that
alias. An explicit unsupported error prevents an out-of-range stored value.
Existing projects and legacy caller-owned CSV generators are not silently migrated;
their previous64-bit SQL allocation remains a bounded legacy contract.

DAO AddNew reserves each generated value before returning, so the new-record
buffer exposes it immediately. CancelUpdate, Close and a failed Update do not
return the reserved value to the sequence. Reservations made inside a native
Workspace transaction survive Rollback: capture the current generator frontier,
ROLLBACK TO the workspace savepoint, restore consumed values while retaining the
transaction lock, then RELEASE. Business rows roll back; consumed numbers persist
in the same project file. No counter sidecar, service or Java dependency is used.
Schema actions after reserved allocations in that transaction are refused3251
until transaction completion; arbitrary reentrant/external transaction changes
are not claimed to match Jet. Existing new-buffer NULL assignments preserve an
already reserved generated value; this is not proof of every DAO assignment case.

Imported managed counters reject explicit SQL NULL with3162. Source-backed
counter UPDATEs raise3113; DAO existing-record assignment raises3164. Proven
writable direct-table Long generators expose recordset Field.Attributes49 and
DataUpdatable=True. Proven TableDef Long-generator fields expose Attributes17
and DataUpdatable=False (catalog field, not a record buffer); catalog assignment
remains read-only383. Bound form controls reject existing counter assignment2448;
normal form updates omit unchanged generated fields. These properties are not
inferred for unsupported contexts.

Persisted32-bit generator metadata supplies Long typing to native field readers
and UI models. Noncompound query projections with verified main-database physical
column lineage preserve the Long subtype for these generators. Expressions,
compound queries and unproven origins receive no guessed counter type.

DAO DDL wraps the schema action and counter/schema/index metadata cleanup in one
savepoint. DROP removes the managed sequence catalogs and original-field/index
metadata for the removed table. Workspace rollback restores both table and state;
a new import under the same name can start from its source seed. Historical logs
and saved query definitions remain. Managed/imported renames are refused atomically
until default/index/source metadata rebinding is implemented. This is not complete
relationship-constraint lifecycle handling.

## Native COUNTER declarations with seed and increment

Local CREATE TABLE supports COUNTER and COUNTER(seed,increment) in actual column
type positions through the existing SQL token stream. Seeds are signed-Long
integer literals; increments are positive Long integer literals. At most one
COUNTER declaration per table is accepted. The compiler emits INT, NOT NULL and a
trusted managed DEFAULT; text literals, identifier text and table constraints are
not scanned as counter declarations. Unsupported/nonliteral/zero increments are
refused before mutation. General CREATE/ALTER/DDL compatibility is not implied.

The new table, generator and optional `_la_counter32_steps` configuration commit
in the existing DDL savepoint. Seed100/increment7 generates100/107/114 as observed
under Access2003; SQL/DAO use the same step, cancellation/rollback retain consumed
values and both step/frontier survive project reopen. Default increments remain1,
including existing32-bit project state without the optional steps table. DROP
also removes step metadata. This does not decode custom increments from arbitrary
original Jet binaries or provide all ALTER/field-property/reseed operations.

## Actual Microsoft Access 2003 reference

The user ran the technical probes with Access11.0 / DAO3.6. Only technical counter
observations are retained in `tests/access2003_reference/counter-observations.json`,
with collector/archive hashes. No original schema, business rows, query SQL or VBA
is included. The reference shows:

- Deleted highest/all rows still allocate4 after allocating1/2/3.
- AddNew allocates before Update; cancellation, failed Required validation and
  rollback consume2/4/6, followed by3/5/7.
- Recordset Attributes49, DataUpdatable=True; existing field assignment3164,
  SQL update3113, explicit SQL NULL3162.
- COUNTER(100,7) produces100/107/114.
- After explicit2147483646, allocation produces2147483647, -2147483648,
  -2147483647.

The reference is not an implementation verdict. Optional Jackcess-generated actual
Jet4 files independently cover deleted-highest/empty-table stored state. Regression
checks exercise new native tables, source imports, immediate AddNew values,
cancellation/failure/rollback, positive/negative wrap, NULL/update errors, persistent
reopen, typed query projection and atomic lifecycle behavior in existing targets.

## Remaining F171 requirements

Full source binary mode/property decoding, ALTER/reseed/field-property semantics, Random and
replication generators, negative source-state import, complete field properties,
form/event allocation timing, raw SQL statement-failure allocation gaps, external
transaction/crash behavior, concurrent application/locking semantics, general
schema rename and relationship lifecycle remain open. The current source decoder
refuses negative state, UUID generators, multiple Long generators and exposed
NewValues/Seed/Increment settings. Absent DAO/MDBTools properties do not prove every
possible counter mode; do not claim generic custom-step/random source support.

Managed SQL defaults are allocated within the inserting statement; failed raw SQL
can still roll back their allocation. DAO reserves in a preceding statement, so its
failed-Update/cancel/workspace-rollback cases have the implemented consumed-number
behavior. Do not transfer that claim to all SQL/form/clipboard event paths.

The fixed456/384/72 catalog is unchanged. No full AutoNumber or Access parity claim.
