# Read-only native TableDef.Fields

TableDef.Fields now returns a real DAO.Fields interface backed by current main-
schema metadata. Count, named/zero-based Item, default field lookup, Refresh and
For Each expose the actual visible PRAGMA table_info columns without reading
user rows. Held collection/item instances reuse live native objects. Explicit
DAO.Fields/DAO.Field bindings and TypeOf use real native interface identities.

Fields expose Name and OrdinalPosition. Type uses the same proven integer,
Currency, Date and BLOB mappings as the existing Recordset.Field adapter; unknown
source types explicitly raise3251. TableDef field Value and other unimplemented
properties are3251; schema writes383. This is no full DAO.Field/property/size/
constraint/Append/Delete contract. Existing immutable import metadata are read,
not rewritten. Collections store metadata readers, not rowsets or UI hosts.

Main schema version and native workspace rollback revision guard each held
metadata reader. Any main-schema DDL conservatively invalidates held metadata
with3420; reacquire TableDef.Fields from the still-valid table to refresh. This
is stricter than a claim of full Access metadata lifetime. Refresh does not revive
an invalidated reader. Workspace close/connection replacement and wrong-thread
access retain native guards. TableDef's existing validity check now also verifies
the original workspace connection. TEMP table/view shadows of the table or import
schema catalog are explicitly3251, so unqualified legacy type helpers cannot
silently consume shadow metadata. External version manipulation/unobserved external
rollback, all drop/recreate identities and full Access refresh semantics remain open.

Existing DAO execution cases run typed VBA over actual schema columns, named
Item lookup, iteration, field type and ordinals, then independently ALTER TABLE,
check stale held collection/field refusal, reacquire the four-field catalog,
exercise a TEMP view shadow and verify close invalidation. No new CTest target.

Validation: focused DAO/layout2/2 (11.08 s) before the final shadow-view guard;
final full desktop48/48 (84.55 s) including that guard, package/build ID/SQLite
and all four original fixture hashes passed. Last fresh headless core1/1 (1.66 s)
is unchanged by this DAO-only change. Catalog456/384/72; F121 remains open.
