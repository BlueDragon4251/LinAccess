# Read-only native TableDef.Fields

TableDef.Fields now returns a real DAO.Fields interface backed by current main-
schema metadata. Count, named/zero-based Item, default field lookup, Refresh and
For Each expose the actual visible PRAGMA table_info columns without reading
user rows. Held collection/item instances reuse live native objects. Explicit
DAO.Fields/DAO.Field bindings and TypeOf use real native interface identities.

Fields expose Name and OrdinalPosition. Type uses proven integer, Currency, Date and BLOB declarations and now also
explicit source metadata from main._la_table_schema, as described below. Unknown
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

## Source type declarations and fixed Size

A checked, unique main-schema source metadata array now classifies unambiguous
BOOLEAN/BIT, TINYINT/SMALLINT/INT, CURRENCY, SINGLE/DOUBLE, DATE/DATETIME,
VARCHAR, BLOB, MEMO and GUID declarations into their DAO enum values. Unknown
or ambiguous source kinds (including undifferentiated TEXT/FLOAT) remain3251;
no source distinction is inferred from row values. Invalid/duplicate/foreign
field metadata refuses explicitly. These are read-only declaration identities,
not proof of full underlying Jet storage/coercion or original-format round trips.
Direct main-table Recordset fields now share proven source identity. Query and
frozen projections remain narrower where no source identity is carried.

Field.Size returns a Long for proven fixed types: Byte/Boolean1, Integer2,
Long/Single4, Currency/Double/Date8, GUID16, and LongBinary/Memo0. The last two
are their defined metadata sizes, not the length of a record's data. Legacy text maximum lengths remain3251 when unavailable. Fresh imports now
retain proven capacities and expose them as Size; see SOURCE_TEXT_CAPACITY.md.
Recordset.Field.Size shares the same fixed-size helper for its proven native
field types, including BLOB0 even with an empty recordset. Writes raise383.
Existing FieldSize remains the current stored non-NULL binary observation length.

The existing DAO target checks actual source metadata for Text/Double/Memo/
Boolean and the original native fields, Long size identity, BLOB metadata size0,
unknown text capacity refusal and size reads from an actual empty recordset.
Focused validation1/1 (0.59 s), package/build ID/SQLite/four fixture hashes passed;
latest full baseline immediately before this extension48/48 (84.55 s), not rerun.
Last fresh GUI-free core1/1 (1.66 s) is unchanged. Catalog456/384/72, F121 open.
Reference: [DAO Field.Size](https://learn.microsoft.com/en-us/office/client-developer/access/desktop-database-reference/field-size-property-dao),
[DAO Field.Type](https://learn.microsoft.com/en-us/office/client-developer/access/desktop-database-reference/field-type-property-dao).
No Windows Access2003 execution oracle is claimed.

## Physical Recordset field metadata

Direct main-table recordsets now read the same checked source metadata as
TableDef.Fields and keep immutable scalar DAO type/size maps with the loaded
rowset data. Field.Type/Size therefore expose confirmed source Text/Memo/Double/
Boolean and text capacities without reinterpreting row values. Clone shares the
same data-family metadata. Requery refreshes it, and failed reload restores the
prior maps with the existing rowset rollback guard. No native document, host,
VBA object or live SQL cursor is stored in the maps. Closed/rollback/thread/
connection guards apply before any field metadata read. Temporary-only sources
keep the previous native declaration fallback; main-table TEMP shadow refusal
is explicit. Query/frozen projections without proof keep their previous subset.

This preserves source declaration identity, not complete Variant subtype/storage
semantics (notably Single and all query lineage) or full DAO schema lifetime.
The existing import test checks actual data, source type/size, Clone and closed
field refusal. The original-copy test verifies57 text capacities and3 Memo
identities on real physical recordsets as well as the table catalog.
Final full desktop48/48 (84.98 s), package/build ID/SQLite/four original hashes
passed; fresh headless runtime1/1 (1.66 s) is unchanged. Catalog456/384/72.

Supported native literal Field.DefaultValue is now readable as a String on
checked TableDef fields and direct-table Recordset fields. Canonical expression,
limits and unsupported setters/projections: [default scope](SOURCE_LITERAL_DEFAULTS.md).
This does not close the complete F121 metadata/schema contract.
