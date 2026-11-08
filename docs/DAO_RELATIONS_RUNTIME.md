# Native DAO.Relations scope

Database.Relations now exposes a read-only native catalog for relationship names
preserved by the real Access importer in main._la_access_relations. It does not
invent names for SQLite foreign keys or infer missing Access definitions.

Supported members are Count (checked VBA Integer), Item(Index), Refresh,
enumeration, and Relation.Name, Table, ForeignTable, Attributes (VBA Long), Fields.
Table is the referenced primary table; ForeignTable is the referencing child.
Each Relation field has Name from the primary table, ForeignName from the child,
and its original ordinal. Asymmetric composite keys retain their order. Fields
and their items retain native DAO.Fields/DAO.Field identity; relation collections
and items support typed DAO.Relations/DAO.Relation declarations and TypeOf.
Named Index, canonical names and held object identity follow the existing DAO
collection adapter. Imported properties are read-only (383); unknown keys give
3265, Null indices94, malformed calls450.

Every read validates the actual main user tables and columns, exact metadata
shape, supported attribute bits, contiguous field ordinals and unique mappings.
For enforced relations it matches the complete ordered SQLite foreign key and
its update/delete CASCADE actions. Every actual native key must have a matching
source definition; an unnamed or otherwise unproven key refuses3251, including
Count, rather than returning an incomplete catalog. Source dbRelationDontEnforce
metadata remains visible without pretending a native constraint exists. TEMP
objects cannot shadow either source metadata or table/field resolution.

Held relations/fields verify schema version, workspace lifetime, rollback revision
and current complete source definition. Observed schema/definition changes refuse
3420. Malformed fresh source metadata refuses3251. Refresh validates the live
catalog; it neither reloads the original MDB nor manufactures lost wrappers.
Relation field Type/Size/defaults are unproven and explicitly unsupported.

Public DAO property documentation establishes the primary/foreign roles:
https://learn.microsoft.com/en-us/office/client-developer/access/desktop-database-reference/relation-table-property-dao
https://learn.microsoft.com/en-us/office/client-developer/access/desktop-database-reference/relation-foreigntable-property-dao
https://learn.microsoft.com/en-us/office/client-developer/access/desktop-database-reference/field-foreignname-property-dao
The public PIA Count signature evidence is recorded in DAO_COUNTS_RUNTIME.md.
These are interface references, not execution against Access2003 on Windows.

No Relation creation, Append/Delete, property mutation, Indexes catalog, linked
metadata, external generation tracking or complete Jet locking/refresh behavior
is claimed. Unobserved drop/recreate or metadata change-and-revert remains open.
F055/F121/F122 remain open; the fixed catalog is456/384/72.

Verification: the existing dao_execution and complex_real_mdb_table_copy targets
pass2/2 (11.87 s). The new uploaded Kassa copy independently passes the same
full import/cell/counter/relation regression: eight original relation definitions
and eight field pairs, without changing source bytes. Technical composite-key
cases exercise real SQLite foreign keys, asymmetric names, typed VBA/named
arguments/enumeration, held identity, read-only refusals, metadata changes,
rollback, TEMP immunity, unknown native keys and malformed ordinals. No new
CTest target or complete Access2003 parity claim.

The immutable6cd40d1 binary-import follow-up also completed Debian12 and
Ubuntu24.04: four existing targeted checks each (18.40 s /17.90 s), package
installation/backend/fixture hashes passed. Those snapshots predate both the
Count subtype increment and this Relations implementation.

Aktuelle App/Debian-Paketprüfung einschließlich Build-ID/SQLite3.54.1 und vier Originalfixture-Hashes bestanden. Letzte Vollsuite bleibt6cd40d1:48/48 (91,47 s); kein neuer Vollsuite-/Distributionsnachweis für Relations.

Indexrichtungsnachweis: vollständiger aktueller Debian13-Desktop48/48 (89,51 s) bestanden, einschließlich tatsächlicher Jet-Binär-/Indexreferenzen, Originalrezepte für Datenblatt-PDF, Originalkassenhandler und aller originalen Kassa-Indexfeldreihenfolgen/-richtungen. Schließt DAO.Relations und Integer-Count ein. Keine neue Distributionsmatrix/GitHub-CI für diese Inkremente;456/384/72 unverändert.

## Detached native relation authoring

Database.CreateRelation, Relation.CreateField, Field.ForeignName, Fields.Append
and Relations.Append/Delete now support explicitly non-enforced native
relationships (dbRelationDontEnforce2). Names/table/column ownership, compatible
proven types/Text capacities, duplicate fields/mappings and field bounds are
checked before an atomic catalog mutation. Native definitions carry DAO_NATIVE
provenance, not fabricated MDB source identity. No foreign-key constraint is
created for this explicitly selected non-enforcing mode. Default/enforcing or
cascade creation/deletion remains3251 until actual constraint migration exists.

Own relation generations prevent old Relation/Field objects from reviving after
same-name/same-definition delete/recreate, even without schema DDL. Workspace
rollback invalidates old wrappers and restores the actual catalog. Table/field
schema deletion/rename refuses declared relationship dependencies until those
relationships are explicitly removed; original files are untouched.

Existing DAO/import2/2 (1.33s), app built: real typed/named VBA relation creation,
persisted mapping observed by an independent file connection, explicit orphan
acceptance and empty physical FK list for DontEnforce, unsupported enforcement
refusal, invalid-field atomicity, deletion, same-name identity and rollback.
Dependency refusals preserve source rows, and removal permits later schema
changes. No new test target/broad suite;456/384/72 unchanged, F121 still open.
