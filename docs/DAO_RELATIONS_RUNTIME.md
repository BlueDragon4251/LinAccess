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

## Enforced native numeric relationships

The subsequent native migration path implements actual foreign-key creation and
removal for ordinary numeric relationships, including update256/delete4096
cascade actions. Supported key identities are matching proven Boolean, Byte,
Integer, Long or Double fields and a complete checked non-null unique parent
key. The child must have complete current DAO_NATIVE declarations. Text/Date/
Currency/Single keys, one-to-one flags and original/unknown child reconstruction
remain explicitly unsupported. DontEnforce2 remains a separate metadata-only
mode; default Attributes0 now requests an actual enforced constraint.

Enforced migration requires an enabled foreign-key connection and no active
transaction. It temporarily controls foreign_keys/legacy_alter_table around its
own savepoint, reconstructs only a strictly matching known native table definition,
and restores the connection settings with verification. Unknown/custom table
constraints are refused before reconstruction rather than dropped. A failure to
restore settings marks the connection for reopening; a committed change is
explicitly reported as stored, and native write scopes refuse that connection.

The reconstruction copies explicit rowids and every column, verifies streamed
row/value digests including Double bits and BLOB bytes, restores the original
index/trigger SQL, revalidates the complete index/type catalogs and executes actual
foreign_key_check. Counter frontiers/increments are kept, and copy does not execute
old insert triggers or allocate replacement counters. Invalid existing references
roll the entire schema/catalog change back. Physical table generations invalidate
old DAO cursors/TableDefs after successful reconstruction. Foreign-key removal
uses the same path and preserves rows rather than executing cascade deletion.

This is bounded native enforcement, not full Jet collation, lock, transaction,
crash-recovery, original-MDB migration or Access error-number parity.

Enforcement validation: final existing DAO1/1 (1.33s), app built. Invalid existing
references roll back; corrected retry creates real update/delete cascades. Rowids,
Unicode text, exact BLOBs, maximum Double, view, audit trigger and index directions
survive reconstruction. Deleted maximum counter114/step7 survives; rejected FK
insert consumes121, next successful128; constraint removal permits next135. A
further referencing child table survives the rebuild without cascade deletion.
Unknown additional CHECK remains enforced after refused migration. Actual VBA
cascade-flag OR and active-transaction/settings guards pass. Initial GUI/import
parts passed before the focused fix (3.99s/0.06s); no fresh combined/full-suite
claim. Four original fixture hashes match. Fixed456/384/72 remains unchanged.

## Graphic relationship management

**Ansicht → Beziehungen bearbeiten** opens the actual project relationship
catalog and a checked parent/child field mapping editor. Integrity is enabled by
default and creates real numeric foreign keys; update/delete cascade choices use
the same DAO migration. Unchecking integrity explicitly selects DontEnforce2 and
disables cascade choices. Removing a relationship requires confirmation and keeps
saved rows. Table documents/editable queries must be closed before opening the
editor. Imported/unproven child reconstruction and nonnumeric enforced keys
retain the explicit backend boundaries described above.

Stored schema changes followed by connection-restoration failure have a distinct
native exception subtype. The dialog reports the already stored result and
requires reopening, rather than classifying it as a repeatable failed creation.
Catalog-refresh failure likewise disables further mutation after reporting the
stored result. No recovery failure or fault-injection success is invented.

Existing mainwindow/DAO2/2 (5.33s), app built: real modal create/remove workflow
on an independent technical database, actual FK/cascade clauses inspected, original
row preserved, removed constraint permits a later orphan insert. No new target or
broad suite;456/384/72 unchanged. This is not a complete Access relationship UI,
source-format writer or demonstrated Jet/Windows parity.

## Explicit native one-to-one relationships

Native dbRelationUnique1 now requires complete checked unique parent and child
keys for matching proven numeric fields. It combines with enforced/cascade modes
or DontEnforce (3). Actual unique indexes enforce the one-child-per-key constraint;
FK mode additionally enforces parent existence. Native catalog reads revalidate
both unique keys. Unique-index deletion on a table participating in a declared
one-to-one relationship is refused until the relationship is removed. No unique
index is fabricated or silently created. Imported source flags retain their prior
readonly source-observation contract.

The relationship dialog exposes explicit1:1 and displays the declared mode. Mode
changes do not silently replace an unsupported field choice with another key;
selection must be made explicitly. Existing mainwindow/DAO2/2 (5.64s), app built:
missing child uniqueness refusal, actual duplicate/orphan failures, dependency
protection, enforced and DontEnforce3 lifecycle, and real modal1:1/cascade creation/
removal passed. No new target/broad suite, exact Jet null/index/flag inference
semantics remain open.456/384/72 unchanged.

## Native Single keys

Enforced and one-to-one relationship validation now accepts matching Single
keys only when both sides have explicit own native properties and the actual
SINGLE_REAL declaration. Ordinary imported/legacy REAL fields labelled Single
are refused here: their declared type alone cannot prove quantized future keys.
Checked parent uniqueness/NOT NULL and known-child reconstruction requirements
remain in force. The graphic relationship editor offers proven native Single
fields and creates real foreign keys/cascades, not metadata-only flags.

Existing GUI/DAO2/2 (6.26s), app built. Actual native Single parent/child keys
retain their binary32 expansions and BLOBs through relation creation. Orphans
fail, values rounding to the same parent key remain duplicate keys, update
cascade stores0.2f in the child, delete cascade removes linked rows while Null
children survive, removal succeeds, and unproven legacy Single parents refuse.
The real modal designer also creates/exercises both cascades on its own technical
database. Its delayed error helper is now limited to its own warning title; no
production change after the successful check. Full Jet collation/key/lifetime
semantics and fixed456/384/72 remain open/unchanged.
