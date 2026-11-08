# Read-only native DAO.Indexes

TableDef.Indexes reads the original, versioned main._la_access_indexes metadata
preserved by the read-only MDB helper. Original names are checked against the
complete actual main SQLite index catalog before returning Count or any item.
Legacy projects without source evidence refuse3251; generated SQLite names are
never returned as invented Access identities. The source metadata/import scope
and new build dependency are in SOURCE_INDEX_METADATA.md.

Supported native types are DAO.Indexes/DAO.Index, including explicit declarations
and TypeOf. The collection supports Integer Count, named/zero-based Item(Index),
Refresh and For Each. Index exposes Name, Boolean Primary/Unique/Required/
IgnoreNulls and Fields. Held items/field collections retain identity. Each index
field exposes its original Name, OrdinalPosition and Long Attributes: dbDescending1
or0. Source field Type/Size/defaults, Index.Foreign/Clustered/DistinctCount and
Append/Delete/CreateIndex/property mutation remain unproven/unsupported. Imported
index and field property writes refuse383.

Every read checks current main table columns/key order, exact catalog shape,
unique source definitions and supported raw flags. Ordinary indexes must have the
actual importer-assigned physical name, complete field order/directions, matching
Unique, BINARY collation and no partial/expression fields. Primary definitions
must match actual primary constraints; a verified single ascending INTEGER
primary may use SQLite's implicit rowid alias. All native indexes must be
accounted for. Additional or changed unknown indexes refuse3251 instead of
returning a partial catalog. Catalog size is bounded4096.

Logical MDB kind2 records with unproven empty fields, IgnoreNulls configurations,
unknown flags, unproven Required constraints and collated/expression/partial
indexes explicitly refuse. This includes source tables with logical relationship
records even when some physical indexes are understood. No Foreign property is
guessed from incomplete logical records. These are bounded native guarantees,
not general Jet collation/NULL/index-update semantics.

Readers retain their TableDef owner, main schema version, workspace lifetime,
rollback revision and original source definition. Observed DDL/drop/recreate,
rollback or replaced valid definitions refuse3420 on held wrappers. Reacquire
TableDef.Indexes after invalidation. Malformed/incomplete fresh metadata refuses
3251. Refresh validates the live catalog; it does not reconstruct an original MDB
or revive invalid readers. TEMP objects cannot shadow the main metadata or index
resolution. Conservative whole-catalog invalidation may be stricter than Access.
Unobserved external generation changes and full Windows lifetimes remain open.

Existing DAO tests cover typed VBA, VarType, named lookup, field directions,
enumeration/identity, readonly/unknown refusals, TEMP shadows, metadata changes,
rollback and unknown native indexes. The technical real Jet reference checks
original named composite/single DESC primary indexes through actual imported
DAO objects. The original Kassa test checks all20 original primary and48 other
index identities/properties/ordered fields after real import. No new CTest target
or full Access2003 oracle claim. Fixed catalog456/384/72; F055/F121/F122/F130 open.

Validation: existing DAO/import/original Kassa3/3 (11.69 s) passed, plus the
new uploaded Kassa copy with all68 original DAO index identities. Current app/
package, app/helper Build-IDs, installed helper load, SQLite3.54.1 and four
original/three upload hashes passed. Last full suite74d75c7:48/48 (85.91 s),
preceding this increment. Public PIA Indexes.Count signature200006 confirms
Integer; field Attributes uses the existing Long interface contract.

## Detached indexes for new native tables

New detached TableDefs support CreateIndex, Index.CreateField, Index.Fields.Append
and TableDef.Indexes.Append before physical TableDefs.Append. Names, field ownership,
field count1..10, final table-column identity, duplicate names/fields and at most
one primary index are checked before DDL. Index.Primary, Unique, Required and
field dbDescending are supported; IgnoreNulls=True and Binary/Memo indexes remain
refused. Primary implies Unique. Primary/Required index columns receive actual
NOT NULL constraints; no nullable primary identity is fabricated.

The table primary constraint, ordinary/unique indexes, managed counters, declared
native field metadata and explicit native index identities share one outer
savepoint. Ordinary physical names follow the existing table_index namespace;
collisions fail the whole table creation. Native declared names are stored in
the same checked versioned format as imported names, with the actual new table
columns; they are not represented as original MDB observations. Empty native
index catalogs are stored explicitly as empty, rather than reported as unknown.

After attachment, held builders resolve through the existing immutable live
catalog. Catalog/schema/generation/rollback/close checks apply. This is new-table
construction only: mutation of already attached/imported indexes, relationships,
Jet locale/collation/IgnoreNulls, original-format writing and complete F121/F130
remain open. Fixed catalog456/384/72 is unchanged.

Validation for detached construction: existing DAO/import2/2 (1.22s), app build
passed. Actual VBA creates primary and descending unique indexes; checked DAO
catalogs reflect physical definitions, and SQL duplicate insertion is rejected.
A deliberate physical namespace collision after table/counter creation proves
complete rollback and corrected retry with first generated ID1. No new target or
broad suite; no Windows oracle or attached-index mutation claim for this stage.

## Index mutation on an existing checked table

Existing TableDefs with a complete, verified native/imported index catalog now
support CreateIndex and Indexes.Append/Delete for ordinary indexes. Detached
builder ownership is tied to the existing table generation/workspace. Final
field identity, supported type, field order/direction, duplicate name/field and
Required semantics are checked again at mutation. A Required index on existing
columns requires their actual NOT NULL declarations; no data-driven guess or
silent table rebuild is used. Adding/deleting primary indexes remains3251.

DDL and the complete updated index definition share one savepoint. Failed unique
creation leaves source rows and the stored catalog unchanged. Delete resolves
actual declared names and removes the matching physical table_index object.
Unique-index deletion on a referenced table is refused pending full dependency
analysis, preserving existing foreign-key validity. Unknown, partial, expression,
collated or incomplete catalogs retain their previous explicit refusals.

Attached builders keep their actual immutable index reader rather than looking
up a replacement definition by name. Old index objects/fields invalidate after
schema changes and rollback; reacquire the table's Indexes collection as before.
This does not implement Access-compatible locks, locale collation, primary-key
rebuilds, IgnoreNulls or full Refresh/lifetime semantics. F121/F130 remain open.

Validation for existing-table mutation: DAO/import2/2 (1.31s), app build passed.
Actual existing data, composite ASC/DESC, duplicate-data unique failure, successful
delete, primary refusal, stale object identity and workspace rollback were checked
in the existing DAO target. No extra target or broad suite.

## Graphic index management

Select a native table in the project explorer and choose **Ansicht →
Tabellenindizes bearbeiten**. The actual modal dialog lists the complete checked
index catalog with field directions, primary and uniqueness information. New
ordinary/unique indexes accept an ordered1..10-field list, ascending/descending
per field and the existing Required constraint option. Normal indexes can be
deleted after explicit confirmation; primary deletion is disabled. Actual DAO
refusals still protect dependent unique indexes and unsupported catalogs.

All changes use the same DAO mutation functions; there is no second schema
writer. Failed operations retain the entered draft and show the real failure.
Dirty table-model drafts must be saved/discarded first. The dialog does not
edit primary keys, rebuild table constraints or pretend unsupported source
indexes are available. Its scope is ordinary index management, not a complete
Access table designer or original-MDB writer.

Graphic-flow validation: existing mainwindow/DAO2/2 (5.11s), app build passed.
The real modal dialog creates a descending index on the UI-authored table,
reloads its actual catalog, deletes that index through confirmation and preserves
both existing rows. No new test target or broad suite.
