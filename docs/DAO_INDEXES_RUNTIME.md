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
