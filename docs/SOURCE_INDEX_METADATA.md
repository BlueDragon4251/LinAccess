# Read-only original index metadata

A separate native linaccess-mdb-metadata process reads actual MDBTools index
structures through its documented public library headers. It opens with
MDB_NOFLAGS (no writable flag), reads only a requested table's columns/indexes,
and never loads/executes VBA. It emits a bounded version1 JSON object with actual
source names, raw kind/flags and ordered field names/directions. The process
inherits the existing MDBTools timeout/output/diagnostic checks. Names must be
bounded valid UTF-8, index names and fields unique, arrays complete and column
ordinals valid. Its generic failure message avoids copying private values/names.

The MDBTools1.0.x mdb_read_indices function returns NULL even on success; the
helper verifies the populated array against num_idxs instead. Library metadata
is kept separate from CLI schema export, whose primary-index names are generated.
The helper is found next to the build binary or in the installed sibling
lib/linaccess directory; it is never substituted from an arbitrary PATH entry.

AccessCatalog appends validated-size JSON as a hex schema metadata marker.
The strict shared decoder checks version, actual ordered columns, names,
integer kind/flags and boolean field directions. Actual physical non-primary
source names replace backend-generated SQL export names on new imports.
Primary names and raw flags remain in main._la_access_indexes within the same
existing import transaction, including empty catalogs and logical kind2 records.
The primary field order must match the real table key. Non-primary declarations
must match the captured source catalog completely before any DDL. Old project
files without this evidence are not assigned invented original index names.

Primary directions are retained in native constraints. Single descending primary
keys use an inline PRIMARY KEY DESC so SQLite cannot silently turn their
constraint into an ascending rowid alias. Descending native AutoNumber primary
keys explicitly refuse because their allocation/alias behavior is unproven.
Non-primary IgnoreNulls is currently unsupported; Required indexes need actual
NOT NULL/primary columns. Other Jet NULL/collation and all original-format write
semantics remain open. Logical kind2 records are preserved without pretending
their empty field arrays are complete DAO.Index definitions.

This adds a build dependency on mdbtools-dev and pkg-config (GLib headers are
transitive), already documented in README/CI. The helper dynamically uses the
platform's MDBTools/GLib libraries; no private source or Java/.NET is shipped.
The package checker requires the helper's exact content/permissions and matching
Build-ID and actually loads its installed binary with --version.

The zero-row technical JetIndexCatalogReference.java independently writes and
reopens an explicitly named composite primary index with mixed directions and
a second ordinary index. LINACCESS_JET_INDEX_CATALOG_REFERENCE enables the
existing import target's real source/import/index_xinfo/persistence/hash proof.
The original Kassa test also checks20 real primary names and48 other index names
across29 tables. This is metadata/import evidence, not a Windows Access2003
execution oracle or a complete mutable DAO.Indexes contract. Fixed catalog
456/384/72; F014/F015/F121/F130 remain open.

Existing GUI/import/original Kassa targets3/3 (15.27 s) passed. Final expanded
technical import target1/1 (0.96 s) also verifies actual single-column DESC
primary index structure, beyond the composite case. No new CTest target.

Indexmetadaten-Grundlage vollständig verifiziert: aktueller Debian13-Desktop48/48 (85,91 s), neue echte Jet-Referenzen einschließlich einzelner und zusammengesetzter DESC-Primärindizes, Originalrezepte, Originalkassenhandler und Originalindexnamen bestanden. Aktuelles Paket/Build-ID für App und Metadatenhelfer, tatsächliches Laden des entpackten Helfers, SQLite3.54.1 sowie vier Originalfixture-/drei Uploadhashes bestanden. Kein neuer Katalogpunkt und keine vollständige DAO.Indexes-/Windows-Parität;456/384/72 unverändert. Lokale Debian12-/Ubuntu24.04-Folgeprüfung der neuen Build-Abhängigkeit folgt separat.
