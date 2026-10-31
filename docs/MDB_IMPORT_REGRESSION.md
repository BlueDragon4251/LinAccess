# MDB import investigation

The published Debian 13 environment used MDBTools 1.0.1. The original parser
accepted only column lines inside CREATE TABLE. A separate `, PRIMARY KEY (...)`
line made it return an empty schema. This deterministically rejected 20 Kassa
tables and both recipe tables. The remaining nine Kassa tables could be copied;
the old test threshold was only 20/29. This is a confirmed LinAccess parser defect.

The pre-fix sources (commit `70c83efd170c6b6f402f15b09aabfd31093c2c00`)
were checked byte-for-byte before an isolated version/library matrix. All runs
used the same SHA-256-verified, readonly MDB test copy. Debian package versions
were MDBTools/libmdb 1.0.0+dfsg-1.1 and 1.0.1-0.1.

| CLI reports | Loaded libmdb | Separate PK clauses | Old importer | Current full schema/data test |
|---|---|---:|---:|---:|
| 1.0.0 | 1.0.0 | 0 | 29/29 | 29/29 passed |
| 1.0.0 | 1.0.1 | 20 | 9/29, failed | 29/29 passed |
| 1.0.1 | 1.0.0 | 0 | 29/29 | 29/29 passed |
| 1.0.1 | 1.0.1 | 20 | 9/29, failed | 29/29 passed |

The same CLI executable hash produces different schema formats when only the
library search path changes. Thus the version string alone does not identify
the schema backend. This reproduces a partial import even with a 1.0.0 version
string and ties it to the 20 rejected PRIMARY KEY clauses. The current parser
and structural metadata fallback pass all four real 29-table schema/data tests.

No beyond-EOF diagnostic occurred in this matrix. The earlier intermittent EOF
reports still require independent failure evidence; their historical cause is
not established. The verified parser/library-format cause does not establish an
explanation for every older failure. Current consistent 1.0.0 also passes repeated
full regressions; its SQLite backend omits keys/indices, recovered from MySQL
structural output.

`scripts/mdbtools-diagnostics.sh` records CLI versions, executable/loaded MDB
library hashes and installed package metadata without opening a database. Linux
CI records these diagnostics. Corpus rows, SQL/VBA literals and raw failure logs
are excluded from committed evidence. See
[the counts-only matrix](regressions/mdbtools-version-matrix.json).

The adapter now parses standalone primary keys and exported indices. Where SQLite
output omits them, structural MySQL-backend key/index metadata is read separately;
none of that source SQL is executed. Composite keys, mandatory columns and index
uniqueness are rebuilt with quoted identifiers. Source-reported column types and
all relation metadata (including ordering, enforcement and cascade flags) are kept
in internal SQLite metadata tables.

Full import reads a SHA-256-verified temporary copy, creates all tables and indices
before inserting records, and commits once. Enforced relations become deferred
SQLite foreign keys, with Access cascade-update/delete flags. Non-enforced Kassa
relations remain non-enforced, as in the source. Existing target tables cause a
failure before data changes. A failed row, constraint, index or commit rolls back
the entire import. MDBTools error diagnostics are checked even after exit code 0.

The strengthened Kassa regression requires 29/29 tables, compares all CSV-exported
values and NULLs as multisets (primary-key rowid order can differ from Jet page
order), checks key positions/indices, full import relationship metadata and
foreign_key_check, and verifies the original file hash. Recipe tests check actual
foreign-key rejection and cascading deletion on a separate SQLite copy. Synthetic
schema tests cover composite keys, quoted identifiers and transactional rollback.

Remaining limits: these comparisons use MDBTools as the extraction oracle, not
Microsoft Access. Exact Jet type precision, Currency/Decimal semantics, original
AutoNumber seeds, defaults, validation, collations, lookups and non-NULL OLE remain
unsupported. Non-enforced relationships do not become new restrictions. Full
Access compatibility and historical EOF instability are not claimed resolved.
