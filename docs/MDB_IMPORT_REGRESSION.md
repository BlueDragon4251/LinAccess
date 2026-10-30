# MDB import investigation

The published Debian 13 environment used MDBTools 1.0.1. The original parser
accepted only column lines inside CREATE TABLE. A separate `, PRIMARY KEY (...)`
line made it return an empty schema. This deterministically rejected 20 Kassa
tables and both recipe tables. The remaining nine Kassa tables could be copied;
the old test threshold was only 20/29. This is a confirmed LinAccess parser defect.

The unchanged original code was also compiled separately and run with Debian 12's
signed MDBTools 1.0.0 packages: it copied 29/29 in this environment. That version's
SQLite backend emits neither primary keys nor indices. The historical intermittent
EOF/0–20 results are not reproduced here and their root cause remains unresolved.
Do not infer that the parser fix explains every previous failure.

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
