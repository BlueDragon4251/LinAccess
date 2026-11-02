# Declared integer field subtypes

The existing Access import metadata retains unambiguous MDBTools MySQL source
names: TINYINT (Jet Byte), SMALLINT (Jet Integer), INT (Jet Long). Native table
column metadata now recovers INTEGER1/INTEGER2/INTEGER4 aliases from those names.
Qt controls and direct-table DAO fields return Byte/Integer/Long value subtypes
instead of a Double for every imported integer. Scalar field writes use checked
conversion/range/rounding; failed conversion preserves the previous draft/data.
Malformed fractional/binary integer storage is rejected instead of narrowed.

Native DAO Field.Type now reads dbByte/dbInteger/dbLong (2/3/4) for these known
integer types and returns an Integer value. It remains read-only on these
recordset fields (native 383), diagnoses arguments (450), closed owners (3420)
and unknown source types (3251). Metadata remains readable at EOF. Generic
SQLite INTEGER columns retain the existing exact-through-2^53 Double contract;
the adapter does not guess a Jet width for them.

The native integer_field_documents Qt/VBA/DAO/SQL CTest verifies metadata/type
codes, boundary/default values, fractional programmatic conversion, pending
Text/Value and explicit save, no mutation on range failure, DAO cancellation,
readonly/closed/EOF type access and corrupted legacy storage diagnosis. It also
preserves a generic SQLite integer 9007199254740991 without narrowing to Long.
Targeted native/DAO/original tests pass 3/3 (6.98 s); GUI suite now has 31 CTests.

The strengthened real MDB table-copy regression checks all 127 original INT
columns and 471904 non-NULL native Long field values. Native integer-row/NULL
multisets match an independent SQL projection; the existing SQL-vs-MDB-CSV full
cell comparison remains active. The strengthened test passes (7.58 s). Analysis
used a SHA-verified MDB test copy, and no original column names/values were logged.

This is scalar/direct-table API type recovery, not complete Jet storage/coercion.
Raw SQL constraints/coercion, general query/expression type propagation, full
DAO Field.Type/property/error identity, locale/editor parsing, 64-bit Variants,
Single/Currency/Decimal/Date storage and a Windows Access oracle remain open.
Raw SQLite writes can still violate imported ranges; reads diagnose that invalid
storage. Source data and stored values are not migrated by this change.

MDBTools' MySQL backend maps both MDB_MONEY and MDB_FLOAT to FLOAT. That metadata
is insufficient to distinguish Currency and Single, so no blanket Single recovery
is added. Exact original types need additional metadata before that work.

References:
- https://learn.microsoft.com/en-us/office/client-developer/access/desktop-database-reference/datatypeenum-enumeration-dao
- https://learn.microsoft.com/en-us/office/client-developer/access/desktop-database-reference/field-type-property-dao
- https://github.com/mdbtools/mdbtools/blob/v1.0.1/src/libmdb/backend.c

MDBTools source was inspected outside this repository; no external source or new
runtime dependency is included in the app.
