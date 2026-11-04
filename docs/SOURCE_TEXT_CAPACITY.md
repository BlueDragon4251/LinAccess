# Imported source text capacity and provenance

Fresh MDB imports now retain the origin of the structural source type and
VARCHAR capacity1..255 in main._la_table_schema. AccessCatalog reads MDBTools'
MySQL schema without executing it and emits encoded structural origin/size
annotations; exact floating/Currency overrides retain their Access-schema origin.
The parser rejects invalid, conflicting, foreign or unpaired size/origin metadata.
No data row or VBA literal is placed into these annotations.

TableDef.Field.Type can now distinguish confirmed MySQL TEXT/Memo declarations
from undifferentiated old SQLite TEXT. Memo has defined Size0. Proven VARCHAR
capacity is exposed as Long Field.Size. Old projects without the required source
provenance/width remain explicitly unsupported3251 where the distinction is
unknown. There is no silent legacy migration or guessed maximum length.

Fresh imported tables enforce confirmed short-text capacities through a native
SQLite CHECK callback. It counts UTF-16 code units with the existing native VBA
string implementation, including surrogate pairs and embedded NUL. NULL remains
subject to the existing separate Required/NOT NULL constraint. Overlong writes
fail with catchable3163 through the SQL/DAO adapter; malformed text/scalar types
retain checked errors. The pure validator is registered deterministic/innocuous.
It reads no database, document, host or external state. Native connections already
install these callbacks. Other SQLite writers must install the app's functions
before writing tables that use them; this is an internal .linaccess format detail.
Existing projects/tables are not rewritten to add constraints.

Existing import cases persist actual source metadata and inspect DAO Type/Size.
They accept exactly20 UTF-16 units and reject21 units with an astral character,
reject an overlong embedded-NUL string that SQLite's ordinary length would miss,
check actual DAO error3163 and verify rejected statements add no rows. The
original MDB test verifies57 text capacities and3 confirmed Memo declarations
on its existing working copy, without altering any original fixture.

Validation: final desktop48/48 (85.63 s), Debian package/build ID/pinned SQLite
and all four original fixture hashes passed. The first callback attempt caught
a missing3163 SQL-error allowlist entry; that propagation is fixed in the final
validated build. Last headless core1/1 (1.66 s) is unchanged by these data-only
changes. No new CTest target or catalog ID;456/384/72 remains fixed.

This is bounded source-metadata retention and native constraint execution.
Complete Jet text encoding/collation, UI editing/event timing, old-project
migration, binary/Decimal/AutoNumber semantics, original-format writes and Windows
Access2003 result/error oracle remain open. F014/F015/F121 are not closed.
