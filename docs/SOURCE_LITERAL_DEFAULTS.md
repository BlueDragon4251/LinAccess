# Native literal table defaults

AccessCatalog now explicitly requests MDBTools' default-value export. Fresh
imports compile only finite numeric, Boolean, quoted text and NULL literals into
native SQLite column defaults. TRUE uses Access's signed-1; FALSE uses0. Text
quotes are decoded/re-escaped; arbitrary source SQL/VBA expressions are not
executed. Unsupported exported default expressions reject the unsupported schema
instead of silently dropping its defaults. Source and normalized literal metadata
are retained. AutoNumber always keeps the existing generator, not a reported
source default. Existing rows are imported explicitly and remain unchanged.

The same native literal compiler reads actual main-schema DDL for pending values.
DAO AddNew now exposes these literal defaults before Update; CancelUpdate writes
nothing. The loaded rowset family owns immutable QVariant maps, shared by Clone
and refreshed/restored with the existing Requery/rollback guard. Constants are
read with one scalar SELECT; generators and arbitrary functions never run merely
to inspect metadata. Raw SQLite TRUE has SQLite's value1, whereas fresh imported
Access TRUE has already been normalized in the actual DDL.

Direct-table forms cache actual literal defaults with their loaded source.
New-record bound controls receive them before Form_Current. Explicit control
DefaultValue overrides table defaults; showing defaults leaves Dirty false and
inserts no row. Existing NULL records are not replaced by defaults. Runtime
source reload clears/rebuilds the cache; query/report and non-main sources are
not given invented field defaults. The public scalar reader does not require
full source-type metadata, preserving old counter-only catalogs; explicit invalid
type metadata still refuses when relevant, while missing optional type proof
retains the native declaration fallback.

The original working-copy test verifies96 source declarations,96 actual native
DDL defaults and96 pending DAO values. Existing cases independently verify
numeric/Boolean/text/NULL values, quoted text containing NOT NULL without inventing
a Required constraint, AddNew/CancelUpdate/Update, actual form Form_Current,
control-over-table precedence, no phantom row and the saved SQL values. The
legacy unsupported-counter clipboard case remains unchanged and passes.
No new CTest target or fixed-catalog ID was added.

Validation: full desktop48/48 (85.12 s). A subsequent UI-only import-information
text correction passed its inspector target1/1 (11.04 s); final Debian package,
matching build ID, pinned SQLite and all four original fixture hashes passed.
The last fresh headless runtime1/1 (1.66 s) is unchanged by these Qt data/UI changes.
Catalog456/384/72 remains fixed; full F014/F015/F121 requirements remain open.

Scope is literal defaults, not complete Access default-expression/property/
GenUniqueID/Date/Now, exact Windows event timing, Field.DefaultValue editing,
writable-query defaults, original-format writes or old-project migration. Exported
source expressions use MDBTools' representation, not a Windows DAO string oracle.
Raw existing SQLite dynamic defaults can still run on insert but are not
preinitialized by this literal-only buffer. No Windows Access2003 equivalence
claim follows from these measurements.
