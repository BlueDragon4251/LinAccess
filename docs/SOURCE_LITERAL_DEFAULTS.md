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

## Creating native tables

The existing New Table dialog now accepts optional literal defaults. It validates
all fields before creating actual SQLite DDL and closes only after successful
creation. Validation/database errors keep the dialog and field entries available.
Integer/real defaults must be numeric; Boolean numerics normalize to0/-1; text
uses quoted literals, date literals use the implemented ISO reader, and binary
fields accept only NULL defaults. General expressions and default input beyond
255 characters refuse explicitly. The automatic id field remains the existing
creation-wizard behavior; this is not the complete Access table designer.

The existing main-window case first tries an unsupported expression and verifies
that no table exists and the fields remain, then corrects it and creates actual
integer/text/Boolean defaults. An independent DEFAULT VALUES insert stores7,
the entered text and signed-1. Focused UI1/1 (2.82 s), Debian/build ID/SQLite and
four original fixture hashes passed. Last full data/UI baseline immediately before
this dialog extension48/48 (85.12 s); no repeat full run is claimed for this
localized UI change. Catalog456/384/72 remains fixed.
