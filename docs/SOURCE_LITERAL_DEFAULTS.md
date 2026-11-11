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

## Read-only DAO Field.DefaultValue

Direct main-table Recordset fields and checked TableDef fields now expose
DefaultValue as a String for supported native literals. SQL string quoting is
translated into a canonical VBA quoted expression, numeric literals retain their
value expression, and explicit SQL NULL becomes Null. An actually absent default
returns an empty String. Unknown/query/frozen expressions, generators, binary
field defaults, multiline/NUL text and expressions over255 UTF-16 units refuse
3251; no arbitrary default function is evaluated by a property read. Recordset
maps follow the same data-family/Clone/Requery/closed guards. Writes remain383;
TableDef default editing and exact original Windows lexical representation remain
open. This is a canonical native expression, not a claim of original VBA spelling.

Existing import cases check String identity, numbers, quoted text, NULL, the
table catalog and readonly refusal. The original-copy case checks all96 reported
default expressions while independently checking their pending values.
Focused DAO/import/original3/3 (10.48 s), latest Debian/build ID/SQLite/four fixture
hashes passed. Last full baseline48/48 (85.12 s); headless runtime unchanged1/1
(1.66 s). Catalog456/384/72 remains fixed; no extra CTest target.

## Native Date/Now/Time defaults (2026-11-08)

Own native Date fields now accept the bounded zero-argument expressions `Date()`,
`Now()` and `Time()` in DAO and the table/field designers. Case, whitespace and a
leading `=` are accepted. Stored defaults call private native SQLite functions;
reading Field.DefaultValue returns the canonical Access expression without
executing it. No arbitrary SQL/VBA default is evaluated by these readers.

DAO AddNew, manual-submit datasheet inserts and direct-table form additions
initialize these defaults in the new record buffer. The frozen values are
written at save, including Date defaults without a bound form control. Explicit
Null and later field edits take precedence. SQL DEFAULT VALUES evaluates the
actual local system clock at insertion. Date uses midnight; Time uses the OLE
base date 1899-12-30. Storage retains the existing whole-second limitation.
No counter is consumed by reading Date defaults.

New native Date fields require canonical ISO timestamps through an actual CHECK;
literal Date defaults use the same conversion as field writes. Existing native
Date definitions retain their older constraint during field/index/relation/table
migrations; metadata does not silently upgrade their stored data. The supported
clock defaults survive these migrations and table renaming. General expressions,
imported default-expression evaluation, writable-query defaults, regional text
identity, subsecond storage and exact Windows event timing remain open (F015,
F014 and the existing DAO/form requirements). This does not establish full parity.

Gezielte bestehende GUI-/DAO-Ziele2/2 bestanden (6,24s), App gebaut. Keine neue
Vollsuite-/Windows-Referenzbehauptung; Katalog456/384/72 unverändert.

Native Date defaults also accept bounded ISO #date# constants (optional 24-hour
time, whole seconds), with an optional leading=. Native literal defaults in
general accept leading= and outer whitespace; quoted contents remain intact.
The shared Date literal reader is also used by field rules. Invalid dates refuse
before schema mutation; table-dialog errors retain input. This does not change
the MDB import default parser or claim Windows lexical/locale equivalence.
GUI/DAO2/2 (7.98s), app built; no new target or full-suite claim.

Exact non-NULL Decimal numeric literal defaults now use the bounded encoded
path in [EXACT_DECIMAL_DEFAULTS.md](EXACT_DECIMAL_DEFAULTS.md). General
Decimal default expressions and native field property editing remain open.
