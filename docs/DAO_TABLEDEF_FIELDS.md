# Read-only native TableDef.Fields

TableDef.Fields now returns a real DAO.Fields interface backed by current main-
schema metadata. Count, named/zero-based Item, default field lookup, Refresh and
For Each expose the actual visible PRAGMA table_info columns without reading
user rows. Held collection/item instances reuse live native objects. Explicit
DAO.Fields/DAO.Field bindings and TypeOf use real native interface identities.

Fields expose Name and OrdinalPosition. Type uses proven integer, Currency, Date and BLOB declarations and now also
explicit source metadata from main._la_table_schema, as described below. Unknown
source types explicitly raise3251. TableDef field Value and other unimplemented
properties are3251; schema writes383. This is no full DAO.Field/property/size/
constraint/Append/Delete contract. Existing immutable import metadata are read,
not rewritten. Collections store metadata readers, not rowsets or UI hosts.

Main schema version and native workspace rollback revision guard each held
metadata reader. Any main-schema DDL conservatively invalidates held metadata
with3420; reacquire TableDef.Fields from the still-valid table to refresh. This
is stricter than a claim of full Access metadata lifetime. Refresh does not revive
an invalidated reader. Workspace close/connection replacement and wrong-thread
access retain native guards. TableDef's existing validity check now also verifies
the original workspace connection. TEMP table/view shadows of the table or import
schema catalog are explicitly3251, so unqualified legacy type helpers cannot
silently consume shadow metadata. External version manipulation/unobserved external
rollback, all drop/recreate identities and full Access refresh semantics remain open.

Existing DAO execution cases run typed VBA over actual schema columns, named
Item lookup, iteration, field type and ordinals, then independently ALTER TABLE,
check stale held collection/field refusal, reacquire the four-field catalog,
exercise a TEMP view shadow and verify close invalidation. No new CTest target.

Validation: focused DAO/layout2/2 (11.08 s) before the final shadow-view guard;
final full desktop48/48 (84.55 s) including that guard, package/build ID/SQLite
and all four original fixture hashes passed. Last fresh headless core1/1 (1.66 s)
is unchanged by this DAO-only change. Catalog456/384/72; F121 remains open.

## Source type declarations and fixed Size

A checked, unique main-schema source metadata array now classifies unambiguous
BOOLEAN/BIT, TINYINT/SMALLINT/INT, CURRENCY, SINGLE/DOUBLE, DATE/DATETIME,
VARCHAR, BLOB, MEMO and GUID declarations into their DAO enum values. Unknown
or ambiguous source kinds (including undifferentiated TEXT/FLOAT) remain3251;
no source distinction is inferred from row values. Invalid/duplicate/foreign
field metadata refuses explicitly. These are read-only declaration identities,
not proof of full underlying Jet storage/coercion or original-format round trips.
Direct main-table Recordset fields now share proven source identity. Query and
frozen projections remain narrower where no source identity is carried.

Field.Size returns a Long for proven fixed types: Byte/Boolean1, Integer2,
Long/Single4, Currency/Double/Date8, GUID16, and LongBinary/Memo0. The last two
are their defined metadata sizes, not the length of a record's data. Legacy text maximum lengths remain3251 when unavailable. Fresh imports now
retain proven capacities and expose them as Size; see SOURCE_TEXT_CAPACITY.md.
Recordset.Field.Size shares the same fixed-size helper for its proven native
field types, including BLOB0 even with an empty recordset. Writes raise383.
Existing FieldSize remains the current stored non-NULL binary observation length.

The existing DAO target checks actual source metadata for Text/Double/Memo/
Boolean and the original native fields, Long size identity, BLOB metadata size0,
unknown text capacity refusal and size reads from an actual empty recordset.
Focused validation1/1 (0.59 s), package/build ID/SQLite/four fixture hashes passed;
latest full baseline immediately before this extension48/48 (84.55 s), not rerun.
Last fresh GUI-free core1/1 (1.66 s) is unchanged. Catalog456/384/72, F121 open.
Reference: [DAO Field.Size](https://learn.microsoft.com/en-us/office/client-developer/access/desktop-database-reference/field-size-property-dao),
[DAO Field.Type](https://learn.microsoft.com/en-us/office/client-developer/access/desktop-database-reference/field-type-property-dao).
No Windows Access2003 execution oracle is claimed.

## Physical Recordset field metadata

Direct main-table recordsets now read the same checked source metadata as
TableDef.Fields and keep immutable scalar DAO type/size maps with the loaded
rowset data. Field.Type/Size therefore expose confirmed source Text/Memo/Double/
Boolean and text capacities without reinterpreting row values. Clone shares the
same data-family metadata. Requery refreshes it, and failed reload restores the
prior maps with the existing rowset rollback guard. No native document, host,
VBA object or live SQL cursor is stored in the maps. Closed/rollback/thread/
connection guards apply before any field metadata read. Temporary-only sources
keep the previous native declaration fallback; main-table TEMP shadow refusal
is explicit. Query/frozen projections without proof keep their previous subset.

This preserves source declaration identity, not complete Variant subtype/storage
semantics (notably Single and all query lineage) or full DAO schema lifetime.
The existing import test checks actual data, source type/size, Clone and closed
field refusal. The original-copy test verifies57 text capacities and3 Memo
identities on real physical recordsets as well as the table catalog.
Final full desktop48/48 (84.98 s), package/build ID/SQLite/four original hashes
passed; fresh headless runtime1/1 (1.66 s) is unchanged. Catalog456/384/72.

Supported native literal Field.DefaultValue is now readable as a String on
checked TableDef fields and direct-table Recordset fields. Canonical expression,
limits and unsupported setters/projections: [default scope](SOURCE_LITERAL_DEFAULTS.md).
This does not close the complete F121 metadata/schema contract.

## Native detached table creation

`Database.CreateTableDef` and `TableDef.CreateField` now construct detached native
DAO objects. `Fields.Append` assembles the schema; only `TableDefs.Append` creates
the physical table. Field names, duplicate names (including subsequent renames),
object ownership and reserved table names are checked before creation. A table
without fields is refused. `TableDefs.Delete` uses the existing guarded DDL path.

Supported creation types are Boolean, Byte, Integer, Long, Single, Double, Date, bounded
Text, LongBinary and Memo. Text requires an explicit size1..255 and gets the shared
Jet text-capacity check. Detached Required and supported literal DefaultValue
setters affect actual SQL constraints/defaults. Long plus dbAutoIncrField uses the
managed signed32 generator. Native field-attribute constants are available in VBA.
Currency, Decimal, GUID, linked definitions and nonzero table attributes
remain explicitly unsupported here. This is not an original MDB format writer.

Physical DDL, counter initialization and explicitly declared native type/size
metadata are protected by a common savepoint. Native metadata use DAO_NATIVE
provenance, never fabricated MDB source identity. Once attached, held builders
read through the existing live table/field catalog and retain its close,
generation, schema-change and rollback guards. Attached field mutation remains
refused; complete post-creation Fields/Indexes/Relations mutation is still open.

This extends F055/F121 within the fixed456/384/72 catalog. Full DAO schema parity
and a Windows Access2003 execution oracle for this new creation flow remain open.

## Graphic designer uses the same schema construction

The native New Table dialog now builds its schema through the DAO builders above
instead of a separate raw SQLite CREATE TABLE. Automatic id/PrimaryKey, field
constraints/defaults, counter state and native field/index metadata therefore
share the same checked transaction. DAO immediately sees the real declared
PrimaryKey and can add ordinary indexes to the resulting table.

New TEXT fields have explicit capacity255; MEMO provides long text. BYTE,
SMALLINT and INTEGER map to DAO Byte/Integer/Long (8/16/32-bit), REAL to Double,
DATE to Date, BOOLEAN to Boolean and BLOB to LongBinary. Integer defaults are
checked against the chosen range before creation. Existing tables are unchanged;
this is a new-schema behavior, not a migration of legacy SQLite integers/text.
The GUI displays the TEXT/MEMO distinction. General expressions, custom Text
capacities and full Access table-design property editing remain open.

Validation of designer integration: existing mainwindow/DAO2/2 (4.95s), app
build passed. Actual modal UI creates the table, existing defaults/pending counter
checks pass, DAO observes Text255/Long/PrimaryKey, and a subsequent ordinary index
is created on that exact UI-authored table. No new target/broad suite.

## Native primitive storage and declared properties

New native schema fields retain explicitly declared Required and intrinsic
Attributes metadata. TableDef.Fields exposes their Boolean Required, Long
Attributes and False DataUpdatable. Required is read from the native declaration;
it is not guessed from an imported NOT NULL/primary constraint. A declared True
requires an actual NOT NULL constraint. Intrinsic fixed/variable flags are
normalized and conflicting attribute/type combinations are refused. Native
AutoNumber attributes additionally require a proven managed generator. Older
schemas without these explicit properties remain unsupported for those reads.

New Byte/Integer/Long and Boolean fields receive actual typed storage checks,
including the existing conversion UDF's overflow transport. Out-of-range values
cannot bypass the limits through INSERT SELECT. Double fields reject nonnumeric
or nonfinite storage; Date fields validate textual Date storage using the shared
codec. Memo and LongBinary reject mismatched physical storage kinds. Text keeps
its shared UTF16 capacity check. Detached table fields are limited to255.

Supported DAO and checked SQL writes keep their existing coercion path. These
constraints protect declared native storage; raw/complex SQL coercion differences
can still be refused, and no complete Access conversion/error/Boolean-null or
original MDB property parity is claimed. No existing tables are migrated.

Validation: existing mainwindow/DAO2/2 (5.00s) for storage bounds and UI; subsequent
DAO1/1 (1.12s) for exact native Required declaration provenance; final combined
primitive boundary/storage case DAO1/1 (1.21s), app built. Byte255/Integer-32768/
Long2147483647 and rejected overflows through INSERT SELECT; actual maximum
Double, valid Date, Unicode/newline Memo and exact binary bytes; rejected wrong
storage kinds leave the two valid rows intact. Four fixture and three uploaded
original hashes match. No new target/broad suite/Windows execution claim.

The graphic New Table dialog now exposes a Text size column1..255 (initial255).
Non-text fields keep their fixed/native sizes. Invalid lengths and literal text
defaults exceeding the chosen length are refused before schema mutation, retaining
the complete dialog draft. Actual GUI invalid/corrected-save and native Size12
readback passed in the existing mainwindow target1/1 (3.91s); app built. No broad
suite or new target. Existing table sizes remain unchanged.

## Ordinary fields on existing tables

Existing TableDefs now support CreateField and Fields.Append/Delete for the
supported ordinary native field types. The builder is bound to the original
table/workspace generation. Append checks ownership, final names, maximum255
fields and the complete native definition; adding AutoNumber remains refused.
SQLite performs the actual ALTER ADD inside a shared savepoint, including real
validation of existing rows against defaults, Required and storage constraints.

Field metadata are extended/removed without fabricating unknown original types.
When an explicit checked index catalog exists, its complete table-column list
is updated in the same transaction and revalidated. Existing index identities
and directions are preserved. Delete refuses the last field, counters, primary
fields, indexed fields, unresolved expression/partial-index dependencies and
foreign-key fields. SQLite additionally validates remaining schema dependencies.
Dependent saved artifacts are not rewritten or promised repaired.

Attached field builders keep their actual immutable schema reader, so later DDL
or rollback does not rebind them to a new same-named field. Old field collections
retain the existing conservative3420/reacquire behavior. Full Access Refresh,
locking, linked-field mutation, property alteration, source-format writing and
counter/primary reconstruction remain open. This does not close F121.

Validation: existing mainwindow/DAO2/2 (5.38s), app built. Actual VBA with named
CreateField arguments adds Text8/default to an existing row; physical/default
values and metadata are checked. Indexed deletion refusal, safe subsequent delete,
Required-without-default failure on existing rows, counter-add refusal, last-field
protection, old object invalidation and workspace schema rollback all pass in the
existing target. No new test target/broad suite;456/384/72 unchanged.

## Graphic editing of ordinary fields

Select a closed table in the explorer and choose **Ansicht → Tabellenfelder
bearbeiten**. The real field list shows available declared types/sizes/defaults
and Required properties; unproven original properties are shown as unavailable,
not inferred. The dialog adds supported ordinary fields with Text1..255, Required
and literal defaults through the same DAO builder. A normal field can be deleted
with explicit confirmation, including a warning about saved-artifact references.
Primary/counter/last-field deletion is disabled or refused by actual dependencies.
Open table models must be closed before structure editing.

Successful mutations are reported as stored before refreshing the catalog. A
subsequent refresh failure disables further mutations and reports the already
stored result, rather than offering a blind repeated operation. The ordinary
index dialog uses the same separation. No property alteration, primary/counter
rebuild, automatic artifact repair or complete Access designer parity is claimed.

Own Fields.Append/Delete operations advance a connection-owned table-field
revision. Existing physical DAO recordsets, fields and clones refuse3420 until
explicit Requery; they cannot use old cached column layouts for updates. A
successful Requery reloads metadata/data, retires stale bookmarks and clears
invalid edit/chunk buffers across the clone family. TableDef generation identity
is preserved. Raw/external schema changes and full Jet locking remain open.

Graphic editing/cursor validation: existing mainwindow/DAO2/2 (5.11s), app built.
Actual GUI adds Text8/default to two existing rows, verifies both values, deletes
the field after confirmation and preserves the remaining values/row count. An
open DAO cursor/Field/edited Clone refuses3420 after field addition; explicit
Requery reloads three fields and clears the stale clone edit without persisting
its pending text. No new target or broad suite.456/384/72 unchanged.

Schema result readers are now bound and validated before releasing the write
savepoint. Only successful release publishes the TableDef/Field/Index bindings;
no post-release database read can turn their binding into a falsely reported
mutation failure. Existing DAO target1/1 (1.25s), app built; previous GUI/DAO2/2
(5.11s) remains the latest GUI check. No broad suite or concurrency-oracle claim.

## Refresh of held TableDef collections

Held TableDef.Fields and TableDef.Indexes collections now support explicit
Refresh after own schema DDL or workspace rollback. A fresh checked reader must
succeed before replacing the collection's reader/cache. Fields.Refresh advances
its item epoch, so old field items from that collection cannot revive. Index
items retain immutable prior readers/definitions; refresh clears item reuse.
A missing/replaced table or closed workspace remains3420 even if its old name
is recreated. QueryDef/Index/Relation field collections keep their existing
narrow readonly Refresh contract. Parent reacquisition can return a new collection;
exact COM collection identity/Refresh timing is not claimed.

Existing DAO1/1 (1.23s), app built: held collections refresh after field/index
append/delete and rollback; old items remain3420, and equal-name table replacement
cannot revive old collections. No new target/broad suite;456/384/72 unchanged.

## Editing existing native field properties

TableDef.Field.Required, DefaultValue and Text.Size can now be changed for
explicitly declared native fields. The graphic field editor saves these three
properties together through one atomic schema rebuild. Same-type edits preserve
rowids, every stored value (verified by a streamed typed digest), indexes,
triggers, views, supported foreign keys and counter allocation state. Defaults
apply to future inserts; existing Nulls are never filled to satisfy Required.
Invalid shrink/Required/default changes roll back both schema and metadata.
Signed numeric defaults tolerate the SQL compiler's separated unary sign.

Existing cursors/clones refuse3420 until Requery; the edited DAO Field publishes
a checked fresh reader only after commit. Original/imported/mixed schemas,
Counter properties, unsupported types/nonliteral defaults and unknown custom
constraints remain refused. Relation key capacity changes or relaxing Required
require removing the dependent relation first. Unlabelled foreign keys on
unrelated tables do not block an otherwise proven independent edit; unknown
keys involving the edited table remain refused. Original foreign_keys (0 or1)
and legacy_alter_table settings are restored and verified. Failed restoration
after commit reports a stored change and requires reopening.

Existing GUI/DAO2/2 passed (5.92s), application built. Actual GUI edits an
appended Text field, preserves both old values and reads back its new properties;
DAO checks rejected truncation/Null constraints, atomic batch rollback, defaults,
Byte/Long bounds, Date/Boolean normalization, binary/max-Double preservation,
cursor/clone invalidation, relation dependencies, reentrant identity and
transaction restrictions. No new target or broad suite. F015/F121 remain open;
456/384/72 unchanged.

## Rename existing native fields

DAO Field.Name and the graphic field editor now rename a declared native
non-counter column in one savepoint. Field/index catalog names and native
relationship column pairs follow the physical ALTER TABLE RENAME COLUMN.
SQLite rewrites actual indexes, views, triggers and foreign-key references;
saved Access query artifacts, VBA and form bindings are not refactored.
All rowids and typed values are digest-checked. Duplicate/reserved/invalid
names, original/mixed declarations, Counter renames and unknown dependent
foreign-key identities are refused. The edited field publishes its fresh name
and reader after commit; other metadata readers become stale, and recordsets
need Requery. No original-format write or full Access Name AutoCorrect claim.

Existing GUI/DAO2/2 (5.66s), application built. Actual GUI rename retains two
values; DAO verifies Unicode/case changes, index direction/field mapping, real
parent and child FK renames, relation catalog pairs, rejection of an orphan,
view/trigger behavior, cursor refresh, binary/max-Double values and next counter
allocation. Two initial test sequencing errors were corrected (stale field
reacquisition and a delayed popup helper); no new target or broad suite.
F015/F121/F155 and fixed456/384/72 remain unchanged.

## Native zero-length text validation

Text/Memo fields expose AllowZeroLength on detached/new native TableDefs and
explicitly declared existing native fields. False emits an actual deterministic
SQLite CHECK with the shared innocuous native text validator; empty strings
raise3315, Null remains governed separately by Required, and a space is not an
empty string. Both table creation and field creation/property editors expose
the setting. Existing property changes use the same atomic digest-checked
rebuild; offending stored empty strings prevent the change without data filling
or loss. Original/imported/unproven declarations stay unavailable. Older native
metadata without this explicit property returns3251 until a checked explicit
edit records it; the prior native definition must still reconstruct exactly.

Physical DAO recordset/clone fields now also read checked native Required and
AllowZeroLength snapshots. They are readonly383 there; changes belong to
TableDef fields. Requery refreshes the maps with existing metadata/clone guards;
queries/frozen projections or original fields without proof remain3251.

Existing GUI/DAO2/2 (5.99s), application built. Text/Memo creation and change,
empty rejection, Null/space distinction, atomic failed multirow update and
failed constraint tightening, default insertion, AddNew correction after3315,
clone/readonly metadata, and actual graphic create/edit/rename constraints
passed. The GUI expectation was updated for the real sixth designer column.
No extra test target/broad suite. Full field validation/Access oracle and
F015/F121 remain open;456/384/72 unchanged.

## Physical DAO row identity with user rowid fields

Physical recordsets now select an actually unshadowed SQLite identity alias
(rowid, _rowid_ or oid) from table_xinfo, prove it by an unquoted LIMIT0 probe
and retain it in their shared data family. Read/update/RETURNING/delete use
that identity; user fields with those names remain ordinary data. Requery
refreshes the identity with metadata, including clones. All three names occupied
or a WITHOUT ROWID physical source refuse3251 before an ambiguous recordset is
created; readonly SQL projections remain a separate supported path.

The pinned SQLite3.54.1 deliberately forbids ALTER ADD/DROP/RENAME of its rowid
alias names, even where CREATE TABLE accepted ordinary columns with those names.
The native backend now reports this unsupported reconstruction boundary3251
instead of generic3146. Existing native schema reconstruction still refuses
shadowed row identities; no complete arbitrary-schema support is claimed.

Final existing DAO1/1 (1.58s), app built. Real duplicate user-rowid values with
one and two shadowed aliases update/delete only the intended row; native
AddNew/counter identity differs correctly from a reused physical rowid, user
rowid editing is ordinary data, clone/Requery survive other field additions,
and triple shadows/WITHOUT ROWID/forbidden alias ALTER refuse safely. Temporary
debug instrumentation stayed outside Git and was used only on these in-memory
technical cases. No broad rerun; latest GUI/DAO before this change2/2 (5.73s).
F054/F121 and fixed456/384/72 remain open/unchanged.

## Rename own native tables

Existing TableDef.Name (including an attached builder) and View → Rename Table
now perform an atomic native rename. A checked rebuild first rebinds counter
defaults to the destination; source metadata, counter frontier/step rows, index
catalog/physical index prefixes, native relationship names and stored table
page settings follow. SQLite propagates physical FK/view/trigger table
references. Every rowid/value is digest-checked; deleted highest counter values
and custom positive increments are preserved. Readers/recordsets for the former
physical identity cannot revive under its old name. The assigning TableDef
publishes its new name/generation without postcommit database reads.

Original/mixed/unknown schemas or original relation identities remain protected.
Case-only renames, shadowed row identities and named-index SQL hints in physical
views/triggers are refused pending dedicated reconstruction. Saved Access query
artifacts/VBA/form bindings are not refactored; the dialog says to update them.
Matching native table/editable-query documents must close first. This is no
full Access Name AutoCorrect, linked-table or original-format writer claim.

Existing GUI/DAO2/2 (5.82s), app built. DAO verifies actual rename/old-reader
invalidation, a late physical index-name collision rolling back the entire
operation, comment-separated index-hint refusal while a string literal is
allowed, source BLOB/max-Double/Unicode values, parent FK/name/cascade updates,
view/trigger preservation and next counter114 after deleted107 with step7. GUI
renames a real authored table, retains page settings and inserts the next
counter value. The page-store header now includes its required geometry types
explicitly. Uploaded original3/3 and repository fixture4/4 SHA256 checks passed;
no private sources changed. Fixed456/384/72 and F015/F055/F121/F155 stay open.

## Connection write capability and local metadata

Database.Updatable, physical TableDef.Updatable and Recordset.Updatable now
reflect SQLite main-database readonly mode and the actual query_only connection
flag. Opening a physical recordset on a readonly connection retains its readonly
mode. Write attempts through Execute, schema changes, Edit/AddNew/Update/Delete
and field/chunk setters refuse3027 before counter allocation or data mutation.
Actual SQLite READONLY errors in direct catalog persistence also map3027 by
native error code. A pending edit cannot save while query_only is set;
CancelUpdate remains available. Counter fields expose readonly17/False or
writable49/True attributes/data capability, with their existing value-write
restrictions unchanged. This extends the documented native contract, without
a fresh Windows readonly/locking oracle.

Detached and explicitly owned native TableDefs expose local Attributes0 and
empty Connect/SourceTableName. Same-value assignments are harmless; nonzero
attributes/nonempty external connections refuse3251. Imported/unproven source
metadata are not guessed. Attached-builder delegation and reentrant attachment
checks retain the actual table identity. External link providers/system flags
remain open. Database/TableDef.Updatable setters are readonly383.

Existing DAO1/1 (1.66s), app built. Actual query_only changes while an edit is
pending, CancelUpdate, metadata refresh, file reopening with QSQLITE_OPEN_READONLY,
readonly counter fields, schema/DML/query-catalog refusal, no lost/allocated
records and byte-for-byte unchanged technical database file passed. No GUI/broad
rerun; latest GUI/DAO before this change2/2 (5.82s).456/384/72 unchanged.

## Native Single creation and consistent supported writes

DAO/new-table/field designers now create Single6/Size4 fields. A REAL-affinity
SINGLE_REAL declaration retains explicit Single provenance and an actual check
that rejects unquantized/invalid values. Shared field conversion rounds to
binary32 before SQL storage; DAO, Qt datasheet editors, checked direct SQL
VALUES/UPDATE and the supported physical INSERT SELECT path share it. The SQL
field callback returns a Double expansion of the rounded float rather than
truncating it to an integer. Source SINGLE metadata also selects this converter
on proven REAL backing. Direct physical-column query lineage preserves Single
type/value identity; compound/unproven expressions are not guessed.

Native literal defaults are stored after quantization, including post-creation
DefaultValue changes/rebuilds. Shared record defaults and supported omitted-
column/DEFAULT VALUES SQL inserts also normalize literal defaults from older
Single-declared tables. Null stays Null; Boolean conversion retains True=-1.
Overflow6, invalid input13, existing unvalidated underflow3251 and negative
zero3251 do not silently damage a prior value. Raw native SQLite statements on
new constrained fields can refuse an unquantized value instead of coercing it.
Legacy unconstrained/raw SQLite and unsupported complex/joined write shapes,
full arithmetic/coercion/locales/rendering and Access reference parity remain
open. Enforced Single relationship keys remain explicitly unsupported.

Final existing GUI/DAO2/2 (6.12s), app built; preceding import target1/1 (0.06s),
not rerun after the lineage fix. Actual cases verify default0.1, 16777217→
16777216, update refusal/retained data, Null, negative zero/underflow, typed
projection, bounded append, post-creation defaults and omitted legacy defaults.
GUI creates Single, stores a rounded data-sheet edit and rejects overflow, then
retains the field through index/field/table schema edits. The test model now
uses its required manual-submit mode. No fresh3/3/full-suite/Access oracle claim.
F014/F015/F041/F055/F105/F121/F122/F175 and fixed456/384/72 remain open/unchanged.

Own native quantized Single keys now also support real relationship enforcement
and cascades when both sides have proven native SINGLE_REAL declarations.
Imported/legacy REAL declarations remain refused for this purpose. Current
GUI/DAO2/2 (6.26s), app built; see DAO_RELATIONS_RUNTIME.md.456/384/72 unchanged.

## Native numeric field ValidationRule

Own numeric non-counter fields support a bounded ValidationRule in DAO and the
field designer, on creation and through atomic existing-field property changes.
The accepted grammar is explicit numeric comparisons (=, <>, <, <=, >, >=)
joined by And/Or, or Between lower And upper. And has higher precedence than Or.
Up to16 comparison terms and255 UTF-16 units are accepted; finite invariant
numeric constants only. Functions, field references, subqueries, arbitrary SQL,
locale-dependent text, date/text rules and parentheses refuse3251.

Rules become real SQLite CHECK constraints through an owned deterministic
validation function. Null satisfies this field rule; Required is independent.
SQL, DAO, datasheet and form saves therefore share storage enforcement rather
than relying on a designer-only warning. Violations report3317. A failed
multirow SQL statement remains atomic. Editing a rule checks existing rows during
the protected table rebuild, and failure retains the old definition and values.
Clearing a rule removes its constraint. Supported field/table renames preserve
rules; unknown CHECK constraints still block unsupported migrations.

ValidationText, arbitrary Access expression rules, table-level rules, imported
validation extraction and exact Windows diagnostics/timing remain open. This is
an implementation within F015/F121, not completion of either full requirement.

Existing GUI/DAO2/2 passed (7.08s), app built. No new CTest target or full-suite
claim; fixed catalog456/384/72 unchanged.
