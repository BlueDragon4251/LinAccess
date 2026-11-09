# Proven Jet4 field captions/descriptions

The existing read-only MDB helper reads Caption and Description through
the actually available mdb_col_get_prop API. The normal metadata response has
an optional field_properties array and field_properties_format=JET4. It binds
every actual ordered column, including proven absence; empty strings remain
present values. Other formats omit this component, so absence of the component
does not masquerade as missing original properties. No extra helper process per
table, VBA extraction or row-value export is added.

The parser checks the version/format, source digest representation, exact column
order, unique names and bounded Unicode strings. Values are limited to4096
UTF-16 units/16KiB UTF-8 per property, within the existing1MiB helper output cap.
This preserves MDBTools' decoded text, not every original binary property-blob
representation. Existing whole-file digest/copy/import-transaction guards apply.
The copied schema stores provenance, actual values/absence and lease identities.
Original-source properties are readonly; Inherited is explicitly unproven. The
existing own-property editor/limits remain separate.

Datasheet headers/hints and native form/report creation use actual imported
captions. DAO Field.Properties can read the proven source strings; missing
supported properties report3270, and mutation refuses383. Older projects without
this metadata are not silently backfilled or assigned invented source captions.
Full field property lists/Count, arbitrary types, Jet3/ACE blobs and Access2003
runtime/rendering equivalence remain open within F015/F080/F121.

The optional zero-row generator tests/jet_binary_reference/JetFieldDisplayReference.java
uses the already documented Jackcess5.0.4 artifact/checksum in JET_BINARY_IMPORT.md.
It independently reopens the real Jet4 MDB and checks Unicode/quotes, description
newline, absent Caption and explicitly empty Caption. It contains no business
records or VBA. No generated MDB/Java runtime dependency is shipped.

Compile/run the generator into an empty private scratch directory as documented
for the other technical Jet references; set LINACCESS_JET_FIELD_DISPLAY_REFERENCE
to its field-display-probe.mdb, then run the existing table_import_contracts target.
The existing case compares native extraction, real MDB→SQLite import, DAO readonly
metadata, actual model headers and unchanged source hash. Existing import/DAO2/2
passed (2.57s), app/helper built. No new CTest target, full-suite, platform-matrix
or Windows oracle claim; fixed catalog456/384/72 unchanged.

## Jet4 Required and AllowZeroLength

The same bounded component also preserves explicitly present Required and
AllowZeroLength booleans. The real MDBTools representation was checked against
an independently reopened zero-row Jackcess Jet4 file: yes/no, not guessed
numeric or localized spellings. Unknown representations fail; absent properties
remain unknown. AllowZeroLength is accepted only for proven text/memo columns.
Imports enforce Required using actual NOT NULL and disallowed empty strings using
the existing linaccess_jet_nonempty_text CHECK (NULL remains independent).
The DAO field and intrinsic Property readers expose the proven booleans readonly.
Original ownership is explicit in the field-designer metadata, so an imported
field never becomes editable merely because Required and Attributes are known.
Older saved projects are not silently assigned these properties or constraints.
The optional development generator JetFieldConstraintReference.java contains no
records/VBA and refuses replacement. Set LINACCESS_JET_FIELD_CONSTRAINT_REFERENCE
to its generated field-constraints-probe.mdb for the existing import target.
Full property blobs, source mutations, validation expressions, collation and
Access2003 runtime equivalence remain separate unproven requirements.
