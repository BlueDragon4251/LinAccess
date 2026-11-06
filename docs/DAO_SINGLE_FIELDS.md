# Proven direct-table Single field values

Physical main-table DAO recordsets now apply the existing checked VBA Single
conversion when the field's captured, validated DAO type code is6 (dbSingle).
Field.Value, recordset bang/default lookup, Clone and staged Edit/AddNew values
therefore return numeric subtype Single rather than a generic Double. Null stays
Null; Empty assignments become Single0. Literal defaults use the same value
read conversion. Assignment narrows before changing the edit buffer, and Update
persists the narrowed binary32 value into the native REAL column. Overflow6,
invalid numeric text13 and unproven underflow3251 preserve the prior buffer.
The existing module conversion supplies these rules; no second converter exists.

There is no inference from stored numbers or generic REAL/FLOAT declarations.
The checked source metadata/native field classifier supplies the type code. A
generic Double field keeps its current behavior. This increment is confined to
direct physical DAO recordsets; arbitrary query lineage, form controls, frozen
projections, SQL arithmetic and all locale/error precedence remain open.

The optional technical Jet4 reference generator
`tests/jet_binary_reference/JetSingleReference.java` creates actual FLOAT columns
with0.1f,16777216f,1.5f and NULL, then independently reopens and compares their
native Float values. It refuses overwriting an existing MDB and uses the already
verified optional Jackcess tool; no Java dependency is shipped in LinAccess.
Set LINACCESS_JET_SINGLE_REFERENCE to single-type-probe.mdb to exercise the
existing table_import_contracts target. It checks real full import, Field.Type6,
Size4, actual Single values/NULL, a narrowed native write/Requery and unchanged
source hash. The existing DAO target also exercises VBA VarType/CSng, Clone,
Edit/Update, AddNew defaults, Null/Empty, refused conversions and Double isolation.

This is not proof that MDBTools exports every Float bit pattern losslessly.
Its1.0.0 data.c MDB_FLOAT formatter uses %.8g. For example, technical IEEE754
bits509022793 (about1.1385786510416537e-20) format as1.1385787e-20, which can
reconstruct a neighboring Float. The original-format/precision import requirement
remains open; the three chosen non-NULL reference values do round-trip. No broad
exact Float storage/import or Windows Access2003 execution claim is made.
F014/F041/F055/F105/F122 remain open; fixed catalog456/384/72.

Focused existing DAO/import targets2/2 (1.94 s) passed with all optional Jet
references enabled. Last full suite4548ce7:48/48 (89.51 s), preceding this change.

Aktuelle App/Paket/Build-ID/SQLite3.54.1 und vier Originalfixture-/drei Uploadhashes bestanden.
