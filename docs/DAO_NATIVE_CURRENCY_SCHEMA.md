# Native Currency schema preservation

Known native Currency defaults now reconstruct from the actual checked storage
value to canonical Currency text before schema serialization. The same helper
handles original and modified field definitions and no-op/default comparisons.
New defaults use checked Currency text after storage conversion, preserving the
four-place creation contract rather than treating binary64 serialization digits
as a new excess-precision source literal.

Existing guarded reconstruction paths can therefore preserve native Currency
columns during DefaultValue/Required changes, ordinal reordering, index append/
delete and table rename. Field rename keeps its checked SQLite path. All native
schema equivalence, exact value/rowid digests, savepoints, constraint enforcement,
index metadata and connection restoration checks remain. Unsupported original/
custom schemas, full-range Currency storage, numeric type conversion involving
Currency and Currency relationship-key/Seek semantics remain bounded. Attached
DAO Field.Type remains read-only383; designer conversion remains separately3251.

App/DAO built. The first existing DAO bundle reached a fixture error-code
assumption after the actual schema operations: attached Type correctly refuses
383. That fixture was corrected and the actual designer conversion3251 was also
checked. Only affected DAO was rerun, passing1/1 in3.71s. Independent raw SQL
compares exact binary64 bits of existing Currency payload/default values across
actual default/Required/reorder/index changes, checks rollback after tightening
Required over a NULL row, and verifies renamed fields/table plus subsequent
new-row defaults. No new CTest/catalog IDs;456/386/70 unchanged; no Windows schema
oracle/full Jet migration claim. Last checked DEB/Original remainsa23c7009 before
these later changes; private published predecessor90f672b9 was tree-verified.
