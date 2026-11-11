# Native typed filtered and sorted query snapshots

Filtered native query plans use the proven output columns and ordinal result
metadata when those names are unique. An explicit outer projection and typed
criteria retain Date/Boolean/Single/Currency/Decimal provenance. The existing
read-only derived-query wrapper keeps exact Decimal keys inside SQLite through
filtering and sorting; values decode to canonical text/Decimal only at the result
boundary. No temporary physical Decimal table or writable Decimal field is made.

DAO Filter/Sort plus OpenRecordset(dbOpenSnapshot) now support these exact Decimal
query results. Derived cursors retain explicit snapshot type, read-only status,
field types and existing independent cursor handling. Repeated filtering/sorting
wraps the already-validated source plan; this is a trusted internal wrapper, not
permission to compile arbitrary nested Decimal SELECT statements. Source parameters
and established TOP/ordering/source/budget checks keep their existing bounds.

Duplicate Decimal output names refuse3251 rather than choosing a field. Other
insufficiently proven outputs retain the existing filter path; no type is inferred
from fetched rows. Exact Currency-tick projections retain their prior refusal.
Full parameterized/qualified filter grammar, arbitrary nested/saved-query lineage,
editable derived cursors and complete Access snapshot lifetime/order parity remain
outside this scope. Physical Jet Decimal creation/precision-scale/round trips stay
open.

The existing DAO/query-model scopes cover numeric Decimal sort order, chained
snapshots, exact values up to96 bits, NULL/empty results, native Boolean/Single
filtering, field types/read-only status, Find on a derived Decimal cursor and
unchanged source rows/cursor. No new CTest or catalog ID; fixed456/386/70 unchanged.
F014/F041/F055/F067/F068 remain open.

App built; final existing DAO/query-model bundle passed2/2 in5.91s
(5.46s/0.44s). A Jet-versus-SQLite identifier quoting error found in the first
bundle was corrected before repeating only the affected targets.
