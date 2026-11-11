# Native typed DAO Find criteria

FindFirst/FindLast/FindNext/FindPrevious compile criteria against the recordset's
already-proven native field metadata. Each candidate still supplies its actual
cached values to a prepared SELECT; the search does not infer types from samples,
create physical Decimal columns, materialize a new business database or modify the
source records. Existing cursor direction, NoMatch, deleted-position handling,
revision checks, pending-edit behavior and execution/time budgets are retained.

Native Boolean, Date, Single and Currency comparisons and membership/range
predicates decode source kinds before evaluating criteria. Exact Decimal query
values now participate in Find, including ordinary numeric comparisons and native
IN/BETWEEN, without passing through Double. Their canonical Qt text is decoded
using the explicit Decimal field metadata. SQL subexpressions without a proven
Decimal adapter refuse3251 even when Decimal appears only as a field reference;
CAST-to-REAL no longer bypasses the exactness boundary in that context.

Exact Currency-tick projections retain their prior explicit refusal. Typed Decimal
Filter/Sort derived snapshots are separately documented in
[Native typed query snapshots](NATIVE_TYPED_QUERY_SNAPSHOTS.md). Full locale/wildcard/
qualified-field/aggregate Find semantics and complete Access error-order/evaluation
parity are not claimed. Unsupported Decimal LIKE requires an explicit CStr.

The existing DAO scope checks actual typed matches, NULL nonmatches, searches in
both directions, NoMatch/EOF recovery, exact96-bit Decimal matching, decimal scale
comparisons and position preservation after a refused lossy cast. No new CTest
or catalog ID; F014/F041/F055/F067/F068/F373 remain open. Fixed456/386/70 unchanged.

App built; existing query model0.42s and final affected DAO1/1 in5.99s passed.
A CAST fallback bypass found by the bundled negative case was fixed directly.
