# Exact Decimal values in native SQL

Native SELECT expressions use the existing checked 96-bit Decimal engine for
CDec, addition, subtraction, multiplication, division, unary signs and comparisons.
Supported typed scalar conversions and Min/Max/First/Last retain the Decimal kind.
DAO exposes proven computed Decimal results as dbDecimal (20); the query grid
shows exact canonical decimal text and keeps computed results read-only.

The internal DECIMAL_EXACT transport is a 24-byte big-endian sortable key:
(value × 10^28) + 2^191. Equal numeric values share a key, so SQL ordering and
grouping compare numbers without a Double intermediary. Decoding validates the
length, Decimal range and exact re-encoding. Result values cross the Qt boundary
as canonical decimal strings, with their explicit Decimal metadata retained.
Trailing-zero scale and signed-zero representation are not preserved by this
transport; numerical exactness does not establish Access formatting identity.

Supported SELECT projection aliases retain their proven types. Unknown functions,
unsupported aggregates such as Sum, nested SELECT/UNION, action/make-table queries
and unproven Decimal predicates refuse instead of executing SQLite blob arithmetic.
Physical Jet Decimal creation, import, precision/scale properties, writable Decimal
fields and lossless MDB round trips remain open. Derived DAO Filter/Sort snapshots
and textual Find also refuse when Decimal lineage cannot be retained. Field.Size
is not fabricated from the query transport.

Existing core, DAO and query-model targets cover the codec, 96-bit limits, tiny
fractions, values above 2^53, arithmetic, rounding, error boundaries, type metadata,
ordering, grouping, selectors, aliases and unchanged data after refused writes.
These are native implementation checks, not new Windows Access observations.
The fixed 456-item catalog remains unchanged: 386 confirmed partial scopes and
70 open full requirements. This extends existing F014/F041/F067/F068 scopes.

Validation: application build succeeded; existing core target passed in0.58s,
query model in0.45s and final affected DAO target in5.41s.
