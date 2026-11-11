# Proven native selector aggregate types

Min/Max with one argument and native First/Last retain the proven argument type
through the existing bound query plan, DAO values and result metadata. This
covers supported Date, Currency, Boolean, Byte/Integer/Long, Single and Double
sources and already-proven scalar expressions. Min/Max DISTINCT retain the same
argument type. No type is inferred from sample result values or invented for
unknown SQLite columns; exact untyped SQLite integers retain their raw values.

For proven scalar sources, selectors decode each relevant input with the shared
native field rules. Min/Max compare those native values, including Boolean -1/0,
Single rounding and signed Date serial ordering. One owned value per group is
retained and encoded through the shared scalar result path. Unknown sources keep
the existing raw aggregates. Min/Max DISTINCT is equivalent for extrema and uses
the same two-argument internal typed callback without a DISTINCT modifier. Empty
aggregate results remain SQL NULL while keeping proven field metadata. First/Last
include the selected NULL; Min/Max skip NULL as before. Multi-argument SQLite
Min/Max are not classified as selector aggregates. Nested aggregates, unknown
provenance, complete text/collation semantics, full saved-query lineage and an
Access result/order oracle remain outside this scope. First/Last source order
is the existing query execution order; no implicit ordering guarantee is added.

The binding allows surrounding native typed scalar calls to receive the actual
Date/Currency/subtype instead of raw SQLite text/real transport. Grouping, readonly rules and source-generation checks retain their existing bounds.
F067/F068/F107/F014 remain open, fixed catalog456/386/70 unchanged.

App built. Existing DAO/query-model bundle passed2/2 in5.54s (5.04/0.50s).
The final affected DAO-only check passed1/1 in5.38s after preserving generic
multi-argument SQLite Min/Max scalar argument binding. Cases cover actual field
types/values, Boolean ordering with raw1/0, Single rounding, negative Date serial
ordering, Date/Currency nested type inspection, DISTINCT/grouping/empty results,
First/Last NULL and untyped integers above2^53. No new CTest or Access oracle.
