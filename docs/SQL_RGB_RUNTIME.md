# Native SQL RGB expressions

VBA RGB and Access SQL RGB now call the same native channel-conversion primitive.
The Access query compiler rewrites function calls to the private SQLite adapter;
quoted RGB field names remain identifiers. Case-insensitive calls, parameterized
queries, alias reuse and compiled views use actual registered SQLite functions,
not precomputed display values. Raw SQLite callers use the private adapter name.

SQL integer/real/text channels use the existing VBA scalar/numeric conversion,
nearest-even rounding and documented >255 clamping. SQL NULL remains Null and
raises 94, unlike VBA Empty (zero). Invalid text or binary channels raise 13,
negative rounded channels 5, nonfinite values 6 and wrong arity 450. The numeric
SQL error protocol now retains 450 rather than degrading it to generic 3146.
Results are exact packed 24-bit integer colors. Complete Jet locale/coercion and
Windows error-order equivalence remain open; this does not introduce SQL object
Variants, Date/Currency fidelity or every VBA function into SQL.

RGB is recognized as a pure scalar expression in the existing bounded alias and
joined-update/view paths. Actual DAO OpenRecordset/Execute calls read calculated
fields, report affected rows and persist real updates. A failed color in an
UPDATE rolls back the statement; joined updates also retain their existing
materialized-source and ambiguity safeguards. Conflicting source colors produce
3251 without choosing an arbitrary source value. This adds no support for
previously rejected target/query shapes. Query-model errors clear stale results.

## Evidence

Targeted core/DAO/query-model/join contracts pass 4/4 (1.07 s). Native SQL tests
cover standard colors, mixed-case calls, quoted RGB fields, aliases, typed query
parameters, clamping/rounding, False/True, NULL, invalid text/binary and arity.
Actual VBA DAO code reads computed fields and writes two rows; a later NULL
channel preserves all preexisting target values. Joined updates calculate colors
directly and through a compiled source view; invalid channels and conflicting
view sources leave targets unchanged. Existing VBA/Qt color behavior uses the
same primitive. Final full validation is recorded in PROGRESS.md.

All 100 original query artifacts contain zero lexical RGB call tokens; original
32-module RGB call inventory is also zero. These are lexical inventories, not
additional original query/business execution. Existing original compatibility
counts are unchanged; no full Access SQL claim follows.

Public function contract:
https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/rgb-function
