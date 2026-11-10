# Actual First/Last values within fixed pivots

Fixed-heading TRANSFORM First/Last now uses the existing native first/last
aggregate with a SQLite FILTER for the actual pivot match. Nonmatching source
rows never enter the selector. A matching Null remains a real first/last value;
it is not skipped as it would be by Min/Max. Replacing nonmatching values with
Null inside a CASE expression would select the wrong row and is not used here.

The existing checked physical-field conversion and result type are preserved
for native Currency, Single, Byte/Integer/Long, Boolean and date fields. Real DAO
Fields.Type/Value exposes the resulting native type. Unproven/text/Decimal
inputs, dynamic headings, joins and saved-view source typing retain the existing
explicit boundaries. The SQLite FILTER syntax is generated internally for this
bound transformation; domain-expression clause validation remains unchanged.

Selection follows the actual native input visitation order. Access itself does
not promise a meaningful sorted First/Last result merely from the final output
ORDER BY; this scope does not assert a Windows/Jet physical-order oracle or
invent an implicit user ordering.

The existing currency_schema_queries case places nonmatching rows before/after
matching rows, includes matching Nulls, missing pivots and absent groups, and
checks Integer/Currency/Date selections against independent native rowid scans.
It also checks real DAO types/Variant values and all previous fixed aggregate
cases. No new CTest or catalog ID. Full TRANSFORM remains open under F068.
