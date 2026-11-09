# Fixed-heading native Count crosstabs

The Jet query compiler accepts bounded TRANSFORM Count(field) / Count(*) queries
with simple grouped row fields and a fixed PIVOT field IN ('heading',...) list.
It rewrites tokens into conditional counts, preserving heading order, literal
escaping, explicit aliases, WHERE/HAVING/ORDER clauses and declared parameters.
Count(field) omits NULL values; Count(*) counts matching rows. Absent categories
produce zero counts. Names are checked for case-insensitive duplicates and row
heading collisions; the output is limited to 255 columns and 64-character pivot
headings. Count outputs have proven native Integer4/DAO Long type, including an
empty result. QueryDefs classify as dbQCrosstab (16).

Direct DAO recordset source text and saved native QueryDefs use the same query
compiler. Aggregate results remain readonly. Simple source-column lineage uses
the existing SQLite origin metadata path. Native execution budgets, parameter
binding and result bounds remain in effect.

Dynamic headings, numeric heading literals, numeric aggregates, expression row
headings, compound/join sources and arbitrary TRANSFORM grammar refuse3251.
This implements native fixed-heading counting; no new Windows Access result,
collation, regional heading or full crosstab compatibility oracle is claimed.
F068/F067 remain open; no fixed-catalog point is added.
