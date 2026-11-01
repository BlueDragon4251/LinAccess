# Jet derived tables and query column origins

Original SQL bytes are preserved. The adapter recognizes Jet's bracketed
`[SELECT ...]. AS alias` derived-table syntax, including nested brackets and
quoted text. The inner SELECT is compiled independently. Explicit inner
PARAMETERS and TOP materialization remain unsupported; nesting is bounded.
An ordinary bracketed identifier starting with SELECT remains an identifier
unless the derived-table terminator is present.

Jet ignores DISTINCTROW for a single source and when all fields of all sources
are output (`SELECT DISTINCTROW *`). Native tests retain duplicate inner-join
rows and duplicate unmatched outer-join rows. Projecting only some joined source
records still requires full Jet record-identity mapping and stays unsupported.

Saved queries sometimes refer to the original base table through another saved
query. The adapter obtains actual SQLite column table/origin metadata by
preparing a zero-row projection, without stepping the statement. A base-field
reference is rebound to a visible source column only when exactly one lineage
match exists. Source aliases and output labels are retained. Explicit/nested
JOIN sources are supported; comma joins, nested scopes, computed-field origins,
intermediate view lineage and more complex qualification remain incomplete.
Ambiguous repeated origins are rejected by the normal SQL error path.

Earlier SELECT aliases can be expanded when their expressions use only known
pure functions and no source column shadows the alias. Expressions involving
live object callbacks, aggregates, unknown functions or volatile calls are not
duplicated. Expression and total expansion sizes are bounded. Ordinary owned
views are repaired in dependency passes; user views and tables are not replaced.

## Original regression evidence

All 100 unchanged Kassa queries remain hash-checked on an isolated MDB import.
The current run executes 72 without extra parameters and five with parameters:
77 total, without a Microsoft Access result oracle. For those five, a non-Null
parameter is selected from actual unfiltered results; the DAO filtered row count
must be positive and equal an independent C++ filter count. Values are not
included in reports. Synthetic tests separately check exact returned values,
alias dependencies, nested derived tables, duplicate preservation, ambiguous
origins, and that a live callback is not expanded/evaluated during metadata work.

One of the six original parameterized derived-table queries remains unresolved
(artifact SHA-256
`4a8db1b0b4f74446d4d95098591dbb4f36c871de05533f3c182b23bfa6127e6e`).
It requests a column absent from its referenced view's six exposed columns.
The regression checks that schema mismatch and reports error 3146. It does not
substitute a similarly named column from another table. A Microsoft Access
comparison is required to distinguish an original query/source mismatch from
additional Access name-resolution behavior. This query is not counted executed.

Remaining corpus classifications: nine missing query dependencies, eleven
other Jet/binding failures, that one source-column mismatch, and two joined
UPDATE boundaries. The runner now attempts the INNER JOIN action on an isolated
full data copy; differing candidates are rejected with 3251. The other action
requires an outer join and a saved-query target (3251). One unchanged original
INNER JOIN action separately changes one independently verified value using its
actual table schemas and synthetic unique matches; that separate prepared case
does not increase full-data corpus coverage.
No complete Jet/ACE semantics or result equivalence is claimed.


## Conditional projection aliases

The known-pure expansion whitelist now includes native Jet IIf. Every nested
function is still inspected; a live Forms/Reports callback, volatile function,
unknown function or aggregate prevents expansion. A real source column retains
priority over a same-named projection alias. Synthetic SQL tests verify explicit
independent tax totals across ordinary, zero-rate, Null-rate and Null-amount rows;
no live callback runs while preparing a rejected alias chain. SQL numeric literals
such as `.2`, `-.25` and `.5e+2` remain single tokens through rewriting.

Three newly executing unchanged original queries are required to produce positive
record counts and a non-NULL first-row value on the imported original MDB copy:

- `aee3c2eb8334d24f8268bda68943e5c59c5a993da65d96ea824174da55ed16bf`
- `c603f4a1eb63c4c0ad7a7427a87ae0978485d1bbea492b1caebedb387ad52c00`
- `e074e74467212cd73520402ffb4239aef220557883d048a2be672de10c3d48de`

These include a dependent lookup query; they are not three independent IIf
implementations. The corpus now executes 77/100 (72 ordinary plus five with
independently checked data parameters). Overall execution counts have lower-bound
regression assertions. This demonstrates native execution on actual rows, not
Microsoft Access result equivalence or exact Currency arithmetic in SQL.

Validation for the conditional-alias stage: all 18 CTests passed (56.43 seconds),
a fresh Debian package passed current executable Build ID/content/permission
checks, and all four original fixture SHA-256 checks passed unchanged. The
GUI-disabled runtime sources did not change in this SQL-only stage.
