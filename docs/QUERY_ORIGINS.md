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
wildcard logical lineage and deeper qualification remain incomplete.
Ambiguous repeated origins are rejected by the normal SQL error path.

Earlier SELECT aliases can be expanded when their expressions use only known
pure functions and no source column shadows the alias. Expressions involving
live object callbacks, aggregates, unknown functions or volatile calls are not
duplicated. Expression and total expansion sizes are bounded. Ordinary owned
views are repaired in dependency passes; user views and tables are not replaced.

## Original regression evidence

All 100 unchanged Kassa queries remain hash-checked on an isolated MDB import.
The run executes 73 without extra parameters and five with parameters.
For those five, LinAccess takes a non-Null value from the unfiltered results.
The DAO filtered row count matches a positive, independently computed C++
filter count. Reports omit the parameter values. Synthetic tests separately check exact returned values,
alias dependencies, nested derived tables, duplicate preservation, ambiguous
origins, and that a live callback is not expanded/evaluated during metadata work.

One of the six original parameterized derived-table queries remains unresolved
(artifact SHA-256
`4a8db1b0b4f74446d4d95098591dbb4f36c871de05533f3c182b23bfa6127e6e`).
It requests a column absent from its referenced view's six exposed columns.
The same schema mismatch occurs in two other artifacts. The sanitized Access 2003
DAO comparison reports 3061 for all three. The adapter now reports 3061 when
SQLite rejects a qualified field explicitly present in the prepared SELECT.
The corpus regression still verifies the omissions against the actual saved
view or table schema. These queries remain unresolved.

Remaining corpus classifications: nine missing query dependencies, eight
binding/context failures, three source-column mismatches, and two joined
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
implementations. At the conditional-alias checkpoint the corpus executed 77/100 (72 ordinary plus five with
independently checked data parameters). Overall execution counts have lower-bound
regression assertions. This demonstrates native execution on actual rows, not
Microsoft Access result equivalence or exact Currency arithmetic in SQL.

Validation for the conditional-alias stage: all 18 CTests passed (56.43 seconds),
a fresh Debian package passed current executable Build ID/content/permission
checks, and all four original fixture SHA-256 checks passed unchanged. The
GUI-disabled runtime sources did not change in this SQL-only stage.

## Three-part saved-query qualification (2026-11-03)

A visible saved-query/view source can qualify a direct output field by its
source relation: query.relation.field. The binder resolves that three-part path
to the exact actual SQLite output label, including duplicate labels such as
id:1. Source aliases are retained; a unique match is required within one visible
view. Four-part paths, ambiguous owners, duplicate origin projections and unknown
origins retain checked SQL failure rather than choosing a candidate.

SQLite physical metadata alone cannot distinguish two input views over the same
base table. The binder therefore follows direct, checked projections in the
actual stored temp/main view definitions, matching their positions to prepared
output labels. Direct wrappers propagate verified paths through ranked TOP
views. Relation aliases/input labels must resolve uniquely. Traversal is bounded
to 16 levels and never steps value expressions. Compound/nested SELECTs,
computed projections and wildcard expansion do not supply inferred logical paths.
Physical lineage for this new three-part path is used only through proven direct
projections; UNION's first reported physical origin is not treated as a complete
origin proof. Existing ordinary alias expansion retains its purity guards.

Unchanged original artifact
32b25d06328e208b7d2e94ec33a17422fca977ef7b5988e6b040123562c890ec now
executes on the full isolated import. Synthetic real SQL checks distinct values
from duplicate labels and two view inputs over the same base table, WHERE/ORDER
binding, and refusal of ambiguous/unknown/compound/overqualified paths.
The current original corpus is 78/100 executed: 73 ordinary and five with checked
data-selected parameters. The Access 2003 DAO comparison establishes the error
boundary described above. Result equivalence remains open.

Three remaining source-column cases are now independently classified: two
unchanged saved source definitions omit the requested field, and one requested
field is absent from the original MDB table schema as well as its import. The
runner checks the original bytes/schema; it does not invent another column or
infer a complete Access application state. Nine missing dependencies, eight
binding/context failures and two joined-action guards remain alongside those
three cases. These 22 cases are not claimed executed.

Full desktop suite passes 48/48 (90.47 s), GUI-free core 1/1 (1.67 s), Debian
payload/permissions/Build-ID/pinned SQLite checks and all four original fixture
SHA-256 checks pass. No new test targets or original source changes.
