# Verified Currency columns in simple table queries

Read-only query preparation now resolves Currency columns from the live physical
table's verified source metadata. For bounded single-table SELECT projections,
plain/qualified/aliased Currency columns retain fixed values and dbCurrency;
Sum(column) uses the exact integer aggregate without requiring CCur(column).
Native DAO, query-table and read-only Qt form/report consumers receive the proven
types through the existing result descriptor path.

The resolver requires one physical table with verified CURRENCY/REAL metadata.
Joins, views, nested SELECTs, HAVING, set operations and repeated SUM calls in
later clauses stay outside this bounded context. Unmarked REAL columns retain
their generic type. Wildcard/ambiguous projection safeguards remain active;
this is not a complete SQL binder/type system or SQL money operator engine.

Original SQL and compiled-statement identity are retained. Changing a plan's SQL
invalidates its old descriptors instead of recompiling a stale statement. A
previously schema-resolved source is refreshed before reuse, so removed/restored
Currency metadata does not keep stale type assumptions. Filter wrappers preserve
the resolved source descriptors without rerunning the original statement.

currency_schema_queries first reproduced missing implicit Currency metadata
(0/1). New/previous Currency tests pass 4/4 (0.21 s), including qualified aliases,
Nullable columns, implicit exact sum, filters, modified plans, metadata refresh,
plain REAL/join/HAVING exclusions, invalid stored values and actual read-only Qt
form/report fields. Expanded source/query/object/sum/original layout/Kassa/table
checks pass; full-suite validation below rebuilds every test against the final
libraries. The original regression now uses Sum(column) directly: all 43 Currency
column sums still match independent CSV integer sums, with no explicit CCur cast.
All 121613 individual Currency value checks remain active. GUI suite has 44 CTests;
full validation is recorded in PROGRESS.md.

General field/view/join/alias/expression type propagation, writable saved queries,
action-query target coercion, complex Currency aggregates/arithmetic, full-range/
regional/Windows identity remain open. No complete original business workflow is
added by this bounded live-schema path.
