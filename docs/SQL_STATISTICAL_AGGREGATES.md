# Native SQL Var and StDev

SQL queries now accept Var(numeric_expression) and StDev(numeric_expression),
including grouping, DISTINCT, aliases, bound parameters and normal saved queries.
Both use the same native sample-variance accumulator already used by report and
domain aggregates. Results are Double, with divisor n-1. NULL values are omitted;
fewer than two numeric values return NULL, including an empty ungrouped result.
The compiler retains the original calculated column labels and proves the Double
result type before reading records.

A centered online calculation uses extended intermediate precision and constant
state per group. It avoids the catastrophic subtraction in Sum(x*x)-Sum(x)^2/n
for ordinary values with a large common offset. The final Double must be finite;
nonfinite input or final overflow raises6. Text/BLOB inputs raise13 instead of
silently becoming zero. State is freed by the SQLite aggregate finalizer on both
success and failure. Existing SQL execution/memory/result bounds still apply.

These SQL functions consume numeric SQLite values. This does not add implicit
locale-dependent text/date conversion, exact Decimal variance, population
VarP/StDevP, or general TRANSFORM statistical aggregates. Currency REAL values use
the existing native storage limits and are evaluated numerically; this is not an
additional exact Currency arithmetic claim. Report/domain behavior is unchanged.

The existing access_query_model test covers independent numeric expectations,
NULL/empty/singleton groups, large offsets, DISTINCT, parameters, text/BLOB refusal
and overflow. Existing Currency SQL targets check the surrounding execution path.
No new test ID or business fixture is introduced.

Microsoft documents sample/population distinction and Null handling here:
- https://learn.microsoft.com/en-us/office/vba/access/concepts/criteria-expressions/var-varp-functions-microsoft-access-sql
- https://learn.microsoft.com/en-us/office/vba/access/concepts/criteria-expressions/stdev-stdevp-functions-microsoft-access-sql

These documents were read during implementation; there is no new executed Access
2003 oracle or full SQL compatibility claim. F067/F068/F071 remain open and the
fixed456 catalog is unchanged.
