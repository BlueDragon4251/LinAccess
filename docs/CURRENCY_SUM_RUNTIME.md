# Exact sums of proven Currency expressions

Bounded SELECT projections Sum(CCur(...)) and Sum(declared Currency/Money
parameter) now use a native integer Currency accumulator. Parentheses/unary
forms and a surrounding CCur retain the proven encoding. This removes residual
binary floating sums such as 0.1+0.2-0.3 and returns fixed Currency through the
existing query metadata/DAO/Qt path.

The accumulator adds checked int64 Currency units via the native fixed-point
runtime. It supports intermediate totals beyond the +/-2^39 output backing while
preserving exact units for cancellation; int64 overflow diagnoses 6. The final
result uses the existing strict lossless REAL output interval and diagnoses 3251
outside it. Null inputs are ignored; no inputs/all Null returns Null. CCur itself
retains its Null diagnostic, so nullable columns in Sum(CCur(column)) require
the appropriate non-NULL source filter.

currency_sql_sum reproduces the prior residual/generic result (0/1), then checks
exact cancellation, grouped/empty/Null sums, Currency parameters, large positive/
negative cancellation, final-range rejection, int64 overflow and actual read-only
Qt report Currency display. Sum/query/object tests pass 3/3 (0.16 s). The original
regression independently sums CSV decimal units in a wide integer accumulator:
all 43 original Currency-column sums match the native SQL aggregate, with zero
outside the supported backing interval (7.91 s). The 121613 individual Currency
value comparisons and existing CSV/Long/Boolean/date/counter checks remain active.
GUI suite now has 43 CTests; full validation is recorded in PROGRESS.md.

The compiler rewrites only proven bounded projection expressions, using its
existing token inspection. General implicit field/view type inference, wildcard/
set-operation/alias/complex aggregate cases, SQL money operators/full-range output,
full parser/locale/Windows type/error identity remain open. Plain unproven SQLite
SUM is unchanged. This adds no full original business workflow or blanket money
compatibility.
