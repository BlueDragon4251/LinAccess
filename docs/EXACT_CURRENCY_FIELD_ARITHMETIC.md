# Exact Currency source arithmetic and sums

Proven Currency operands with at least one CURRENCY_TICKS source support +, - and
* through the existing native VBA Currency engine. Results retain signed64 ticks
and Currency type; NULL propagates. Unary negation uses checked Currency negation,
including error6 for the minimum value. Unproven numeric mixing still refuses.

SUM of a declared tick source uses the existing checked Currency accumulation
policy with integer input/output, NULL skipping and NULL for no values. Integer
encoding is required; every addition uses the checked native Currency operation.
Overflow6 is retained, including intermediate overflow; arbitrary Jet aggregate
ordering/equivalence has not been measured. Legacy REAL Currency SUM keeps its
previous transport and range contract. AVG/statistics/CAST and physical tick SQL
target conversion remain outside this scope.

App/helper/parameter/DAO/model/sum targets built. Existing parameter/DAO/model
bundle3/3 in6.89s (0.13/6.30/0.45), existing legacy sum scope0.09s passed.
Actual signed64 endpoints, last1e-4 digits, unary/operation/sum overflow6,
NULL accumulation and arithmetic in a checked derived filter are covered.
All seven protected hashes unchanged. No new ID;
F014/F055/F067/F068 remain open, fixed456/386/70 unchanged.

Bounded physical Currency SQL target conversion is now described in
[EXACT_CURRENCY_SQL_TARGETS.md](EXACT_CURRENCY_SQL_TARGETS.md).
