# Exact Currency query wrappers and source-kind propagation

Full signed64 Currency parameter projections retain their checked binding source
through bounded derived Filter/Sort wrappers. The generated SQL and binding proof
are carried together, including a hash of the regenerated original binding SQL;
a changed SQL plan or substituted original expression refuses3251. Read-only DAO snapshots can
filter, sort, chain snapshots and use typed Find without converting ticks to REAL.

Known CURRENCY_TICKS inputs preserve Currency through CVar/CCur and supported
scalar calls, including exact CStr and VarType. Comparisons use proven Currency
operands or explicit Decimal conversion; unknown numeric mixing refuses3251.
NULL keeps existing SQL propagation. Native declared tick sources order by their
actual signed64 values. Proven Currency arithmetic, checked negation and SUM are now documented in
EXACT_CURRENCY_FIELD_ARITHMETIC.md. Lossy CAST, remaining aggregates and physical
SQL target conversion remain refused. Existing bounded
parameter-only arithmetic is unchanged. Original-format MONEY import, arbitrary
promotion/locale/round-trip/Windows equivalence remain outside this scope.

App/helper/DAO/model/parameter targets built. Existing DAO/model scopes passed
(5.58/0.61s); after correcting the stale SQLite True=1 expectation to the already
implemented Access True=-1 and refusing the uncovered raw SUM path, only the
affected parameter scope was repeated: final0.27s passed, including binding-root
substitution rejection. No new CTest ID; seven protected hashes unchanged. No new ID;
F014/F054/F055/F067/F068 remain open, fixed456/386/70 unchanged.
