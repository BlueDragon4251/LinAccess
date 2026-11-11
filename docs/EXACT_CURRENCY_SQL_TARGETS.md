# Input-aware native Currency SQL targets

Bounded physical INSERT VALUES, one explicit INSERT SELECT and direct UPDATE
carry their proven expression kind into Currency conversion. Native tick targets
receive signed64 units; legacy REAL targets receive the existing checked dollar
encoding. This prevents copying12500 internal ticks as12500 dollars. Proven
Decimal input uses the existing Currency rounding/range contract. Binary input
without numeric provenance refuses13; overflow6 and legacy range3251 preserve
rows. NULL and already encoded physical DEFAULT VALUES retain their meaning.

Native tick contexts refuse unsupported joined/multi-source action fallback.
Access2003 action SQL does not expose SQLite RETURNING; that clause is refused
before execution. Checked path acceptance is recorded explicitly, including a
statement whose rendering happens to equal the input. Other target kinds need
explicit conversion from Currency ticks. Original-format MONEY import, arbitrary
parameter/dialect/locale/default/round-trip parity remain outside this scope.

App/helper and existing parameter/insert/update/DAO/model targets built.
Insert/update/DAO/model passed (0.13/0.14/6.25/0.50s). The parameter
check exposed action-local pure-projection tick inference conflicting with the
outer REAL parameter binding. That optimization now stays outside action
expressions; only the affected parameter scope repeated and passed in0.18s.
No new CTest ID;
F014/F055/F067/F068 open, fixed456/386/70 unchanged.
