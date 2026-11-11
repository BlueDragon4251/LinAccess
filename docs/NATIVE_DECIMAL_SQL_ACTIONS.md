# Exact-fit physical Decimal SQL actions

Bounded physical INSERT VALUES, one physical explicit INSERT SELECT and direct
UPDATE use the same proven precision/scale and real CHECK constraint as DAO/grid
writes. Known expression kinds accompany values to the Decimal store callback.
A BLOB is decoded only for an expression proven Decimal; arbitrary binary source
values refuse. NULL, exact keys and text retain the existing conversion contract;
values requiring rounding/greater precision refuse6 before a statement completes.

Decimal action syntax is deferred until a physical action adapter proves the
conversion path. Unsupported joined/subquery actions and fallback execution stay
refused3251. A Decimal input needs a Decimal target or explicit conversion;
make-table output has no proven target precision/scale and stays refused. Bare
Decimal criteria require explicit Boolean conversion. Ordinary read-only results
retain their query-only contract.

Existing import scope covers actual source-derived fields, exact arithmetic
UPDATE, VALUES/SELECT append, NULL, typed Delete, scale/precision refusal, binary
input refusal and retained values. DAO/model scopes retain the other action paths.
App/helper/DAO/model/import built; existing bundle3/3 in6.86s
(5.96/0.43/0.46) with the actual Jet fixture. Same existing import/action scope
with actual ACE2010 fixture passed0.36s. Seven protected hashes unchanged. No new CTest/catalog ID; fixed456/386/70, F014/F055/F067/F068.
