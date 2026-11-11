# Proven native types in physical SQL actions

Bounded direct-table UPDATE expressions and WHERE predicates now share known
Date, Boolean, Single, Currency, integer and Decimal source kinds with SELECT.
Direct DELETE WHERE predicates use the same expression compiler, including
physical aliases and parameter declarations. Native conversion guards on target
fields remain in place. Source typing requires an unshadowed physical main table.

This prevents stored Boolean/Date/etc. values from silently acquiring SQLite's
scalar kind while evaluating VarType, arithmetic or conditions in an action.
It does not introduce Decimal target coercion, joined/delete queries, additional
subquery grammar or original-format writes. Existing action transactions and
rollback remain responsible for atomicity. NULL semantics are unchanged. Mixed Date/non-Date action comparisons refuse3251
unless explicitly converted: the targeted DELETE exposed an untyped SQLite
text/number fallback. Date/Date and Date/NULL retain their existing paths. This
strict check is scoped to the already resolved action-expression compiler, so
field types are established before it runs. General read-only mixed Date
comparison compatibility remains outside this scope.

App/helper/DAO/model/GUI built. Existing GUI passed18.36s; after localizing
the action Date guard, only affected DAO/model repeated: 2/2 in6.19s
(5.74/0.45). Exact Date/Boolean/Single/Currency updates, alias Delete, NULL
retention and rejected mixed Date action are covered. Seven protected hashes
unchanged. Fixed456/386/70, no new CTest ID or Windows oracle.
