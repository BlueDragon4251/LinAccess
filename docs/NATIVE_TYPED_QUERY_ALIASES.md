# Proven native types through computed AS aliases

Unique explicit AS aliases for supported native scalar results retain their
proven type in subsequent SELECT expressions and native WHERE/HAVING predicate
binding. Supported kinds are Date, Boolean, Single, Currency, Decimal and the
existing proven integral/Double kinds. No type is inferred from result samples.
The existing closed pure-expression alias expansion still controls SQL expansion;
this change supplies type provenance rather than inventing a separate evaluator.

Only unique projection names without a source-field collision acquire new alias
bindings. Source columns retain their existing lookup behavior. Non-Decimal
ambiguous/shadowed aliases keep the previous path. Decimal retains its stricter
explicit collision refusal because exposing its binary transport would be unsafe.
Bare implicit aliases, forward references, complete nested/saved-query lineage,
volatile alias evaluation order and general Jet name resolution remain outside
this scope. General AS alias support does not imply writable computed columns.

The existing DAO scope checks actual Date/Boolean/Single/Currency values through
CVar aliases, subsequent Year/VarType expressions, native membership/range filters,
ordering and source-column shadowing. No new CTest/catalog item or Windows oracle.
F014/F041/F067/F068/F107 remain open; fixed456/386/70 unchanged.

App built; existing DAO/query-model bundle passed2/2 in5.83s (5.42s/0.40s).
