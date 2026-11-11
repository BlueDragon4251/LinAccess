# Native exact Decimal IN and BETWEEN predicates

Proven native Decimal SQL operands use dedicated typed predicates rather than
SQLite comparisons on the internal binary transport. Each operand is evaluated
once by SQLite and decoded using its proven source kind. Supported operands are
Decimal and the existing proven numeric kinds, or literal NULL; unproven/text/date
mixtures refuse3251 rather than choosing undocumented coercions.

IN returns native Boolean -1/0, with NULL when no candidate matches and a NULL
candidate is present. A successful match remains True despite another NULL
candidate. A NULL tested value returns NULL. NOT preserves that unknown value.
BETWEEN includes both bounds and returns NULL when any operand is NULL.
Reversed bounds explicitly refuse3251 until their Access behavior is verified;
no undocumented ordering rule is substituted. Lists are bounded to128 operands.
Nested SELECT/UNION, physical Decimal fields and general Jet coercion remain open.

SELECT projections and WHERE predicates retain proven Boolean metadata; Decimal
projection aliases retain their exact type. Decimal LIKE requires an explicit
text conversion, including aliases, instead of exposing an internal blob to
SQLite's text matcher. Grouping and exact numeric ordering retain their existing
transport and bounds.

Microsoft's public BETWEEN contract documents inclusive bounds and NULL
propagation:
https://learn.microsoft.com/en-us/office/vba/access/concepts/miscellaneous/between-and-operator
This is documentation evidence, not a newly executed Windows Access2003 oracle.
F014/F041/F067/F068/F071 remain open. Fixed456/386/70 unchanged.

App built; existing DAO/query-model targets passed2/2 in6.50s
(6.02s/0.47s). No new CTest target or catalog ID.
