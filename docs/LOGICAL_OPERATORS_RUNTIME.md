# Native Xor, Eqv and Imp

VBA and native Access expressions now execute logical exclusion, equivalence and
implication. Logical precedence follows Not, And, Or, Xor, Eqv, Imp; arithmetic,
comparison and concatenation retain their documented higher levels. Existing
power/prefix precedence remains covered by the core/native arithmetic regressions.

Boolean pairs return Boolean values. Xor/Eqv propagate Null; Imp implements the
documented Boolean/Null table. For bounded numeric operands, checked Long
conversion precedes bitwise xor/complement/or operations. Known zero antecedents
and all-one consequents determine true even with a Null counterpart. Two tagged
Integer operands retain Integer; other supported numeric combinations return Long.
Object scalar defaults resolve once; invalid text/range input rejects with 13/6.
Full mixed-type/Variant/error timing/Byte promotion and existing And/Or numeric
Null/result-type edge semantics remain open; this is not complete Boolean/VBA
type-system parity. SQL Xor/Eqv/Imp rewriting is not added.

Core tests cover Boolean truth tables, all nine Imp Boolean/Null combinations,
public bitwise examples (10 Xor 8 = 2, 10 Eqv 8 = -3, 8 Imp 10 = -1), precedence,
subtypes, rounding, errors and default evaluation count. The fixture-independent
Qt/VBA/SQL form test calculates a comparison/Xor control before/after a pending
edit and persists true=-1/false=0. Native VBA also stores real bitwise SQL audit
values and proves range errors prevent later writes. Targeted tests: 2/2
(0.97 s). Full validation is in PROGRESS.md. Original application business
execution coverage is unchanged.

References:
- https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/xor-operator
- https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/eqv-operator
- https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/imp-operator
- https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/operator-precedence
