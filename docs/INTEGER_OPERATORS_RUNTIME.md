# Native VBA integer division and remainder

VBA and native Access expressions now execute integer division (backslash) and
corrected Mod. Operands convert to checked signed Long values with ties-to-even
rounding before division; quotient truncates toward zero and remainder retains
the dividend's sign. Mod no longer applies floating fmod directly. Multiplication/
division bind before integer division, which binds before Mod; addition follows.
Unary signs retain their higher precedence.

Null propagates, Empty is zero, rounded-zero denominators raise 11 and range/
quotient overflow raises 6. Numeric text and existing object scalar defaults use
the shared conversion path; core probes verify one default read per operand.
Two explicitly tagged Integer operands retain an Integer result with its range
check; other supported operand conversions yield Long. Literal types, Byte/
complete Variant promotion, all typed arithmetic and exact Windows diagnostics
remain open. SQL Mod/backslash rewriting is not added by this change.

Core tests cover fractional rounding, precedence, signed results, Null/Empty,
subtypes, invalid text, overflow, rounded-zero division and object reads. The new
fixture-independent integer_arithmetic_documents CTest constructs an actual Qt
form with bound data and calculated Mod/backslash controls, reads them through
native VBA and writes real SQL audit rows. A pending edit recalculates without
implicit record save; explicit save persists the record. Zero division and Long
overflow stop before the subsequent SQL effect. Targeted core/GUI tests: 2/2
(0.96 s). Full GUI suite now has 23 tests; validation is in PROGRESS.md.
Original business execution coverage is unchanged.

References:
- https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/mod-operator
- https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/backslash-operator
- https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/operator-precedence
