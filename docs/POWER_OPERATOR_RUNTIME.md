# Native VBA exponentiation

VBA and native Access expressions now execute the caret power operator. Power
binds above unary signs, uses Double values/results and evaluates chained powers
from left to right, matching the Microsoft VBA example 3^3^3 = 19683. Prefix
parsing now respects the enclosing operand precedence, keeping signed right-hand
operands separate from subsequent powers while retaining parenthesized groups.
Core tests include -2^2, (-2)^2, 2^-2^3 and mixed multiplication.

Null propagates, Empty is zero and numeric text/object scalar defaults use the
shared checked conversion. A negative base requires an integer exponent (error
5); nonfinite results reject with 6. Zero base with a negative exponent and
nonzero-base results that completely underflow to zero are explicitly diagnosed
3251 pending Access validation. Full libm boundary/rounding, locale, exact Windows
error identity and all numeric promotion/literal rules remain open. SQL caret
rewriting is not added; this is native VBA/Access-expression execution.

Core tests cover the documented left-to-right example, signs, parentheses,
negative exponents, Null/Empty, Double result identity, errors and one-time object
reads. The fixture-independent native Qt form regression now calculates powers,
reads them through VBA and stores actual SQL audit rows before/after a pending
edit. Domain/overflow errors prevent subsequent SQL effects. Targeted core/GUI
tests: 2/2 (1.19 s). Full validation is in PROGRESS.md. Original application
business execution coverage is unchanged.

References:
- https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/caret-operator
- https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/operator-precedence
