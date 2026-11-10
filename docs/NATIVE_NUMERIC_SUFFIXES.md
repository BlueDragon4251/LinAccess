# Native numeric literal type characters

Native VBA decimal literals now support % (Integer), & (Long), ! (Single),
# (Double) and @ (Currency). They establish their actual numeric subtype and
proven fixed declaration category. Integer suffixes require integer digits and
retain the existing checked32-bit ranges. D/E exponents retain native support.
Currency literals go directly from decimal text to the checked fixed-point
converter; a full positive signed64 Currency boundary never passes through
Double. Single literals use direct decimal-to-float from_chars decoding, avoiding
a Double-rounding intermediate, with existing range/underflow boundaries.

The lexer mode is explicit: source VBA enables these suffixes; Access expressions
and observed Access Eval retain their prior mode. Existing adjacent numeric
concatenation is preserved where & introduces an operand. A delimited & before
an operator/end/argument boundary is the Long type character. Date delimiters,
source offsets, precedence and native class-type metadata scans remain intact.
This is not complete identifier type-character syntax, signed-minimum constant
folding, all compiler/IEEE diagnostics or new raw SQL/Eval support.

Application/core/DAO targets built; first existing bundle passed2/2 in8.61s
(core5.38/DAO3.23). It checks all five characters, bounds/errors, native exponent
and precedence, Date/concatenation preservation, exact Currency ties/boundary and
a Single decimal just above a rounding midpoint. A real native VBA function
uses all five type characters plus Optional Long=2&, and executes its default
through SQL/DAO Type5/Size8 with exact Currency units20002. No new CTest/catalog
IDs; fixed456/386/70 and the full F040/F041/F043/F122 requirements remain open.
The checked4cd0f848 DEB predates this and the other later arithmetic scopes.
