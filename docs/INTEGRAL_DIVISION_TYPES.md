# Native primitive integer division and Mod types

For supported native integral operands, \ and Mod now retain Byte for Byte/Byte,
Integer for small integral mixtures (including Boolean/Empty), and Long for the
wider supported integral paths. Their results carry actual proven declaration
categories into later arithmetic. Fixed signed-minimum division overflow remains
6; a proven Variant Integer can widen to Long and Variant Long to Double.
Operands still use the existing checked conversion helper; calculation uses
int64 for the supported32-bit operands, avoiding signed32-bit division overflow.

The existing truncation toward zero, signed remainder, zero-divisor11 and Null
behavior remain. In particular, the general MS-VBAL Mod wording does not replace
the bounded negative-remainder behavior with an unverified absolute value.
Mixed floating/String/Currency/Date/Decimal conversion rules retain the previous
bounded Long path rather than inventing complete operand/range equivalence.
Access Eval retains its separate prior path. Full AST, suffix/64-bit, locale,
all mixed coercions and Windows parity remain open; totals remain456/386/70.

Application/core/DAO targets built; the first existing bundle passed2/2 in8.72s
(core5.55/DAO3.17). Actual tests cover subtype combinations, quotient/remainder,
Variant widening, fixed overflow, result use in later arithmetic, signed Mod
and unchanged Null propagation. A real native Byte function verifies the Byte
quotient inside its body and returns a Mod result through SQL/DAO with actual
Type2/Size1/Byte value1. No new CTest targets/catalog IDs.
The checked4cd0f848 DEB predates the later logic and integral-division scopes.
