# Checked Single and Double VBA values

Module fields, procedure locals, parameters, results and defaults now support
checked Single/Double storage. Single retains a distinct numeric tag and rounds
to native IEEE 32-bit float before storage; Double retains the existing 64-bit
representation. VarType/TypeName report 4/Single and 5/Double.

CSng is executed natively. CInt/CLng/CSng/CDbl/CBool share one-argument dispatch,
one-time object scalar resolution and checked conversion. Numeric conversion
accepts surrounding ASCII whitespace. Single overflow raises 6 before assignment
or procedure body effects. Null/error/invalid-text conversion retains native
94/13 handling. Defaults preserve their declared kind and known ByRef subtype
conflicts reject with 13. Matching references and parenthesized temporary values
retain their existing bounded semantics. Single text conversion uses seven
significant digits, avoiding generic Double tails for common values such as 0.1.
Exact Windows formatting, notation, locale and all edge cases remain open.

Single nonzero inputs below the smallest representable subnormal are explicitly
diagnosed 3251 pending Access validation. Single For counters are explicitly
unsupported before procedure effects rather than silently using Double counter
semantics; bounded Double fractional loops retain their existing execution.
Full Single loop rounding/termination, underflow boundaries, numeric literals,
arithmetic promotion/intermediate overflow, Byte/Date/Decimal and complete
Variant/reference/locale/compiler/error timing semantics remain open. Len storage
semantics and complete DAO/SQL Single coercion are not added.

Core tests cover 0.1 rounding, 16777217 -> 16777216, distinct Double identity,
Boolean/Empty input, object evaluation count, formatting, argument errors,
overflow and explicit underflow diagnostics. Native Qt/VBA/DAO tests store
Single=16777216 and Double=16777217 through actual parameterized SQL, verify
typed defaults/fields/results, unchanged rows after bad input, overflow value
preservation, ByRef rejection/matching/temporary calls and Double loops.
Targeted tests: 4/4 (42.76 s). Full validation is in PROGRESS.md.
Original application execution coverage is unchanged.

Reference: https://learn.microsoft.com/en-us/office/vba/language/concepts/getting-started/type-conversion-functions
