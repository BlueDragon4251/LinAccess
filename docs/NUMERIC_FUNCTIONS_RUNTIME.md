# Native Abs, Int, Fix and Sgn

Native VBA/Access expressions now execute Fix (truncate toward zero) and Sgn
(Integer sign result). Abs and Int now propagate Null as documented, alongside
Fix. Sgn Null rejects with 94. All four functions validate one argument (450),
resolve object scalar defaults once and use checked numeric input conversion.

Known Integer/Long/Single/Double values retain their supported numeric subtype
through Abs/Int/Fix. Abs of the signed Integer/Long minimum rejects with 6 rather
than widening silently. Currency retains exact fixed-point operations: Abs and
Int reuse checked Currency functions; Fix truncates scaled units without passing
through Double, including the lower Currency boundary. Boolean/Empty inputs use
bounded Integer conversion; numeric text uses the existing checked Double path.
Full Variant/Byte/Date/locale/formatting/promotion/error-timing identity remains
open. SQL function rewriting is not added.

Core tests verify Null behavior, negative floor versus truncation, sign/subtype
results, Currency boundary precision, arity/range/error handling and one-time
object evaluation. A second actual Qt form in the fixture-independent arithmetic
CTest calculates Abs/Int/Fix/Sgn for -3.75 and writes the native values
3.75/-4/-3/-1 through VBA into SQL. A Long Abs overflow prevents the next write;
reading calculated values leaves the record clean. Targeted tests: 2/2 (1.02 s).
Full validation is recorded in PROGRESS.md. Original business execution coverage
is unchanged.

References:
- https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/abs-function
- https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/int-fix-functions
- https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/sgn-function
