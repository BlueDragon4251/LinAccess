# Checked Integer/Long procedure values

Procedure declarations now retain parameter type names through execution. Local
Dim Integer/Long fields, numeric parameters and function-result assignments use
the same checked 16/32-bit rounded conversion as module fields. Defaults remain
zero. Failed writes leave their previous values intact. ByVal inputs convert
before body execution without changing the caller argument; Null/type/overflow
failures use 94/13/6 and return an error result rather than escaping as C++ errors.

Known numeric ByRef declarations carry reference type metadata. Conflicting
known non-Variant widths/types reject with 13 before callee effects. Numeric
references also reject a conversion that would change their initial value
representation. Matching references remain live and checked; parenthesized
argument expressions create temporaries, preserving the caller value. The
existing alias/error propagation and copyback paths remain in use. Variant,
object-default, external-reference and mutating-alias type identity is not fully
established; this is not a complete VBA ByRef compiler contract.

Integer/Long For counters now use checked writes, including Next. Their supported
bounds and step are integral numeric values checked against the counter width;
fractional typed limits/steps explicitly reject with 3251 instead of inventing
VBA coercion semantics. Untyped/other counter paths retain prior scope. Known
numeric For Each counters and Set-object targets reject before procedure body
effects (3251), rather than accepting an impossible scalar/object assignment.
Full compile diagnostics, all loop coercion and boundary Resume/handler behavior
remain open.

Core tests cover local rounding/write preservation, checked function results,
ByVal conversion, matching/mismatched ByRef behavior, temporary arguments,
copyback overflow, typed integral loops, overflow/fractional guards and invalid
object/enumerator targets. Actual native GUI/VBA/SQL tests persist a value derived
from rounded input/local/reference/result stages; invalid arguments and references
perform no SQL effects. A temporary reference permits the intended callee effect
without changing its caller. Targeted core/main-window/layout/original-heavy set:
4/4 (43.64 s). Full validation is in PROGRESS.md. Original business execution
coverage is unchanged.

Values still use the existing Number representation. VarType/TypeName/Len declared
subtypes, typed intermediate arithmetic/promotion/literals, other primitive local
and parameter types, complete ByRef/loop/type-error timing, arrays/classes and a
full VBA AST remain open. These checked storage/assignment paths do not establish
complete VBA numeric parity.
