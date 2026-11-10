# Native unsuffixed decimal numeric literals

MS-VBAL 3.3.2 defines unsuffixed decimal integer tokens as Integer through 32767,
Long through 2147483647 and Double above the supported Long range. Decimal-point
and exponent tokens have Double type; native exponent letters D/d and E/e are
accepted. The native evaluator now uses these types instead of assigning Double
to every unsuffixed decimal literal. This composes with the existing in-range
unary/binary arithmetic, including Single mixed with Integer versus Long.

The separately observed Access Eval integer/decimal rules remain unchanged.
D exponents are supported in native VBA, without claiming new Access Eval or
SQL grammar support. This does not implement full type-character/suffix binding,
LongLong, static-versus-Variant overflow promotion, negative-boundary constant
folding or full declared/result type parity. F040/F041/F122 stay open.

The application, core, DAO and query model targets built. DAO (3.09 seconds) and
query model (0.20) passed in the first bundle. Core exposed two stale expectations
for literal-dependent Hex/Oct and Not results: these now follow the already
implemented Integer paths. Explicit Double fixtures preserve their previous
behavior. After correcting those expectations, only core was rerun; it passed in
5.41 seconds. This is not a claim that the initial three-target bundle passed.
No new CTest IDs or catalog points were added; totals remain 456/386/70.

Later proven declaration/Variant contexts are extended in [BOUND_NUMERIC_ARITHMETIC.md](BOUND_NUMERIC_ARITHMETIC.md); its explicit unbound/full-parity limits still apply.
