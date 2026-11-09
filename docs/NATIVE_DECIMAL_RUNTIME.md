# Native Decimal Variants and the Access Eval numeric context

CDec now produces a native Decimal Variant (VarType14, TypeName Decimal), retained
in Variant locals/arguments/arrays and unbound form controls. Its authoritative
representation is invariant decimal text backed by a checked96-bit coefficient,
separate sign and scale0..28. Internal arithmetic uses Boost.Multiprecision integer
headers; no Java, COM provider or additional Boost shared library is required.
The compiled application remains native C++. libboost-dev>=1.74 is a build-only
dependency, including headless builds; Boost Software License1.0 applies to those
headers (https://www.boost.org/LICENSE_1_0.txt).

Addition/subtraction/multiplication/division, comparisons, unary sign, Abs/Fix/Int,
Round through28places and integral conversion use decimal/integer operations.
Scale reduction rounds once from the original coefficient, with ties to even;
overflow reports6 and division by zero11. Decimal→Byte/Integer/Long/SQLite integer
conversion rounds before range checking, without a Double intermediate. Currency
conversion uses its existing exact four-place text compiler. Explicit floating
conversion still returns Single/Double. CDec floating inputs use bounded invariant
7/15-significant-digit conversion; full locale and extreme IEEE/Automation
conversion equivalence remain unproven. Text input is bounded/invariant rather
than pretending to accept all regional numeric syntax. Signed-zero/scale-bit
identity is not claimed. VBA does not acquire a new As Decimal declaration type.

Homogeneous native report Decimal Sum/Min/Max/First/Last retain Decimal values;
mixed sums and unproven Decimal average/variance paths refuse instead of silently
summing through Double. Unbound native controls retain their Decimal kind and
exact text. Explicit typed destination conversions keep their own existing
bounds. This does not enable native dbDecimal fields, exact imported Jet Decimal
storage, general SQL Decimal arithmetic/indexes, Decimal dictionary keys or full
original-format write support. Those broader requirements remain open.

The actual user reference was executed with Access11.0, DAO3.6, de-AT/de-DE.
tests/access2003_reference/expression-observations.json retains104 observed
technical expressions and sanitized provenance from archive SHA256
9a78930834b1ad811bb73ab33e4573771e29b06f7d0bff5f321c28a0680c0076.
It contains no database paths, source metadata or private VBA. The missing expressions for eight error records were recovered by exact
SHA256 case-ID matching against all112 expressions in the technical collector;
the expression-registry hash is retained. The numeric verification group has56
observations plus those8 actual errors; remaining48 locale/string/date observations
are now covered by the desktop checks documented in NATIVE_FORMAT_RUNTIME.md.

Application/VbaHost Eval now has an explicit numeric context separate from ordinary
VBA and calculated-control expressions. It uses observed Long integer literals,
Integer True/False, invariant Decimal point literals, unary/exponent precedence,
Long VarType results and Nz(Null)=Empty. Calculated controls retain their existing
Nz default. Binary errors in Eval become Error Variants until the selected result
is returned, allowing the observed IIf(True,1,1/0)=1 without changing ordinary VBA
argument evaluation. The observed Eval Switch/Error-Variant conversion reports6;
CDec remains a VBA conversion and the observed Eval context refuses it2439.
This is not a complete Jet expression AST or function-error
propagation oracle. Null bitwise certainty now requires actual zero/all-one bits;
an arbitrary nonzero known operand does not resolve unknown Null bits.

Existing core/DAO/window targets cover the actual numeric observations plus native
extreme coefficients, exact sums/division/rounding, overflow, one-time scalar
resolution, arrays, aggregates and real unbound control readback. Transcendental
and financial Double observations use an explicit relative tolerance2e-12; this
is not bit-for-bit financial/libm identity. Full F014/F041/F068/F118/F164 and
Access2003 parity remain open. Fixed catalog456/384/72 is unchanged.

Str/Str$ now retain native Decimal/Currency invariant text without an intermediate
Double and resolve scalar defaults once. Native Double→Currency conversion uses
checked integer mantissa/exponent scaling and ties-to-even, independent of the
platform long-double width. Five existing core/DAO/window/Currency targets passed
(10.03s), app built; full Currency SQL/storage/original-format scope remains open.
