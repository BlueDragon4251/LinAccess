# Native Currency execution

Currency Variants use a signed 64-bit integer scaled by 10,000. CCur preserves
that subtype through values, arguments, returns and unbound form controls.
Decimal strings are parsed without passing through Double; rounding uses ties
to even. Both endpoints, discarded decimal digits, exponent notation, overflow,
Null (94), type mismatch (13) and argument count (450) have executable tests.
Text conversion preserves all four fractional places when required.

Currency/Currency addition, subtraction and multiplication use checked wide
integer intermediates. Multiplication rounds back to four places. Comparisons
between two Currency values use their scaled integers. Unary signs, Abs and Int
preserve Currency; overflow is reported rather than wrapping. Division returns
Double. VarType identifies Currency as 6; TypeName reports Currency. General
native object class-name introspection is explicitly unsupported, except Nothing.

The unchanged original complex form contains one CCur calculated expression.
The regression sets a synthetic non-Null value in its actual original input
control and verifies a Currency result, then restores Null. In the initial Null
state, eager IIf argument evaluation reaches CCur(Null), producing error 94.
This remaining calculated error is inventoried explicitly. There is no Microsoft
Access execution oracle proving how that original control renders in Access.
The implementation does not change the original expression to hide this error.

## Remaining boundaries

- Conversion currently accepts invariant decimal syntax, not every Access locale,
  currency symbol or grouping separator.
- Mixed arithmetic/coercion combinations, typed declarations, Currency literal
  suffixes and Decimal/Date subtypes do not yet implement the full VBA rules.
- Explicit conversion to Double can lose precision, as expected for that type.
  Currency/Currency calculations and string conversion avoid that conversion.
- Original Jet MONEY columns still use the existing SQLite REAL import. Their
  storage, aggregate accumulation and DAO/SQL parameter/scalar bridges are not
  an end-to-end exact Currency implementation. The import checklist stays open.
- These tests establish bounded native execution, not complete VBA compatibility.

The runtime is C++20; wide integer intermediates are supported by the target
Debian/Ubuntu GCC toolchains. No Python or COM service performs these operations.

## Nested ByRef regression

A reproduced defect in reference synchronization compared Variant kind and Double
payload but omitted the Currency scaled-integer payload. A nested ByRef procedure
could therefore lose an exact 0.0001 mutation while appearing to complete. The
new regression failed on the unchanged runtime before the comparison was fixed.

The comparison now includes scaled units. Core tests verify an exact unit increment,
ByVal isolation, two parameters referencing the same caller variable, and mutations
retained before an error handled by the caller. The native corpus test also passes
the actual Currency result of the unchanged original calculated expression through
a separate generic, nested native VBA helper (12500 to 12501 scaled units). Original
code bytes remain unchanged; this tests original-expression value interoperability,
not an original procedure calling that synthetic helper. Exact imported/SQL Money
storage, declared Currency variable/parameter coercion and general typed VBA remain
outside this fix.

ByRef-fix validation: all 18 CTests passed (59.57 seconds), the GUI-disabled core
build/test passed, a fresh Debian package matched the current executable Build ID
and passed content/permission checks, and all four fixture SHA-256 checks remained
unchanged.
