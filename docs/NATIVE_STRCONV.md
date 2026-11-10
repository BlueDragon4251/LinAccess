# Bounded native StrConv

The common VBA evaluator implements StrConv for native ASCII text/Empty with
vbUpperCase, vbLowerCase and vbProperCase. It includes documented conversion
constants, named String/Conversion/LCID argument binding, native VBA-qualified
calls and a String result. Proper-case word boundaries are exactly the documented
NUL, TAB, LF, VT, FF, CR and ASCII space; punctuation inside a word is retained.
Scalar native object defaults are read once. No host/symbol/result cache is added.

This is an explicitly bounded native function, not complete StrConv parity.
Non-ASCII/codepage/East-Asian conversion, explicit LCID, non-text and Null source
contracts refuse3251 rather than claiming unobserved Windows behavior. Null
conversion/LCID uses the existing checked numeric argument coercion. Source
business values are not substituted. All six native evaluator hosts can use the
shared function; this does not register a separate Jet SQL builtin.

App/core built; the existing core_tests bundle passed1/1 in0.52s. It covers upper/
lower/proper case, named arguments, all seven separators including embedded NUL,
apostrophe/hyphen word behavior, Empty, scalar getter count, guarded unsupported
cases and a real String-returning VBA function calling VBA.StrConv. No new CTest
or catalog identity; F041/F043/F071/F118/F164 remain broader open requirements.
Latest checked package/Original/query inventory is code299b946f; this function is
newer than that package. No new Windows or full-suite claim.

Public contracts (not an Access execution oracle): MS-VBAL6.1.1.14 and
6.1.2.11.1.37; https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/strconv-function
