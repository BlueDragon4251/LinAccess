# Native VBA Replace

VBA and native calculated controls now execute Replace in C++, including the
three required arguments and optional start/count/compare omissions. Matching is
non-overlapping and replacement text is not rescanned. The result begins at the
one-based UTF-16 start position; it does not preserve the preceding prefix.
Empty find/count zero retain that suffix, empty replacement deletes matches,
and a start beyond the input returns an empty String.

Object scalar inputs are resolved once. Start/count/compare use checked Long
conversion, including ties-to-even rounding. Invalid start/count/comparison
raise 5; Null raises 94; required omissions/arity raise 450; existing overflow,
Error Variant and Nothing diagnostics propagate. A start splitting a surrogate
pair explicitly diagnoses 3251 because native strings currently require valid
UTF-8. Embedded NUL and intact supplementary characters are preserved.

Omitted comparison uses Binary. Explicit -1 uses the module Option Compare;
0/1/2 use the existing Binary/Text/Database host comparison adapters. Core Text
comparison remains bounded to ASCII; native Qt uses its existing case folding.
Full Windows locale/LCID/Jet collation, variable-length linguistic equivalences,
isolated-surrogate BSTR strings and exact error/evaluation identities remain
open. No dollar alias, named arguments or SQL Replace rewrite is added. SQLite's
raw Replace remains separate. DAO tests bind native VBA results into SQL rather
than claiming complete Jet SQL Replace identity.

Fixture-independent core and replace_function_documents tests cover suffixes,
counts/omissions, numeric inputs/rounding/types, scalar-once inputs, deletion,
non-overlap, NUL/Unicode, errors, module comparison and actual Qt text/calculated
controls. Pending text retains the accepted Value; explicit save refreshes the
calculated result. Programmatic replacement remains dirty until save/Undo.
Real DAO QueryDef parameters store results; rejected arguments prevent SQL side
effects. Initial targeted tests pass 2/2 (1.13 s); full validation is in
PROGRESS.md. This adds no full original business workflow or blanket VBA parity.

Reference:
https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/replace-function
