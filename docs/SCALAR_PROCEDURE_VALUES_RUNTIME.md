# Checked Boolean and String procedure values

Boolean/String declarations now participate in procedure-local assignment,
parameter entry, function-result storage and known-reference copyback checks.
They use the existing module-field scalar conversion: numeric Boolean values
become true/false, text True/False is recognized, and String targets receive text
rather than retaining a numeric/Boolean value. Empty conversions keep existing
native defaults; invalid Null/type writes use 94/13 without replacing prior data.
ByVal conversion does not mutate the caller argument vector.

Known declared references must match their parameter type. Checked reference
inputs reject conversions that would change their initial value representation
instead of silently borrowing an incompatible value. Matching Boolean/String
references remain live; parenthesized expressions are temporary values. Full
Variant/external-reference/object-default and changing-alias semantics remain
open, as do exact VBA compile diagnostics and locale/error-description identity.

The pre-effect invalid-target guard now also covers known scalar Set/For Each
variables and known Boolean/String/object For counters. Valid Variant/object
enumeration retains its previous scope. This prevents an invalid declared target
from executing earlier procedure effects, but does not establish a full VBA AST
or a general compiler/type checker.

Core tests verify False text, numeric/String conversion, typed return kinds,
preserved values after Null errors, ByVal input, live references and invalid
scalar loop targets. Actual native GUI/VBA/DAO tests bind converted Boolean/String
values through a temporary parameterized QueryDef, store false=0/true=-1, preserve
Unicode and apostrophes, and prove invalid inputs/references have no SQL effects.
Temporary references permit the intended callee effect while retaining the caller
String. Targeted core/main-window/layout/original-heavy set: 4/4 (43.12 s). Full
validation is in PROGRESS.md; original business execution coverage is unchanged.

The synthetic test initially used the unsupported SQL parameter alias BIT; it
now uses the existing BOOLEAN spelling. BIT and other Boolean SQL aliases were
a separate parser gap at this stage; subsequent support is documented in
SQL_BOOLEAN_PARAMETERS.md, not implied by typed VBA values.
Other primitive types, numeric subtypes/promotion/literals, full ByRef/loop/error
semantics, fixed-length Strings, arrays/classes and complete VBA compatibility
remain open.
