# Native VBA CBool

CBool now returns an actual Boolean value through the native expression runtime.
Numeric zero and Empty are false; finite nonzero numbers, including Currency,
are true. Case-insensitive True/False strings and numeric strings with surrounding
spaces convert through the shared checked Boolean conversion. The same numeric
text whitespace handling applies to typed Boolean variables and parameters.

Null raises 94, Error variants and invalid/empty text raise 13, Nothing raises
91, and wrong arity raises 450. Object scalar defaults resolve once. Core tests
check kinds, VarType/TypeName Boolean identity, Currency, numeric/text values,
errors and single evaluation. Native VBA/DAO tests persist CBool(False text) as
zero through a BIT parameter, verify true=-1 and caught conversion errors with
no unintended SQL changes. Targeted tests: 2/2 (1.18 s).

This does not implement CInt, full numeric subtypes, Date/locale/LCID conversion,
SQL CBool, full object-default recursion or exact Access error descriptions.
Original application execution coverage is unchanged. Full platform validation
is recorded in PROGRESS.md.
