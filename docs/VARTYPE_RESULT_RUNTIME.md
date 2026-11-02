# Integer result of native VarType

Microsoft specifies an Integer result for VarType. Previously native VarType
returned the correct supported type-code number in a Double, so nested VarType
and TypeName produced the wrong observable type. The result now uses the checked
Integer value subtype, including the existing Byte-array code 8209.

Before/after core and native Qt/VBA/SQL regressions reproduce both failures
against the previous implementation (0/2) and pass after the change (2/2, 1.14 s).
Tests cover nested type inspection, Null/Error codes, Byte-array code/subtype,
an actual calculated form field displaying Integer, a Variant assignment
retaining the result subtype and SQL storage of the observed nested type code.
The existing no-premature-source-save regression remains active.

This changes the return subtype, not full VarType semantics. General object
class/default-property inference, arrays beyond supported Byte arrays, Date/
Decimal, complete declared-variable storage and all Windows boundaries remain
open. No additional original business workflow is claimed. Full validation is
in PROGRESS.md.

Reference:
https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/vartype-function
