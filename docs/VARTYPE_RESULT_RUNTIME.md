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


Native SQL scalar inspection (2026-11-10)
---------------------------------------
IsNumeric, VarType and TypeName are registered native SQL scalar functions using
the existing VBA evaluator and shared typed argument decoder. Proven Date,
Boolean, Single, bounded Currency and Byte/Integer/Long/Double declarations,
including nested conversions, reach the actual inspection operation. IsNumeric
returns BOOLEAN, VarType INTEGER2 and TypeName TEXT; DAO exposes the proven
Boolean/Integer result kinds and dbBoolean/dbInteger field types. TEXT does not
acquire a guessed Jet Text/Memo declaration.

A nullable declared field is inspected as its actual runtime value: NULL yields
IsNumeric=False, VarType=1 and TypeName="Null". An ISO TEXT source remains String;
Date is not guessed from its contents. Boolean literal binding now survives
parentheses as well as direct arguments, without treating quoted [True]/[False]
field names as literals. Supported simple projections, nested calls, predicates,
grouping and ordering use the same binding. Native numeric-text recognition uses
the current QLocale host policy, not a new Windows regional profile.

Unknown physical/literal input retains the existing conservative SQL transport:
unproven SQLite numbers are Double within the exact envelope, BLOB/object/array
inspection through this scalar transport is unsupported. This is not full Jet
literal typing, Object/UDT/type-library introspection or a Windows Access SQL
oracle. The existing Access Eval profile and its observed VarType Long return
remain separate; this uses the existing native VBA Integer contract.
