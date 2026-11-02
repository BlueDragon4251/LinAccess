# Checked VBA Byte values

Byte is now an explicit numeric value subtype (VarType 17, TypeName Byte), rather
than an accepted declaration with unchecked Double assignment. CByte, module and
local variables, zero defaults, procedure parameters and function results use a
shared checked conversion. Matching typed ByRef calls preserve the subtype;
known mismatches fail before the callee body, and parenthesized temporaries do
not modify the caller. Failed assignment preserves the previous typed value.

Numeric input rounds to the nearest integer, with ties to even, and checks
0..255 before narrowing. Null raises 94, invalid/Error values 13, range overflow
6 and invalid conversion arity 450. Scalar default properties are evaluated once.
The bounded Boolean conversion uses the unsigned Boolean bit pattern
(True -> 255, False -> 0). Abs/Int/Fix preserve an existing Byte subtype.

Core regression covers boundaries, rounding, type/default/ByRef/result/module
identity, single scalar evaluation and failure preservation. The native
byte_value_documents CTest uses actual bound/calculated Qt controls, VBA Byte
variables, DAO PARAMETERS Byte/Parameter.Type, SQL writes and explicit native
record save. It verifies stored integer 12/4, no SQL after conversion errors,
module overflow preservation and no premature save of the source record.
Targeted core/native tests pass 2/2 (0.96 s); GUI suite now has 28 CTests.

This does not implement complete VBA numeric promotion, arrays of declared Byte,
all locale/Boolean/binary-edge behavior or a Windows Access oracle. Declared Byte
For counters explicitly reject before effects until signed-step/end/overflow
semantics are validated. The independent existing SQL Byte parameter path is
not broadened to infer all VBA subtype/storage identities. The inspected 34
original module files contain no As Byte/CByte occurrences; original business
execution coverage is unchanged.

References:
- https://learn.microsoft.com/en-us/office/vba/language/concepts/getting-started/type-conversion-functions
- https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/byte-data-type
- https://learn.microsoft.com/en-us/windows/win32/api/oleauto/nf-oleauto-varui1frombool

The open-source Wine automation implementation was also inspected outside this
repository for numeric rounding and unsigned Boolean conversion behavior; no
Wine source or dependency is included in LinAccess.
