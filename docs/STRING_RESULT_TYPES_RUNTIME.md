# Native string-function result types

Supported Len string measurements and InStr positions now return tagged VBA Long
values; StrComp returns tagged Integer (-1/0/1). Previously their numeric values
were returned as Double, making VarType/TypeName and subsequent subtype-sensitive
expressions incorrect. Null propagation is preserved. String length/position
conversion checks the signed 32-bit result range instead of silently narrowing.

Core regression reproduced the type mismatch before the fix. Tests cover UTF-16
supplementary-character positions, found/not-found/empty-needle results including
maximum Long start, comparison signs, nested VarType/TypeName and Null.
string_result_types_documents checks actual bound/calculated Qt controls, DAO
QueryDef parameter/SQL values and type codes, pending accepted Value versus Text,
explicit save, Null draft and Undo without unwanted source changes. Targeted
core/native tests pass 2/2 (1.04 s); GUI suite has 36 CTests. Full validation is
recorded in PROGRESS.md.

This fixes result tags in the existing bounded string paths. Len storage sizes
for declared numeric/user-defined types, complete locale/LCID/Jet comparisons,
BSTR/byte-string operations, computed SQL/DAO subtype propagation and exact
Windows Access error/event identities remain open. No additional complete
original business workflow or blanket string-function compatibility is claimed.

References:
https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/len-function
https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/instr-function
https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/strcomp-function
