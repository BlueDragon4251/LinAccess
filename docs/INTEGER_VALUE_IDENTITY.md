# Integer and Long value identity

Native numeric values now retain an explicit Integer/Long tag for checked module
fields, procedure locals/parameters/results and integral loop assignments. Default
values retain their declared type. Variant copies preserve the tag. Existing
finite 16/32-bit range checks and ties-to-even assignment rounding remain.

CInt performs checked signed 16-bit conversion with ties-to-even rounding and
one-time object scalar resolution; CLng returns a tagged Long. VarType reports
2/3 and TypeName reports Integer/Long. Standard VbVarType constants are available.
Len storage semantics remain open: Variant storage differs from declared numeric
storage, and a value tag alone does not establish that distinction. The former
String-conversion behavior is retained pending a complete storage-aware design.
Known references now also reject numeric subtype changes rather than refreshing
a checked parameter from an incompatible reference representation. Alias change
comparison includes the subtype tag.

This is a storage/conversion foundation, not complete typed VBA arithmetic.
Literals and general numeric host/DAO values still use the existing Double
representation. General arithmetic/unary operations still produce generic Double
results unless a checked assignment converts them; promotion, typed overflow,
Single/Byte/Date/Decimal identity and full Variant/ByRef/error timing remain open.
In particular this does not claim Windows-equivalent type identity for numeric
literals or all DAO fields. Complete locale parsing and compiler diagnostics
remain open. SQL CInt/CLng are not added.

Len reference: https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/len-function

Validation and actual native SQL regression results are recorded in PROGRESS.md.
