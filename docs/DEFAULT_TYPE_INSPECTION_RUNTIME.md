# VarType of native object default values

Native scalar-default providers now advertise that capability explicitly.
VarType follows those defaults rather than reporting Object for every reference.
The supported control Value, DAO Field/Parameter and Err providers participate;
objects without an advertised scalar default and Nothing retain Object type.
Getter errors propagate unchanged. Nested default references have a 16-read
limit (error 28); each default in a normal chain is evaluated once. TypeName
class-name inference is not changed. Err.Number/default Number now retains Long.

Core regression covers Byte/text/Null defaults, non-default objects, Nothing,
nested/recursive defaults, a failing getter (13) and Err/default Number type.
The fixture-independent default_type_documents Qt/VBA/DAO/SQL test takes real
Control, Field and Parameter references and stores the observed type codes
8/1/17/8/8/3 in SQL. It distinguishes Null from live EOF (3021), closed field/
parameter access (3420), closed Fields.Count and a rejected closed-field write.
The original test table remains unchanged and the native form stays clean.

This regression exposed generic exceptions on closed recordset data access;
those existing guards now use the bounded DAO 3420 code consistently with the
other closed-object guards. Full DAO property/lifetime/error identity remains
open; this is not a Windows Access oracle.

Before the implementation, independent core/native default-type tests fail 0/2.
After it and the closed-recordset correction, targeted core/native/DAO tests pass
3/3 (1.38 s). GUI suite now has 29 CTests. Full validation is in PROGRESS.md.
No extra original business execution is claimed.

This is limited to native advertised scalar-default providers. General COM/
class/default-property binding, non-Value control defaults, full TypeName class
metadata, exact imported numeric/Date/Decimal field subtypes and all Variant/
array/Windows boundaries remain open. No fallback turns a failed getter into a
successful guessed type. No dependency on Windows/Wine is introduced.

References:
- https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/vartype-function
- https://learn.microsoft.com/en-us/office/client-developer/access/desktop-database-reference/recordset-close-method-dao
