# Native object array bindings

The existing closed native object vocabulary is accepted for fixed/dynamic local,
Static and module arrays, ReDim/Preserve, exact ByRef array parameters and Function/
Property Get results. Qualified names include DAO, Access, Scripting and VBA types.
Slots start as Nothing. Array copying retains object identity while preserving the
existing array value semantics. TypeName uses canonical native class names and
VarType is 8201. Erase retains the existing fixed/dynamic distinction.

Set element assignments enforce the actual declared native interface through
isNativeType, rather than guessing from an object's display class name. This also
strengthens scalar DAO/Collection bindings: an ordinary project class named Database
cannot impersonate DAO.Database. Existing real native objects and Nothing remain
accepted. ErrObject scalar/array declarations refer to the actual error object.
Local Forms/Reports bindings now take precedence over global collections in VBA
expressions; calculated Access-expression collection lookup is unchanged.

Validation extends existing core, GUI and DAO targets without adding CTest IDs.
Technical fixtures exercise actual Collection/Dictionary objects, fixed/dynamic
slots, ByRef resizing, returned arrays, type metadata and shared Err references;
real form/control arrays with typed writes and retained identity after rejected
Set; real SQLite-backed DAO database/recordset/field arrays, cursor movement,
returned database arrays and error3420 after closing a retained recordset.
A real project class named Database is rejected as a DAO array element.

Initial checks exposed a stale unsupported-Dictionary-array expectation and a real
local Forms lookup collision. After correcting those, affected core/GUI2/2 passed
24.46s (4.85/19.61); DAO passed2.83s before the lookup correction. App built.
No full-suite, Original-event, Windows-oracle or package refresh is claimed here.
Unknown libraries/interfaces, general COM/type-library binding, UDTs, AsNew arrays,
complete lifetime/GC semantics and complete compiler parity remain open.
Catalog456/386/70 is unchanged; complete F042/F043 remain open.
