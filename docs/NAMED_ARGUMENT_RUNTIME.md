# Native named argument binding

Native VBA Sub/Function calls accept parameter:=expression in expressions,
Call statements and implicit procedure calls. Binding uses declared parameter
names, preserves written expression evaluation order and permutes both values
and ByRef references. Optional defaults/Missing slots retain their existing
checked binding. Unknown names fail 448; duplicates, missing required arguments
and positional arguments after named ones fail 450 before the callee body.

Native DoCmd methods advertise parameter signatures. OpenForm/OpenReport,
Close, GoToRecord, RunSQL, RunCommand, GoToControl, Requery, SelectObject and
DoMenuItem map named values into their existing real document/action paths.
Close distinguishes omitted ObjectType/ObjectName from explicit numeric values.
A native GUI regression executes SQL, opens a filtered child, reads its bound
control and closes it through reordered named arguments.

Unsupported method options remain unsupported after binding. RunSQL's
UseTransaction option now has a native scoped-savepoint implementation; see
[transaction scope](DOCMD_RUNSQL_TRANSACTION_RUNTIME.md). Named built-in functions, Access Eval
named calls, named ParamArray calls, arbitrary DAO/COM method signatures, full
compile-time evaluation order and Windows binder equivalence remain open.
Missing scalar/error coercion still uses the bounded Error-Variant behavior,
not a claim of complete SCODE/BSTR semantics.
