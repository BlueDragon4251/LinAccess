# Native fixed-length scalar VBA strings

Local Dim/Static and module Dim/Private/Public/Global scalar declarations accept
As String * n with1..65526 UTF-16 code units. The bounded length grammar accepts
a decimal integer or an explicitly known scalar Const name, Let-converted to
Long. Module length constants can occur after the variable declaration. Existing
local-constant ordering rules still apply. The declared width remains in native
variable type metadata; TypeName/VarType of the value remain String/8.

Initialization fills the declared width with NUL characters (U+0000). Let assignment
first uses the existing native String conversion, then truncates to width or pads
on the right with spaces (U+0020). Empty assignment therefore differs from initial
NUL storage. Module and static values retain their width across execution. Normal
assignment, LSet/RSet and the bounded existing Mid path enforce the declared width.
Mid$ assignment is now recognized explicitly and uses the same native character
replacement path; an array/object variable named Mid does not shadow Mid$.

Null assignment raises94 before changing the previous value. Invalid length/type
or unknown length constants fail before procedure effects. Native Unicode validation
is retained. Truncation that would create an isolated UTF-16 surrogate refuses3251
rather than replacing a character or claiming arbitrary UTF-16 storage support.
String conversion keeps existing native date/numeric/locale limits. Fixed strings
are scalar variable types, not new SQLite/MDB field types.

This scope does not add fixed-length string arrays/UDTs, typed parameter/result
specifications, project/reference constant lookup or complete ByRef coercion.
Existing known-ByRef-type guards apply; ByVal String receives the actual padded
value. MidB/MidB$, complex Mid targets and previous explicit Mid boundary refusals
are unchanged. General physical Jet text/encoding preservation remains open.

Existing core checks cover NUL initialization, local/module/static state, known
and forward module widths, maximum length, assignment padding/truncation, numeric
and Unicode inputs, String identity, LSet/RSet/Mid$, ByVal access, rejected Null
without mutation and preflight refusal. The initial bundle's GUI passed14.77s;
the core fixture identified missing Mid$ syntax. That specific parser gap was
fixed, followed by only the relevant core rerun:1/1 passed (2.31s). The application
was rebuilt. No new CTest, Windows/original/full-suite or new package claim.

Normative basis: MS-VBAL2.3 initialization,5.2.3.1.4 variable type declarations,
5.5.1.2.5 fixed-string Let conversion and5.4.3.5 Mid/Mid$ statements. No executed
Access 2003 result oracle is newly available. F040/F041/F042/F043/F122 stay open;
the fixed catalog remains456/386 confirmed partial scopes/70 open full requirements.
