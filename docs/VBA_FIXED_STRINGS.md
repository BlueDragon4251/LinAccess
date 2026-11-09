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

Scalar declarations above do not add UDTs, typed parameter/result
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


## Fixed-length array elements

Fixed and dynamic arrays now retain canonical String*n element declarations in
native storage. Public TypeName/VarType remain String()/8200; internal canonical
metadata preserves the width for typed assignments independently of that display
name. Fixed/module/static arrays and ReDim [Preserve] over typed arrays or an
existing Variant support the same bounded decimal/known-Const widths.

Every new element is initialized with n NUL characters. Element writes use the
same scalar conversion, truncation and space padding. Preserve retains prior
values/width and initializes new elements; changing a declared type or preserving
with a different element width refuses13 before replacing the previous array.
Erase deallocates dynamic storage while retaining the declared element type, or
resets fixed dimensions to NUL-filled elements. Existing last-dimension Preserve,
active For Each/ByRef, immutable Variant-copy and array-bound guards remain active.

Before allocating fixed-string elements, a16MiB payload budget bounds count*n*3:
three UTF-8 bytes are sufficient for every UTF-16 unit of a supported value.
This protects later writes as well as initial NUL data. The existing one-million
array-element/rank limits remain. Oversized declarations fail3251 before procedure
body effects. This is a native resource bound, not an Access allocation-limit claim.

This adds no UDT layout, random-access file records, complete typed parameter/
reference coercion or Windows array interoperability oracle. Native Variant copies
retain element declarations; general Access interop/copy semantics remain open.

The existing scalar fixture was moved from the large core_tests.cpp into
vba_fixed_string_tests.cpp within the same CTest target to reduce subsequent build
costs. It also exercises actual array VBA, NUL/default values, preserved/new values,
String identity, copies, For Each, module/static state, Erase, typed Variant ReDim,
failed width changes, failed Null writes and preallocation memory refusal.
