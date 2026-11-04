# Native optional parameters with explicit defaults

VBA Sub/Function declarations accept trailing Optional ByVal/ByRef parameters
with scalar literal/constant-expression defaults, including prior same-module
constants. Defaults fold without runtime host calls; omitted trailing arguments
copy and coerce the folded values for each call.
Supplied arguments retain checked type conversion and existing ByRef aliasing.
No supplied argument slot or reference is accessed for an omitted parameter.

The existing core and DAO execution targets exercise defaulted and explicit
function calls, a string default containing a comma, repeated missing ByRef
arguments without leaked state, supplied ByRef writeback and actual DAO writes
using a function's defaulted suffix. Too few required or too many arguments fail
450 before execution. Required parameters after Optional and runtime-dependent defaults
are rejected by procedure validation. No new CTest target is added.

## Remaining boundaries

Optional parameters without explicit defaults now bind typed zero/empty values
or a distinct Missing Variant. IsMissing distinguishes that marker from Empty,
Null and CVErr; VarType/TypeName/IsError expose its Error subtype. Function and
Sub calls support argument holes and forwarded Missing values without aliasing
defaults back to omitted caller slots. Scalar Missing coercion, exact SCODE/BSTR
semantics, qualified/cross-module constant linking, non-Nothing Object defaults, unsupported external object declarations and full declaration/
compile-time coercion remain open. Named calls are documented in NAMED_ARGUMENT_RUNTIME.md. Bounded
ParamArray evaluation is documented separately in [PARAMARRAY_RUNTIME.md](PARAMARRAY_RUNTIME.md).
This segment does not claim complete VBA argument binding or a Windows oracle.


## Pure default expressions (2026-11-04)

VBA_CONSTANTS_RUNTIME.md defines the folded default-expression path. Arithmetic,
string/date values and previous same-module constants are retained in immutable
parameter metadata. Named overrides and explicit Empty remain separate from
omission; each omitted ByRef argument receives an independent value. Existing
Null/Empty/Collection Nothing defaults preserve their prior meanings. Runtime
calls/variables and bare clocks do not execute during default evaluation.


## Native Object bindings and Nothing defaults (2026-11-04)

As Object parameters now require actual native object references (including
Nothing), rejecting scalar arguments with type mismatch 13 before the body.
ByVal preserves shared object mutation while keeping reference replacement
local; ByRef replacement writes back through supported caller bindings.
Optional As Object with an omitted default or explicit Nothing binds a fresh
Nothing slot for each call, without aliasing omitted arguments. IsMissing
remains false for these typed slots. Non-Nothing object defaults and optional
external types remain unsupported. Supported native DAO Nothing bindings
are described in DAO_OBJECT_BINDINGS.md.

Set targets use checked binding. Object-collection For Each control assignments
use the same checked path, refusing scalar values for As Object controls.
General default-member Let/Set coercion, full compile-time declaration checks,
class/reference-library binding and Access 2003 error precedence remain open.
An existing core case now expects parameter mismatch 13 before a numeric
argument can reach the object's loop, replacing its previous late error 424.


## Checked native Object parameters and optional Nothing (2026-11-04)

As Object binds actual references/Nothing before execution. ByVal shares object
mutation while replacement stays local; supported ByRef writes replacement back.
Optional Object Nothing slots are independent per omitted call. Set and object
collection loops use checked assignment. OPTIONAL_ARGUMENT_RUNTIME.md records
remaining default-member/compiler/early-bound boundaries. Existing core target
extended only. Full desktop 48/48 (77.73 s), freshly rebuilt GUI-free 1/1
(1.24 s), targeted core 1/1 (1.23 s), Debian payload/permissions/Build-ID/pinned
SQLite and all four original hashes pass. Scoped checklist 370 checked / 72
open; not complete Access 2003 parity.
