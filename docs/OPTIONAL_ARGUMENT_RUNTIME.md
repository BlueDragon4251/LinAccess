# Native optional parameters with explicit defaults

VBA Sub/Function declarations accept trailing Optional ByVal/ByRef parameters
with numeric, Boolean, String, Empty or Variant Null literal defaults. Omitted
trailing arguments evaluate and coerce their declared defaults for each call.
Supplied arguments retain checked type conversion and existing ByRef aliasing.
No supplied argument slot or reference is accessed for an omitted parameter.

The existing core and DAO execution targets exercise defaulted and explicit
function calls, a string default containing a comma, repeated missing ByRef
arguments without leaked state, supplied ByRef writeback and actual DAO writes
using a function's defaulted suffix. Too few required or too many arguments fail
450 before execution. Required parameters after Optional and nonliteral defaults
are rejected by procedure validation. No new CTest target is added.

## Remaining boundaries

Optional parameters without explicit defaults now bind typed zero/empty values
or a distinct Missing Variant. IsMissing distinguishes that marker from Empty,
Null and CVErr; VarType/TypeName/IsError expose its Error subtype. Function and
Sub calls support argument holes and forwarded Missing values without aliasing
defaults back to omitted caller slots. Scalar Missing coercion, exact SCODE/BSTR
semantics, user-defined constant defaults, Object defaults and full declaration/
compile-time coercion remain open. Named calls are documented in NAMED_ARGUMENT_RUNTIME.md. Bounded
ParamArray evaluation is documented separately in [PARAMARRAY_RUNTIME.md](PARAMARRAY_RUNTIME.md).
This segment does not claim complete VBA argument binding or a Windows oracle.
