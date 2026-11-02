# Native optional parameters with explicit defaults

VBA Sub/Function declarations accept trailing Optional ByVal/ByRef parameters
with explicit numeric, Boolean, String, Empty or Variant Null literals. Omitted
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

Optional parameters without explicit defaults, Missing/IsMissing, positional
argument holes, named arguments, user-defined constant defaults, Object defaults,
ParamArray and full VBA declaration/compile-time coercion semantics remain open.
This segment does not claim complete VBA argument binding or a Windows oracle.
