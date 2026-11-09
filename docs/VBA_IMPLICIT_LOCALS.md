# Native implicit procedure-local variables

Without Option Explicit, supported native procedure execution resolves existing
local/parameter/constant, module/public project and host instance names first.
Only a genuinely unresolved simple variable name creates a fresh Empty Variant
local. Supported external public functions are invoked; a Sub used as a value or
assignment to a procedure name fails instead of becoming an invented variable.
This follows MS-VBAL5.2.1.3 and5.6.10 within the native namespace boundaries.

Ordinary procedures keep these values only for that invocation. Static procedures
use the existing retained module/procedure slots (STATIC_LOCAL_RUNTIME.md).
Actual ByRef mutation, arrays, objects and loop counters use the normal runtime
bindings. These are real VM locals; no inferred UI control, class, API, SQL parameter
or external reference is manufactured. Existing declared bindings retain priority.
A real unbound TextBox test confirms assignment reaches the control rather than a
new local. New ordinary invocations start again with Empty; module fields persist.

Option Explicit prevents implicit creation on reads and writes. Its current native
error5 is a runtime guard, not full Access compile-time binding/diagnostics.
Def-type directives are refused before procedure effects, and implicit type
suffixes require explicit native declarations (3251). There are at most4096 new
implicit bindings per frame; the retained static path keeps its existing budget.
Complete reference/class/Unicode/default-type/compiler semantics remain open.
Standalone expression evaluation, calculated expressions and SQL parameter binding
do not acquire this procedure-local fallback.

## Verification

Application built. Existing GUI passed16.32s. The affected core/original retry
passed2/2 (core3.11s; real_kassa_dao_event124.12s), after correcting only technical
fixtures: case-insensitive external-name lookup in the core host, and Option
Explicit in the MsgBox fixture that deliberately expects an undeclared-name error.
No GUI retry or new CTest target. Original VBA/source/data copies were not edited.
Original inventory:32 VBA artifacts,19 with Option Explicit,13 without,0 Def-type
candidates. The actual unchanged original event regression still reaches the known
missing-macro boundary2485/1023; this is not a complete positive sale workflow.
No new Windows Access oracle, full-suite or package proof for this extension.
F037/F042/F043 and the fixed catalog456/386/70 retain their existing statuses.
