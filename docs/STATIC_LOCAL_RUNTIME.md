# Explicit procedure-local Static storage

Explicit Static declarations inside a supported Sub/Function create persistent
local bindings in that host's module state, keyed by normalized procedure and
variable names. They remain local: module/public lookup does not expose them.
Supported scalar/native object and existing bounded dynamic/fixed array
declarations retain their values, identities, allocation/bounds/type and fixed
array guards across calls. Explicit Set Nothing and existing Erase semantics
operate on the retained binding. Declaration defaults are native values; As New
and arbitrary runtime initializers are not implemented.

The shared alias mechanism publishes changes between statements and refreshes
state across nested calls, including recursion. Different procedures and hosts
remain independent. Exact module source identity separates edited-source state,
as with existing module fields. Host/module lifetime owns the storage; this
does not implement an Access IDE Reset/End or a general COM lifecycle. Native
procedure translation plans retain only declaration metadata, never live static
values. Current bounded fixed-array declaration initialization may allocate a
temporary default array before an existing static binding is restored.

The existing core target covers repeated counters, Dictionary identity/state,
fixed/dynamic array values, different procedure/host/source scopes and recursion.
No new CTest target is added.

Static Sub/Function and Property modifiers now use the same retained local
storage. Initial Static (after visibility) and trailing Static are supported;
using both is refused before effects. Every nonconstant Dim in such a procedure
is persistent, including declarations in inactive/inline branches. Parameters,
ParamArray, constants, Err and function/property result variables remain local
to each invocation. MS-VBAL5.3.1.2/5.3.1/5.4.3.1 define the lifetime distinction.

With no Option Explicit, actual unresolved reads/assignments create procedure-
private implicit Variant bindings after normal local/module/procedure/instance
lookup. They share the existing alias synchronization, including ByRef mutation,
recursion and error exits. At most4096 implicit/static entries can be extended
through this implicit path. Ordinary nonstatic procedures retain their existing
strict unresolved-name behavior. Option Explicit prevents implicit creation;
this is a native runtime guard, not a complete Access compile-time diagnostic.

Def-type directives are refused before body effects for static procedures.
Implicit type suffixes require explicit native declarations and fail3251 rather
than inventing a Variant subtype. As New, complete declaration/compiler binding,
full reset/End/class lifecycle and a Windows Access2003 oracle remain open.
No unknown class, host API or object member is fabricated by implicit binding.

The application was rebuilt. Existing core/GUI scope passed after correcting an
invalid null owner in the technical GUI host: core3.02s and GUI16.68s. The later
Def-type/suffix guards were checked only in the affected existing core target
(1/1,2.95s); GUI was not repeated for those guards. Actual cases cover fixed/
dynamic arrays, fixed strings, Dictionary identity/state, implicitly created
ByRef variables, recursion, fresh parameters/results, error retention, static
Get/Let, independent real form instances/reopening and shared project state
across fresh standard-module hosts. No new CTest ID. All32 original VBA artifacts
were inventoried: no static procedure candidate; no new original or Windows run.
Catalog remains456/386/70, with F042/F043 still open in full.


## Historical proof before procedure modifiers (2026-11-04)

STATIC_LOCAL_RUNTIME.md specifies actual per-host/procedure/source static
scalar/object/array state, nested-call alias refresh and current initialization
limits. Existing core cases cover persistence, native Dictionary/fixed/dynamic
arrays, independent scopes and recursion. Static procedure modifiers now refuse
incomplete implicit-local behavior instead of silently executing transient
locals. Full desktop 48/48 (79.78 s), freshly rebuilt GUI-free 1/1 (1.45 s),
targeted core 1/1 (1.26 s), Debian payload/permissions/Build-ID/pinned SQLite and
all four original hashes pass. No new CTest target. Scoped checklist 372 checked
/ 72 open, not complete Access 2003 parity.
