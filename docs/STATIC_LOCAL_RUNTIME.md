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

Static Sub/Function modifiers require additional implicit-local/result lifetime
rules and remain explicitly unsupported. They previously passed signature
parsing while locals were transient; refusing them prevents a false success.
Full static procedures, duplicate declaration/compiler diagnostics, classes,
IDE reset semantics and an Access 2003 execution oracle remain open.


## Explicit persistent procedure-local Static declarations (2026-11-04)

STATIC_LOCAL_RUNTIME.md specifies actual per-host/procedure/source static
scalar/object/array state, nested-call alias refresh and current initialization
limits. Existing core cases cover persistence, native Dictionary/fixed/dynamic
arrays, independent scopes and recursion. Static procedure modifiers now refuse
incomplete implicit-local behavior instead of silently executing transient
locals. Full desktop 48/48 (79.78 s), freshly rebuilt GUI-free 1/1 (1.45 s),
targeted core 1/1 (1.26 s), Debian payload/permissions/Build-ID/pinned SQLite and
all four original hashes pass. No new CTest target. Scoped checklist 372 checked
/ 72 open, not complete Access 2003 parity.
