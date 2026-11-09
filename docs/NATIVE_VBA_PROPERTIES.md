# Native VBA Property accessors

The real interpreter supports explicitly selected Property Get/Let/Set bodies,
End Property and Exit Property (including a single-line If). Accessors sharing a
source name have separate translation, invocation, debugger and observation
identities. Original source bytes and line positions remain the module identity;
there is no generated replacement module or independent state for each accessor.

Same-module unqualified reads/writes and public native Form/Report member access
use the actual module state. Getters return their source-named result variable;
object results and Set preserve reference identity, including Nothing. The final
Let/Set value parameter is always ByVal, even when its source says ByRef, as
specified by MS-VBAL5.3.1.7. Parameters/function results remain per-invocation.

Indexed properties use explicit ByVal index parameters and existing typed/named
getter argument binding. Matching accessors must have equivalent index names,
types and passing mechanisms; Get/Let value types must agree. Duplicate accessors
and mismatched signatures fail explicitly. Setter detection and metadata inspection
never execute a getter. A failing accessor propagates its actual native error;
unsupported declarations fail before its body effects.

Actual form instances keep independent state. Private accessors are available to
same-module unqualified code and are excluded from external object metadata.
Closed owners reject access with2450. Existing executing-VBA project-switch and
module-lifetime guards remain in force. Get-only scalar writes fail with native383.
Module selectors and capability observations distinguish all three accessor kinds;
source recognition alone is still not an execution or compatibility claim.

## Explicit boundaries

This is a bounded F042/F043/F050 extension, not complete class support. Generic
user class construction/lifecycle, default-property attributes/implicit default
forwarding, arbitrary cross-module property binding, ByRef index copy-back,
ParamArray, optional setter indices, named setters and complete visibility/COM
semantics remain open. Parameterless Sub/Function/Get declarations now allow omitted parentheses, as
specified by MS-VBAL5.3.1. Module inventory and actual form/core execution use
the same source names, including initial/trailing Static. Setters still require
value parameters; Sub/setter result types are refused before body effects. Static
property lifetimes are documented in STATIC_LOCAL_RUNTIME.md. Complete declaration
grammar remains open. Private
indexed/object writes should use same-module unqualified access; a public object
facade does not grant private access. Existing native value/coercion limits apply.
No Windows Access property/event/locale oracle is claimed.

## Verification

Application built; existing core_tests/mainwindow_documents bundle2/2 passed
(16.12s), after fixing the two actual binding failures found by the first bundle:
bare getters and indexed form argument dispatch. Tests execute the same technical
source in the Qt-free core and real independent form instances. They verify ByVal
setter value semantics, read counters without speculative getters, indexed values,
retained dictionary identity/Nothing, private metadata, errors11/383/2450 and
single-line Exit Property. No new CTest ID. All32 original private VBA artifacts
were inventoried without exposing sources: none declares a Property accessor;
therefore no new original-workflow or Windows proof follows from this feature.
The fixed catalog stays456 IDs,386 confirmed partial scopes and70 open requirements.


Parameterless declaration extension: application rebuilt; existing core/GUI2/2
passed on its first bundle (15.90s). Actual sub calls, bare function calls, typed/
Variant getters, initial/trailing static counters and form instance values were
executed. Invalid parameter syntax, typed Sub and valueless setter were refused.
No new CTest ID; original32-artifact inventory had zero such header candidates,
so no additional original/Windows claim. Catalog456/386/70 is unchanged.
