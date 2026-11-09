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

Indexed properties use native typed/named getter argument binding. Same-module
unqualified calls support ByRef indices; the external object facade currently
requires explicit ByVal indices. Matching accessors must have equivalent index names,
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
user class construction/lifecycle, user-defined default-property attributes and
arbitrary argument/default queues, cross-module property binding and cross-object ByRef index copy-back,
ParamArray mutation/forwarding, named setters and complete visibility/COM
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


## Object-valued getters and native default members

MS-VBAL5.6.13 permits a nonempty positional index list on an object/Variant-valued
property or function that accepts no parameters: the list belongs to the returned
object's default member. The native VM now executes the getter/function once and
forwards that list, preserving object results for Set. Same-module functions and
properties and public form property reads use this path. Nothing fails91; a
nonobject result fails13; the original getter error propagates. An actual scalar
getter with excess arguments still fails450 before its body effects.

Writes first prefer an actual Let/Set accessor. With no setter, a zero-parameter
object-valued getter can supply the native default indexed assignment, including
Set dictionary items. A declared native-object getter with matching index
parameters can supply a scalar default write to the returned control, following
MS-VBAL5.4.3.8. No getter is executed merely to discover whether a setter exists.
Arguments/RHS are evaluated once; getter and default member are called once.
The source owner is checked after getter and default operations. Closing its
real form during the getter fails2450 before touching the former control value.

This covers existing native object capabilities, not arbitrary custom class
attributes, array-valued getter mutation, function-result writes, named unbound
default indices or complete Variant/COM/default argument queues. Public facade
ByRef index arguments still fail explicitly. Existing visibility/lifetime and
native value bounds remain. No Windows Access oracle is claimed.

Application built; first runtime core/GUI bundle2/2 passed (17.19s). Actual native
Dictionary values/reference identity and real TextBox writes/read counters,
Nothing/error propagation/scalar excess arguments and owner-close rejection were
executed. No new CTest ID. All32 original VBA artifacts had0 properties and0
zero-argument object-function candidates; no extra original run was added.
Catalog456/386/70 and full F042/F043/F050 requirements remain unchanged.


## Same-module ByRef index transport

Direct unqualified Get/Let/Set calls in the owning module now reuse the real
procedure argument binder instead of copying evaluated indices into a detached
argument vector. Existing local/module aliases, typed variables and captured array
elements therefore update their actual caller locations. Temporary expressions
stay temporary. The final setter value remains ByVal even if declared ByRef.
A default-value assignment through an indexed object getter uses that same binder.

Changes made before a getter/setter error are published through the existing
alias mechanism; Err/On Error retain their normal native behavior. Native ByRef
type mismatches and active-array-element allocation guards remain in force.
Named setter indices are still refused. `publicOnly=false` is the module-owned
API path; direct caller-owned argument vectors also receive Get index changes.
The public form/object facade still rejects ByRef index signatures with3251,
because its value-only interface cannot supply caller reference bindings.
No external reference copy-back is simulated.

The existing core/GUI bundle passed on its first run2/2 (17.35s), application
built. Actual cases cover Get/Let/Set index changes, array element indices,
temporaries, setter-value isolation, object identity, error-side mutations,
ByRef type mismatch, API argument vectors and a real TextBox default write through
a ByRef getter. Public facade rejection remains checked. No new CTest IDs or
Original/Windows/full-suite proof; F042/F043/F050 and456/386/70 remain unchanged.


## Optional and ParamArray indices

Property parameter grammar now keeps the final required value parameter separate
from the index parameter list (MS-VBAL5.3.1.5/5.3.1.7). Optional indices may precede
that value. Missing positions are inserted before the value, never after it;
caller-owned argument vectors retain their original length. Provided ByRef indices
copy back to their original positions; omitted defaults remain local. Matching
Get/Let declarations can have different default constants as the specification
allows. The setter value stays ByVal regardless of its source mechanism.

Positional holes and empty index lists work for supported property assignments,
including `Cell(,8)=7`, `Cell()=7` and the corresponding object-facade calls.
Named optional getter arguments preserve Missing rather than silently becoming
Empty/0. Empty/omitted indices in unsupported array contexts are refused by the
binding preflight before body effects. Syntax acceptance alone remains unproven
runtime capability, as in the existing capability report.

A property ParamArray gathers only index arguments. Its final setter value is
bound from the actual last argument, including when the ParamArray is empty or
has more elements than there are formal parameters. Scalar/object Let/Set values
remain separate and object identity is preserved. Native signatures advertise
variadic shape so named ParamArray calls are refused consistently. ParamArray
holes, named ParamArray arguments and element mutation/forwarding remain explicit
native limitations; element writes fail3251 before changing even the local input
collection. Existing native type/default/object, reference and budget limits apply.

Application built; first existing core/GUI bundle2/2 passed (16.60s). It executes
ordinary/qualified getter and setter calls, distinct Get/Let defaults, partial
ByRef indices, preserved API vector shape and RHS values, empty/filled scalar and
object ParamArrays, real public form facade calls, and negative array-hole,
ParamArray-write and named-variadic cases. No new CTest ID, new Original/Windows/
full-suite/package proof. The complete F042/F043/F050 requirements remain open;
the catalog stays456 IDs/386 confirmed partial scopes/70 open requirements.

## Typed array function and getter returns (2026-11-09)

Native Function/Property Get declarations accept `As T()` (including spaced
array designators) for Variant, Byte, Integer, Long, Single, Double, Currency,
Date, Boolean, String and Object. Return bindings use the existing dynamic-array
initialization and exact element-type assignment rules: an unassigned return is
unallocated, not Empty; mismatched element types fail13. Returned values retain
rank, lower/upper bounds, typed cells and independent whole-array copy semantics;
Object cells retain actual object identity. Class/UDT/COM array types remain refused.

A no-argument function/getter whose actual result is an array can be read with
indices, including Variant results and public form getters. The producer runs
once and existing rank/bounds checks apply. A function with parameters uses a
separate result-index queue, e.g. `Build(2)(1)`. Array-valued result writes remain
outside this scope. Known typed-array SQL functions fail3251 before executing the
body; arrays are not serialized into fake SQLite scalar values. Ordinary scalar
SQL calls in the same module continue to execute through the native query adapter.

Application built; first existing core/GUI bundle2/2 passed (16.91s), including
actual form facade reads, retained arrays after form close, typed Byte/String/
Object/matrix values, copy independence, bounds/type failures and the real shared
SQL/module-state counter proving pre-body rejection. No new CTest, Windows oracle,
Original or full-suite run. F043/F050 remain open. Catalog456/386/70 unchanged.
Latest checked Debian package is1313196 (before this array-return scope), with
payload, permissions, exact Build-ID and pinned SQLite/MDBTools checks passing.
