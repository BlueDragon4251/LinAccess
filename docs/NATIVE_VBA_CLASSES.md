# Native authored class instances

Native owned class modules are saved as version1 UTF8 VBA documents in the real
project catalog (`kind=class`). The class editor and explorer create/save/reopen/
delete that source. A first body-only save writes a real CLASS/VB_Name header;
subsequent edits preserve the actual source text and line numbering. Readonly
imported VBA bytes are not rewritten. Native class names are ASCII identifiers
up to64 characters. The shared source namespace rejects collisions with ordinary
native/imported module names. AllModules includes actual owned classes, and its
existing editor-loaded state recognizes the real class editor. Public SQL/VBA
standard-module discovery skips class code; class methods are not registered as
free SQL functions or public global variables.

New creates a real independent instance through the native runtime. Module fields,
arrays and Static locals belong to that instance, not to a singleton standard
module frame. A private parameterless Class_Initialize Sub runs once per successful
construction. Ordinary public Functions/Subs, public fields/constants, Property
Get/Let/Set, native array returns, indexed reads/writes and native returned-object
default forwarding use the existing compiler, type and argument binder. Object
identity, Me identity and actual object references are retained. A Sub is permitted
in an object statement, and refused before its body when used as a value. External
class arguments currently require explicit ByVal; direct unqualified owning-class
ByRef calls retain the real existing aliases. No detached value-copy writeback is
presented as ByRef.

The lexical execution context tracks the actual source and host for every VM
entry. Standalone/nested expressions through another host do not inherit the
parent class's private-member privileges. Reentry into another procedure/source
uses that procedure's lexical scope; returning restores the caller's scope.
Private members remain available inside their class, including same-class instance
access, and unavailable from other modules/callback expression hosts. Constructor
references are bound before procedure effects; source strings and raw control
names containing New are not treated as constructors.

Class methods use genuine project services. CurrentDb/DBEngine, public standard
module calls, supported formatting and application commands use an owned service
host; no fabricated hidden Form exists. A class tied to an actual project window
outlives its creating editor/form. A standalone document uses its guarded document
lifetime. Database/workspace identity, project generation and source revision are
checked before/after method calls and before field reads/writes. Once an observed
revision/deletion/context change invalidates an instance, restoring that revision
does not revive it. Unobserved byte-identical delete/recreate, all external editing
orders and full Windows lifetime equivalence remain outside this proof.

Explicit native boundaries: references are held As Object/Variant; named custom
As types/As New, imported generic class construction, Implements/WithEvents,
default/predeclared instances, custom default-member attributes, COM/ActiveX and
complete external ByRef are not claimed. Class_Terminate is rejected before
construction instead of silently pretending cleanup ran. Bang/default member
binding and array-valued getter mutation retain explicit refusals. The existing
one-MiB source/array/procedure limits apply; live instances have a4096 count and
16MiB retained-source budget. This does not complete F040/F042/F043/F050 or claim
full Access2003 class parity. Fixed catalog456/386/70 remains unchanged.

Verification: application and the existing core/GUI/DAO/Original targets built.
First core/DAO passed5.20s/3.49s; Original exposed the Def preflight regression
and GUI hit its20-second watchdog under the concurrent block. The Def binding
was restricted/caught, constructor discovery excludes raw control names, and
private lexical scopes were completed before publication. The final existing
core/GUI/Original block passed3/3 in143.28s: core4.08s, GUI17.58s, Original121.62s.
Only affected targets were repeated; no new CTest IDs or broad full-suite run.
The GUI editor diagnostic observed no unexpected error dialog in the successful
run. All seven actual upload/fixture hashes remain unchanged and core remains
Qt-free. Native core/GUI cases cover independent initialization/state, Sub versus
value calls, fields/constants, properties/indexing/arrays, real ByRef internally
and explicit external refusal, named optional calls, Me identity, object default
forwarding, private callback/reentrant scopes, actual CurrentDb SQL writes,
class SQL-function exclusion, real editor Ctrl+S/header/catalog state, instances
outliving caller windows, sticky revision invalidation and TEMP/deletion guards.
The Original event path retains its known missing-macro boundary, not a positive
initialized sale proof. No Windows oracle or fresh class package proof. Last
checked Debian3fe7245 precedes this scope.
