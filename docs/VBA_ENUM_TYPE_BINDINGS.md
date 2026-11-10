# Source-local Enum type bindings

MS-VBAL5.2.3.4 defines an Enum name in an As clause as the declared type Long.
The runtime now resolves actual source-local Enum declarations accordingly,
instead of treating their names as potential project classes. Ordinary local,
Static and module/class fields, scalar/array parameters, Function/Get results,
Let values, scalar/array Dim, ReDim types and typed local/module constants use
the same binding. Names come from the real module declaration section; no type
is invented from an unknown identifier. Existing member constants, visibility,
qualified member names, duplicate/closed-declaration and signed32 checks remain.

Long coercion, range errors, zero defaults, optional arguments, ByRef Long
compatibility and Long-array metadata apply. Values need not be one of the named
members. Enum is not a class for TypeOf or AsNew. Cross-module/project-qualified
type aliases, full imported reference/type libraries, UDTs and complete Unicode
identifier/case-fold semantics remain open. Other Enum-value syntax limits are
unchanged; this is not full Enum/compiler parity.

Single-line If wrappers inherit real Enum aliases, constants, Def/Option state,
comparison mode and array base. These context-dependent inner plans are parsed
without using an unrelated wrapper cache entry. Original source text/files and
outer source-line mapping are retained. Full declaration hoisting/AST binding
remains F040.

Validation used existing targets, no new CTest IDs. The shared native/project
fixture covers fields, Get/Let, Enum/Long ByRef, Optional, typed constants, fixed/
dynamic/returned arrays, Preserve/zero defaults, inline declarations/constants
with OptionBase1, Long type metadata, arbitrary valid Long values and overflow6.
The pure scalar SQL fixture proves the actual Long result metadata and result.
Modules retaining array/object state keep the existing SQL3251 guard. First
runtime checks exposed an unsupported numeric-suffix fixture and the array-state
SQL guard; those fixtures were corrected without loosening transport guards.
Final core/GUI2/2 passed23.08s (4.60/18.48). App built. Original event target passed
1/1 in121.37s, retaining its known missing-macro boundary; seven upload/fixture
hashes are unchanged. No Windows oracle or full-suite/DEB refresh for this scope.
Fixed catalog456/386/70 remains unchanged; complete F040/F042/F043 stay open.


## Same-source qualified Enum type names

An actual source Attribute VB_Name now supplies the owning module qualifier for
that source's Enum As-type aliases. Both EnumName and ModuleName.EnumName bind to
Long in the existing declaration paths, including typed Const, fields, parameters,
results, scalar/array Dim/ReDim and inline wrappers. The scanner reads only the
preprocessed declaration section, rejects duplicate module identity attributes,
and does not infer a module identity from a filename, class guess or database
label. Without an identity attribute no module-qualified alias is invented.
Other.ModuleEnum and project-wide visibility/ambiguity/reference libraries remain
open. This does not add project-global aliases, inject source declarations or
make an Enum constructible as a class. Source-only parse plans remain source-only;
no database/host type environment leaks into their cache key.

App/core/mainwindow built. One existing core+GUI bundle2/2 passed25.29s:
core5.51s, mainwindow19.78s, unchanged precise20s watchdog. The complete existing
Enum fixture is also executed with real same-source qualified As clauses; the
actual owned SQL module registers its qualified Enum return as Long and executes
EnumSqlIdentity(1234). Foreign qualified type is rejected. No new CTest/catalog
IDs;456/386/70 unchanged. No Windows Access execution-oracle claim.
