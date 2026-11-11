# Native project Enum type bindings

The native compiler accepts an explicit immutable set of project-visible Enum
names. Public/default-public declarations in standard and class modules are
available from the real checked project source catalog. Private declarations
remain source-local. Unqualified names must be unique and cannot replace a
module/class namespace; unique qualified names use the checked project module identity. A native module
need not embed Attribute VB_Name in its editable source. If a public Enum source
does carry that attribute, it must agree with its project owner. The source-only
API still creates qualified aliases only from its actual VB_Name attribute.
A source's own Enum names retain their local precedence. No declaration text is
injected and no type is invented from result values or a filename.

Enum As-types retain the existing Long binding: fields, constants, local/Static
variables, scalar/array parameters and results, Property Get/Let, Dim/ReDim and
inline wrappers use the same aliases. Native standard-module, CodeBehind, class
and pure SQL hosts pass the environment explicitly. SQL function signatures can
therefore expose a proven Long return. Global Enum is the standard-module public
compatibility form; class Global Enum and intrinsic type keyword replacement
are refused. Ordinary member constants keep their existing implementation.

Procedure plan keys include the exact alias set; only names are retained, never
a host/database/document. Module-state keys include the canonical alias set as
well as the actual source. Project metadata caches observe SQLite changes,
main data/schema revisions and temporary schema revisions. SQL callbacks use
the registry's already-checked snapshot instead of querying project metadata for
each result row. Existing source/catalog/SQL execution bounds remain in force.

Unqualified ambiguous names are not guessed; a unique explicit module qualifier
can still bind. Removing a provider removes its external alias on later binding.
The existing source-only API remains source-only when no aliases are supplied.
An independent project with the same consumer source cannot inherit another
project's Enum bindings from cached plans.

References: MS-VBAL5.2.3.4 (Enum declarations) and5.6.10 (type name binding tiers),
local reviewed MS-VBAL specification. Referenced-project/type-library resolution,
project-name qualifiers, complete Unicode/case-fold/compile-time constant grammar,
and general code-edit/reentrant compiler equivalence remain open. In particular,
this does not add general cross-module constant-expression folding merely because
an As-type can now name a public Enum. F040/F041/F042/F043/F122 remain open;
fixed catalog456/386 confirmed partial scopes/70 open full requirements unchanged.

The capability report remains available when a broken module prevents building
the project type environment. It reports that environment as unavailable with an
error number and continues its existing per-artifact invalid-definition/source
classification; it does not reuse stale aliases or claim complete type binding.

Validation used existing targets. Core passed0.51s after correcting the test
service's class-availability compiler to pass the same environment; DAO passed
5.36s. Final mainwindow target passed17.29s (test17.28s), app built. Shared cases
cover fields/constants/ByRef/ByVal/Optional/arrays/results/inline declarations,
qualified and unqualified types, class fields/properties and CodeBehind facade,
public class Enums, Global Enum, actual pure SQL Long metadata/results, private
refusal, ambiguity, removal/restoration, headerless checked owner qualification
and isolated/cache contexts. An actual standard-module call exposed an incorrect
Me requirement in the metadata host; it now checks project lifetime. An existing
invalid-module capability case exposed a report regression; unavailable type
metadata is now reported without aborting artifact classification. Debugger and
existing module/event paths also ran in the same UI target. No new CTest ID or
Access2003 execution oracle. Later code changes removed only temporary metadata
diagnostics and clarified comments; assertions/runtime behavior were retained.
