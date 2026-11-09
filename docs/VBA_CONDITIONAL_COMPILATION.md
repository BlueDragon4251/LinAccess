# Native VBA conditional compilation

The native, Qt-free runtime preprocesses #Const and nested #If / #ElseIf / #Else /
#End If (also #EndIf) before procedure parsing and module-state creation. Excluded
ordinary code is not syntax-validated or registered as live procedures/variables.
Only the first true branch of an included group contributes code. Every directive
expression is evaluated, including ElseIf after a selected branch and directives
inside excluded groups. Every #Const is a module conditional binding, including
constants in excluded groups. Duplicate names and dependency cycles are refused.
References between conditional constants are resolved as module bindings, including
forward dependencies; undefined names evaluate as Empty. These bindings do not
become runtime module fields. A typed suffix on #Const's left-hand name is ignored.

Expressions use the existing native scalar evaluator through a host that cannot
read/write fields, save records or display messages. Literal/operator expressions
and Int/Fix/Abs/Sgn/Len/LenB/CBool/CByte/CCur/CDate/CDbl/CInt/CLng/CSng/CStr/CVar
are allowed. Arbitrary function/object/member access is refused. Existing native
coercion, date, numeric and string limits still apply. Conditions use Let-to-Boolean
conversion; Null is an invalid condition, rather than a runtime If's false branch.
Unsupported grammar returns an explicit error before host-side procedure effects.

The current Linux project defaults are Win16/Win32/Win64/Mac=False, VBA6=True,
VBA7=False. These are native host/language-target flags: Linux does not advertise a
Windows API host. Module #Const may shadow a project-default name. No Windows API,
COM/ActiveX activation or Office installation is supplied by selecting a branch.
Custom project-level compilation-argument configuration is not added here.

Physical line positions survive preprocessing, including continued directives.
Storage, source hashes, debugger identities and observation records retain the
original source. The procedure cache is keyed by original source bytes. Runtime
array bases, comparison mode, module constants and variable/procedure inventories
use included code. SQL VBA, native standard-module/form cross-module resolution,
Eval and stored form/report event binding inspect included declarations. Archival
syntax/reference inventories deliberately still inspect original source and do
not constitute execution proof. The editor's archival procedure list remains a
syntax inventory; selecting an excluded-only routine cannot make it execute.

Limits are one MiB source, 4096 directives, 4096 bytes per directive expression,
64 nested groups/constant dependencies and32 physical lines per continued directive.
Expression text inputs/output are limited to64KiB and total resolved conditional
constant text to1MiB; these budgets stop exponential concatenation before unbounded
expansion.
Binary native expression comparison is used for conditional directives. This does
not claim full Unicode conditional names, every LongLong/LongPtr intrinsic, project
compilation settings, locale/compiler diagnostics, or Windows Access 2003 parity.

The existing core fixture exercises actual selected procedure execution, module
state, local calls, Option Base/Compare, undefined constants, excluded constant
processing, eager excluded-guard errors, continuation, source mapping and actual
source-identity debugger stops. The existing DAO fixture calls a conditional
standard-module function through the real SQL UDF bridge, verifies the selected
result type and absence of an excluded function. No new CTest ID or business data.

Normative basis: MS-VBAL v20260818 sections3.4,3.4.1,3.4.2 and5.6.16.2. These
specification rules are distinct from an executed Access 2003 reference oracle.
F040/F042/F043/F045/F051/F068 remain open; fixed catalog456/386/70 is unchanged.
