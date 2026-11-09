# Native Application.Run

The native Access Application object exposes `Run(Procedure, Arg1, …, Arg30)`.
An unqualified public Function or Sub is resolved through the shared standard-module
catalog used by native VBA, module globals and SQL function registration. Imported
archive sources and newly authored native modules use the same lookup. Private
procedures are not exposed; ambiguous public names fail rather than selecting one.

Arguments retain their native Variant types. Optional parameters use the existing
procedure binder, Functions return their typed value and Subs return Empty.
VBA errors retain number, source and help metadata. Named Application arguments
use Procedure and Arg1 through Arg30; these are not names of the target parameters.
Source compilation and argument checks precede execution of the target body.
Calls remain subject to the supported native VBA language and host context.

Project-qualified names and library/reference search remain unsupported3251.
A missing public target reports2517. At most30 target arguments and32 nested Run
calls are permitted; the latter fails28. This API receives argument values; it
does not provide caller-variable ByRef writeback. There is no new Automation/COM
server or full VBA IDE claim.

Contract reference: Microsoft Learn, Application.Run method (Access),
https://learn.microsoft.com/en-us/office/vba/api/access.application.run .
Documentation is not a fresh Access2003 execution oracle.

Validation: existing mainwindow_documents1/1 passed (6.48s), matching app build.
Actual native module editor/source, named VBA invocation, Function/Sub/optional
arguments, shared globals, Private/missing/project/arity guards, preserved raised
error source and recursive Run failure28 with subsequent successful independent
call are exercised. No new test target or Windows Access2003 oracle.
