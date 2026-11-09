# Native standard modules

**+ Modul** creates an actual native standard-module document. The editor saves
UTF8 text in `_la_objects(kind='module')`, reopens it through the Module explorer
group and supports Ctrl+S without closing. Existing imported VBA archive rows are
never rewritten by this editor. A new name cannot silently replace an existing
native/imported module. Saved names are fixed in this initial editor; renaming,
qualification/refactoring and a full VBA IDE remain open.

The Analyze button produces a syntax inventory, not an execution guarantee.
Standard-module text is limited to1MiB encoded/decoded UTF8 with valid UTF16/no
NUL at the editor boundary. Native definitions contain version1 and `vba`; class
module sources are explicitly refused. Unsupported ordinary VBA can remain saved
as source; actual native calls still undergo the existing compiler/runtime
preflight and supported language limits. No sample business data is supplied.

Native calls, Eval public-function lookup, public module-variable lookup and the
SQLite public-function registry read a common source catalog. It preserves latest
imported versions per case-insensitive module name and merges native module
sources. Native/native and native/imported collisions fail5 rather than choosing
a source. Existing physical-main/archive shape and active TEMP-shadow checks are
retained; native catalogs receive corresponding checks. Limits remain256 latest
modules,16MiB encoded source,1M scanned source rows and a5-second metadata scan
bound. Native definition JSON is limited to4MiB; native text itself remains1MiB.
Original archived bytes and their encodings remain stored unchanged.

Public Functions use the existing bounded native VBA engine. Private helpers remain
callable inside their module. The SQL host retains its existing scalar-state,
argument/result and unavailable GUI/DAO-context restrictions; this is not complete
Access SQL/VBA context equivalence. Module globals and supported Static state share
the existing connection-owned source frame. Editing a body selects its new frame;
unchanged source retains the prior frame. Complete live editing and transactional
VBA-variable semantics remain open.

The SQL registry's existing source/statement revisions now include native module
text. Prepared functions observe actual source updates and restoration after SQL
savepoint rollback. Removed sources cannot continue through a cached old body.
Function registration occurs through normal LinAccess/DAO query entry points;
raw third-party QSqlQuery preparation is not a new public registration API.

CurrentProject.AllModules exposes the supported detected standard modules from
both stores, with actual Name, Count, Item, enumeration and readonly metadata.
IsLoaded/SysCmd acModule state reports a real tagged editor, including modified/new
flags; full Access compiled-module memory state and class-module catalogs remain
open. Held descriptors retain existing project-generation guards and snapshot
name behavior.

Delete removes only the saved native module, with real confirmation, a saved
revision comparison and current project identity/generation checks across the
modal interaction. Tables, imported archives and other modules are not removed.

Validation: existing DAO1/1 (2.55s) and layout1/1 (11.36s) passed in the bundled
run. After correcting actual Qt focus for Ctrl+S and using the normal query entry
point for first SQL function registration, only the GUI target was repeated and
passes1/1 (6.23s); app built. The real editor path verifies save/reopen/update,
AllModules/SysCmd state, shared VBA/SQL Counter and private-helper execution,
private SQL-function exclusion, prepared body replacement/rollback, TEMP-shadow
and imported/native-name collision refusal, native deletion/stale prepared-query
refusal and unchanged pre-existing archive bytes throughout. No new CTest target,
fresh full-suite/platform matrix or Windows Access2003 oracle. F042/F050/F068/
F118/F130/F155 remain open; fixed456/384/72 unchanged.

## Execute a saved module in the editor

The native module editor lists analyzed procedures and provides a Run button/F5.
It executes the selected saved procedure with zero supplied arguments. Existing
Optional defaults apply; required arguments produce the runtime arity error.
Saving refreshes the inventory; typing source disables Run until Analyze/Save.
Unsaved source is refused, and the stored source is rechecked against the guarded
shared module catalog before execution. Selecting a procedure is transient UI
state and does not mark the module source dirty.

This is a real editor/database host. CurrentDb and DBEngine use the current native
project; public standard-module calls and module/Static state share the same
connection-owned frames as form/VBA/SQL calls. Supported locale-aware comparison,
numeric/date formatting, MsgBox and InputBox reuse existing native implementations.
Dialogs and module entry recheck project identity/generation. There is no hidden
or fabricated form. Me/current-record operations and unavailable application/UI
members fail explicitly. This initial editor host does not supply a complete
Access Application/Forms/Reports object model or debugger.

Results show actual completion/typed scalar return text, or runtime error number,
line and diagnostic. Syntax inventory remains distinct from actual execution;
unsupported procedure syntax must fail preflight before its prefix SQL effects.

Validation: existing GUI1/1 passed (7.20s), matching app build. Actual F5,
DAO write and public call into a second module, shared counter, return value,
MsgBox response and arity/context/preflight/unsaved refusal were exercised.

## Shared CurrentProject and CurrentData

The standalone Application exposes CurrentProject and CurrentData, including bare
VBA identifiers through the normal evaluator. They use the existing actual native
catalog implementation rather than a form placeholder or a copied name list.
AllModules/AllForms/AllReports/AllMacros and AllTables/AllQueries retain their
existing bounded catalog, descriptor, readonly and IsLoaded contracts.

The form and standalone applications share the same project/catalog object
identity under the actual project window. The weak cache key includes connection,
database and project generation. A held catalog survives closure of a caller form;
closing/reopening or switching the actual project invalidates old project/data
objects and descendants with3420. Native file Name/Path/FullName retain the existing
file-backed-project requirement. No native AccessObject export/import or arbitrary
project-property implementation is implied.

Validation: matching app built; existing GUI/layout2/2 passed (19.49s), no test
correction. Actual VBA counts, shared identity, caller-form closure, live module
removal and3420 after actual project switching are covered.

## Actual Forms, Reports and Screen from standard modules

Bare Forms/Reports and Application.Forms/Reports in the standalone host expose
actual native project document objects. Count, numeric/name Item access, direct
Application collection indexing and For Each reuse the existing form/report
collection implementation and real per-document references. A document hidden
after opening remains addressable; closing it removes it from subsequent live
lookups. A held document reference still fails when that actual document closes.
No hidden placeholder form is created to supply the collection.

Application.Screen and bare Screen expose the existing actual ActiveForm,
ActiveControl and MousePointer implementation, with its current documented bounds.
MousePointer assignment uses the owned cursor contract and rechecks context after
scalar-default evaluation. Active lookups are isolated by connection and database
identity. Native project-owned collections/Screen survive caller-form closure and
reject stale generations3420 after actual project switching. Existing form-owned
collection/Screen handles retain their prior caller-lifetime behavior.

This is not a complete Access Application/Screen surface: ActiveReport,
PreviousControl and arbitrary screen properties remain unsupported. Global object
identity across independently created Application handles and full original
hidden/modal/new-instance ordering still require broader implementation/oracle.

Validation: matching app built; existing GUI/layout2/2 passed (20.16s), no test
correction. Real visible/hidden forms/report, collection identity of document
references, VBA enumeration/active screen/cursor, live closure and stale3420
after actual project switching are covered.
