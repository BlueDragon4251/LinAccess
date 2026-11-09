# Native macro project commands

Stored native macros now expose Close (optional type/name/save), DeleteObject
(explicit numeric type/name) and Rename (new name/type/old name). They call real
existing document/catalog operations. Rename supports owned closed tables/queries;
DeleteObject supports owned tables/queries/macros, with existing open-source,
revision and running-macro guards. Imported archives remain outside write/delete
ownership. No automatic caller/reference rewrite or implicit delete target is added.

The standalone Close proxy uses actual query/table close APIs or an actual project
form/report. No hidden caller form is created. Default Close uses the actual active
document; unsupported object/editor types refuse3251. Save options0..2 concern
object design. Normal pending record validation/save still occurs with acSaveNo.
Existing close paths clear SQL models before deferred widget disposal, so a
subsequent supported schema operation need not wait for a UI event-loop turn.

Project macro execution now has its own lease instead of pretending to be an
active SQL query. Existing object commands can run while project/application close
and a second UI root macro remain blocked. Shared active-macro frames also guard
project close for direct form/module calls. Native explicit nested RunMacro still
uses its documented shared depth/action limits. Complete Access stop/reentry/window
scheduling equivalence remains open.

Dispatch into a real form/report holds the existing native document execution
scope until the command unwinds, including Close/Requery/RunCommand. Reusing an
already open stored document does not replay prior initialization failures; a new
failed document initialization is still surfaced and cleaned up. This is bounded
native document behavior, not a new Access2003 event/error oracle.

Validation: existing GUI/layout2/2 passed (18.92s), matching app build. A real menu
macro renames/opens/closes/deletes owned table/query objects and another macro,
closes an active query by default and saves a pending numeric record43 with
acSaveNo. Real Form_Unload/Report_Close SQL and subsequent macro SQL execute.
Actual window.close during a MsgBox is denied with unchanged project generation;
the macro continues afterward. Existing historical Load7003 is not rerun, while a
new initial Load7003 remains an error and leaves no live failed document. No new
CTest target, full-suite, package or Windows observation is claimed here.


## SelectObject

The native SelectObject macro action accepts numeric type, object name and optional
InDatabaseWindow Boolean. The standalone DoCmd API uses the corresponding named
parameters. Open table/query selection activates the actual existing document;
closed objects are not silently opened. Real open forms/reports use the existing
validated preview selection path and execution lifetime guard. Missing/ambiguous/
modal targets retain explicit native failures.

InDatabaseWindow supports native table/query/form/report/macro/module selection,
including the corresponding folder when name is blank. Macro/module selection is
limited to the database-window context; no nonexistent editor/VBE window is
fabricated. The real form host accepts these native macro/module selections too.
Imported archive-only module navigation and complete Access context equivalence
remain open. Explorer refresh preserves exact logical selection or archive ID
only within the same project generation, so a macro's final selection survives
its normal catalog refresh.

The standalone execution host implementation now lives in its own Forms C++ unit
with a small declaration header. This avoids recompiling its implementation inside
the main window and test translation unit on every runtime change; the Qt-free
core boundary and application dependencies remain unchanged.

Validation: existing GUI/layout2/2 passed (19.03s), app and matching DEB checked.
Actual open table/query/form/report activation, macro/module/folder selection,
refresh preservation, final menu-macro selection and real form-host4/5 selection
are exercised, alongside unavailable targets. No new test target or Windows oracle.
