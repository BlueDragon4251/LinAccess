# Imported documents opened from the main window

The imported form/report actions in the source-artifact inspector now close the
read-only details dialog and show a managed native AccessFormPreview. They no
longer enter QDialog::exec or keep the details dialog blocking the application.
Modal metadata still creates an application-modal window, without suspending the
C++/VBA caller. This does not implement the separate acDialog execution mode.

The public native launcher reuses one open same-kind/name instance within the
same database connection/project, preserving existing state instead of firing
Load again. New targets validate document kind before construction. Cancelled
Open or initialization-time self-close returns no shown document. Constructor/
source errors are reported through the existing UI error path. Source-preview
data/unsupported-initialization diagnostics retain their former scope; this does
not promote the inspector to a fully validated business execution environment.

The project main window owns the widget; an invoking details/form widget never
owns its lifetime. finished uses the same deferred-disposal mechanism as native
DoCmd-opened documents, waiting for outer VBA calls before scheduling deletion.
Project-close guards and same-instance closing checks prevent reopening during
close. Artifact/module read cursors finish before constructing/activating a new
document, avoiding long-lived source readers.

## Evidence

A new mainwindow_documents CTest compiles the actual production GUI sources and
uses the actual project-open file dialog and explorer signal/action route with
own temporary SQLite projects. No production form/window logic is mocked. It
checks independent nonmodal form/report lifetime after details destruction, one
Load for repeated opening, actual dirty SQL saving, cancelled project switching
through Form_Unload, modal metadata, cancelled Form_Open and self-close followed
by a nested native MsgBox. Deferred-delete delivery during the message box proves
the closing document remains alive until VBA unwinds. The project subsequently
switches to a second temporary database. Existing layout/original-heavy tests
remain in the targeted trio, which passed 3/3 (38.93 s) before final validation.

The fixture corpus gains no new original business execution. This uses synthetic
layouts/modules/rows only; all original fixture hashes remain separately checked.
The desktop suite now includes 22 CTests when all fixture options are enabled.

## Boundaries

The older generated/partially converted LinAccess form editor still runs in its
existing modal dialog and uses its prior nested opener. Converting only that
child to nonmodal would leave it blocked by the still-modal legacy caller. Its
renderer/host migration is a separate open task. Source inspector/editor dialogs
otherwise retain their existing role. Bounded classic F11/explorer focus now exists (DATABASE_WINDOW_RUNTIME.md).
Full database-window command, MDI/menu/
designer, acDialog and Windows modal/focus/event parity are not established here.


## Real table documents in the native window ring (2026-11-03)

Window menu and Ctrl+F6/Ctrl+Shift+F6 now include actual native table tabs between
the database window and native forms/reports. Existing views are activated, with
weak targets rechecked across native validation; source field errors block the
switch, while valid row drafts remain unsubmitted. Mouse tab navigation is
checked before hiding the old page; pure modifier keys do not close an editor.
Native form/table mixed round trips, menu/table direction switches, unavailable
field transitions and preceding numeric/NULL/clipboard behavior are covered in
the existing targets. See WINDOW_CYCLE_RUNTIME.md for scope and remaining MDI/
query/editor-window limitations.
Full build and suite pass 48/48 (92.38 s), Debian payload/permissions/Build-ID/
pinned SQLite checks and all four original fixture hashes pass. GUI-free core
remains 1/1 (1.53 s), with subsequent Qt-only changes. Theoretical checklist:
283 checked / 70 open scoped points, not complete Access 2003 parity.


## Existing table reuse and native SelectObject (2026-11-03)

Opening the same table again activates its existing native page/model instead
of creating a duplicate. Pending row drafts remain intact; the existing model
is not replaced, refreshed or silently submitted. Initial table opening and
activation use the same source validation/modal guards as document switching.
DoCmd.SelectObject acTable,name,False now resolves an actually open table through
the owning production MainWindow and activates its real view. Lookup uses the
existing table metadata with case-insensitive name comparison. The current
source field is validated, activation can fail with 2501, and modal context
returns 2046. A missing/closed table and unsupported owner remain explicit 3251;
exact Access closed-object errors, OpenTable and query-window activation are
not implemented by this increment. The database-window True route remains
separate. Generic table DoCmd commands remain outside the native form/report
command subset; selecting a table must not retarget subsequent commands to the
old caller.
The main-window regression reopens the same dirty table, proves model identity/
one tab/draft retention and unchanged SQL, then runs interpreted SelectObject
with an uppercase table name. Actual focus moves to the table, while Me.Tag
still refers to the caller and closing that caller preserves the table page.
Existing main-window/layout targets pass 2/2 (15.56 s). Last full suite/package
remains 671d997 / private API fea77b0, 48/48 (92.38 s); this UI/runtime increment
is focused-verified. Theoretical checklist: 285 checked / 70 open scoped points.


Native OpenTable ordinary/read-only datasheets are now implemented with checked
argument binding and model mutation guards. See TABLE_DOCUMENT_RUNTIME.md for
verification and the remaining append/design/preview/general-command limits.
