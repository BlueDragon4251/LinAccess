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
otherwise retain their existing role. Full Access database-window/F11, MDI/menu/
designer, acDialog and Windows modal/focus/event parity are not established here.
