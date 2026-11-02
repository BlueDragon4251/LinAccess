# Project open/new keyboard shortcuts

Ctrl+O invokes the existing real project-open dialog and Ctrl+N the project-create
save dialog from the main window and nonmodal native forms/reports. Native
KeyPreview/KeyDown can cancel with KeyCode=0; pending control validation runs
before leaving the editor. Imported modal documents retain their modal context.
Opening a selected project uses the existing project-wide close path, including
Unload cancellation. Cancelling either dialog retains the project. Creating a
project refuses existing paths, preserving their bytes.

Synchronous SendKeys from executing native VBA cannot switch the project while
its old module/connection is on the stack. The native key handler reports its
existing unsupported status instead of opening a dialog; this does not raise a
new VBA Err through SendKeys. This restriction is intentional and remains open
for a future safe deferred project-switch contract.

The production-main-window integration test uses actual Qt file dialogs and
keyboard events on temporary SQLite projects. It checks KeyPreview cancellation,
both dialog cancellations, modal blocking, an Unload-cancelled open, successful
open/create, byte-identical overwrite refusal and the executing-SendKeys guard.
The targeted main-window/layout/original-heavy trio passes 3/3 (39.37 s).

These operations open/create LinAccess SQLite projects through the existing
project flow. Direct native MDB/ACCDB editing, complete Access File-menu/dialog
appearance, real window-manager behavior and every Access keyboard shortcut
remain open. No original business coverage increment is claimed.

Reference: archived Access 2003 keyboard help,
https://documentation.help/MS-Office-Access-2003/AccessKeyboardShortcuts.htm
