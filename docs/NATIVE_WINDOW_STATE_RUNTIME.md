# Actual native window-state commands

Standard modules and native macros now expose Maximize, Minimize and Restore.
They accept no arguments. The specialized VBA DoCmd.Maximize statement also uses
the checked native module host override, rather than the generic host's missing
form error.

An actual loaded active form/report delegates to its existing native window
controller. An active table/query or database explorer addresses its real project
window. No temporary form or fake Me is constructed. GUI thread, project identity,
query/output reentry, modal context and owner lifetime are checked. Module-editor
windows refuse3251, consistent with the documented exclusion of VBE module windows;
other modal designer contexts refuse2046. Selecting a minimized form explicitly
before Restore uses the existing real SelectObject path.

Window changes do not issue a record save or discard table/form drafts. The
existing mainwindow_documents check passes (8.69s): actual Qt state transitions,
VBA's specialized statement, a three-action macro, real form targeting, retained
SQL/cache/Dirty values, extra-argument450 and unchanged module-editor text on
refusal. The layout check passes (11.29s). Qt can preserve the maximized flag
while minimized, so evidence checks the actual state bits rather than assuming
mutually exclusive flags.

This is native Linux window behavior. Full Access MDI propagation across all
windows, popup rules, designer/VBE window control, exact geometry and a Windows
window/event oracle remain open under F052/F088/F135; no catalog count changes.

The native `MoveSize` macro action moves or resizes the active normal form
window. It accepts up to four whole twip values for left, top, width and height.
At least one is required. Omitted arguments keep the current value. The action
uses signed 16-bit twips with positive dimensions, converts at 96 logical pixels
per inch and positions the client rectangle relative to the project window.
The macro persists its arguments. Moving a window leaves the form definition
and any record draft untouched. The action requires a visible, active top-level
form and reports error 2046 for unsupported window states, popup forms, modal
contexts, other document types and positioning on Wayland.

The native macro parser/catalog/load functions now compile once in linaccess_data
instead of being defined in every including translation unit. Their signatures,
validation, limits and runtime meaning are retained. The GUI extension checks use
a stable single entrypoint, allowing further cases to compile in their smaller
source file without editing the large mainwindow test file. No new runtime DSO,
GUI dependency in the data layer, package payload or CTest target is introduced.
