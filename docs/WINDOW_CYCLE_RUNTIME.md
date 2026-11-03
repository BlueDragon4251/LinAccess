# Native project window cycling

Access 2003 documents Ctrl+F6 for the next database window and Ctrl+Shift+F6 for
the previous one. These keys now cycle the real main database window and its
open native table tabs and visible native Form/Report windows. The bounded ring
uses the database window first, table-tab order next, then native document
creation order, wrapping in both directions. Hidden,
closed, embedded, modal and other-project documents are excluded. Minimized
eligible documents restore when selected. Returning to a document restores its
previous Qt focus widget, and database-window selection uses the existing
explorer focus/splitter/command-context path.

Native keyboard dispatch runs first, including KeyPreview cancellation. Current
control validation can cancel the transition; cycling does not call record save.
A native execution-depth guard keeps closed windows alive through callbacks.
Targets are weakly held and rechecked after validation. A handler redirecting
activation, closing the selected target or replacing/closing the project
prevents a stale activation. An active modal widget blocks both directions.

The production-main-window regression uses actual temporary SQL projects,
keyboard events and form/report widgets. It checks both directions/wraparound,
real focus restoration, KeyCode=0 and BeforeUpdate cancellation, hidden skip,
minimized restoration, cross-project exclusion, removal of closed windows and
a target report closed during control validation. Actual SQL stays unchanged
while form edits remain dirty. Modal blocking is also checked. Initial targeted
main-window regression: 1/1 (1.31 s); complete validation is in PROGRESS.md.

This is a bounded native-document ring, not a claim of Windows MDI stacking,
Access's exact window-order policy, query/editor-window cycling, popup
semantics, all multi-instance cases or real window-manager compatibility.
No extra original business execution is claimed.

Reference: archived Access 2003 keyboard help,
https://documentation.help/MS-Office-Access-2003/AccessKeyboardShortcuts.htm


## Real table pages in the window menu/ring (2026-11-03)

The Window menu lists actual open table pages alongside native forms/reports,
with numbered entries and the current table marked. Selecting a table activates
its existing Qt view. Ctrl+F6/Ctrl+Shift+F6 from either native forms or the main
window use the same ring when a production MainWindow owns the project. Other
hosts retain the bounded native-form fallback. Table focus retires the preceding
active form context so a document command cannot silently address its old caller.

MainWindow tracks the current data page and validates its actual delegate editor
before menu/ring/F11 switches. Direct tab-bar mouse navigation is checked before
hiding the old page; tab navigation keys are checked without treating bare
modifier keys as a commit request. Existing row drafts remain manual-submit
caches, with no added SQL save. Form validation/cancellation and active-modal
blocking retain their existing paths. Editor focus registration uses the owning
main window's event filter rather than a global application signal.

The existing main-window target exercises two real table pages, menu selection,
both key directions and the database window, blocked number input via Ctrl+F6,
menu and tab click, Escape, and a mixed native form/table round trip. Targeted
main-window/Currency/layout checks pass 3/3 (16.31 s). Full rebuilt desktop suite
passes 48/48 (92.38 s). Debian payload/permissions/Build-ID/pinned SQLite checks
and all four original hashes pass. Latest GUI-free core is 1/1 (1.53 s), with
subsequent Qt-only edits. This includes the preceding native mouse/context-menu
increments. Windows MDI stacking, exact creation/activation policy, all queries/
editors and complete Access window/focus semantics remain open.
