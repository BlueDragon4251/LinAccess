# Native project window cycling

Access 2003 documents Ctrl+F6 for the next database window and Ctrl+Shift+F6 for
the previous one. These keys now cycle the real main database window and its
visible native Form/Report windows. The bounded ring uses database window first,
then native document creation order, wrapping in both directions. Hidden,
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
Access's exact window-order policy, table/query/editor-window cycling, popup
semantics, all multi-instance cases or real window-manager compatibility.
No extra original business execution is claimed.

Reference: archived Access 2003 keyboard help,
https://documentation.help/MS-Office-Access-2003/AccessKeyboardShortcuts.htm
