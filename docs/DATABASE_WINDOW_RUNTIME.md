# Classic F11 and database-window command context

The Access 2003 help defines F11 as bringing the Database window to the front.
The native equivalent activates the main project window and focuses its actual
object explorer, restoring the explorer if its splitter was collapsed. Repeated
F11 keeps it shown; this does not copy modern Access's Navigation Pane toggle.
The same main-window action handles F11 from its own widgets. Native forms and
reports route F11 after their existing VBA KeyDown/KeyPreview dispatch, avoiding
a global QAction shortcut that could preempt cancellation.

Nonmodal native documents commit pending control input through existing control
validation before switching. Cancelled validation or a KeyCode=0 handler retains
the editor; imported modal windows do not switch. F11 does not call form-record
save. Actual VBA control/deactivation handlers can of course write data. The
explorer regains real Qt focus and the previous native document receives one
logical Deactivate transition. Returning to a native document restores its
logical activation and Screen references.

Database-window focus is now an explicit weakly owned UI context, rather than
implicitly leaving the last form active. Screen.ActiveForm/ActiveControl report
the existing native unavailable errors 2450/2474. Default record commands,
GoToControl and active-control requery cannot silently target a stale form;
record-command context errors are 2046. Database-window default Close and
SendKeys targets remain unsupported (3251), with asynchronous SendKeys target
resolution also refusing stale form focus. Named object operations and explicit
instance/DAO operations keep their existing scopes. In particular Me.Dirty=False
and CurrentDb operations can act on their explicit owner/project while the
explorer is active.

The background-context guard starts after a document's first activation,
preserving the prior bounded early-Current initialization path. This does not
establish complete Windows initialization/focus/error timing. Native message
boxes and external-application deactivation retain the existing document context;
they are not treated as a database-window transition.

## Evidence

The actual production-main-window integration on temporary SQLite projects tests
F11 from a dirty editor, restored splitter width and actual explorer/main-window
focus, repeated non-toggle behavior, KeyPreview and BeforeUpdate cancellation,
modal blocking, one Deactivate transition and form reactivation. SQL rows remain
unchanged after default background save/key/close attempts; native errors are
checked, and an explicit Me.Dirty=False persists the expected value. Screen
references clear/restore, and an early Form_Current command still initializes.
Existing project-switch cancellation and retained close/MsgBox checks continue
to pass. Final targeted trio: 3/3 (39.79 s). Full validation is in PROGRESS.md.

The original 32-module lexical inventory has zero vbKeyF11 and zero SelectObject
references. No extra original business execution is claimed.

## References and limits

- Archived Access 2003 product help: https://documentation.help/MS-Office-Access-2003/AccessKeyboardShortcuts.htm
- https://learn.microsoft.com/en-us/office/vba/api/access.screen.activeform
- https://learn.microsoft.com/en-us/office/vba/api/access.form.deactivate
- Modern reference, whose F11 toggle is deliberately not the 2003 contract:
  https://support.microsoft.com/en-us/office/keyboard-shortcuts-for-access-70a673e4-4f7b-4300-b8e5-3320fa6606e2

Tests use Qt offscreen. Real window-manager stacking, complete MDI/database-window
commands, popup/menu/IME focus, every shortcut and exact Access error/order
identity remain open. This does not implement schema-object deletion, database-
window SendKeys, default database-window Close or Alt+F11/VBE. Bounded
SelectObject support is now documented in SELECT_OBJECT_RUNTIME.md.
