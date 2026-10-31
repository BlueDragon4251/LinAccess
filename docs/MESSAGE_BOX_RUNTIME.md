# Native VBA MsgBox and form-member resolution

The C++ runtime executes MsgBox as a function, a statement or `Call MsgBox(...)`.
The Qt host displays an application-modal QMessageBox and returns the actual
selected button as VBA values 1 through 7. OK, OK/Cancel, Abort/Retry/Ignore,
Yes/No/Cancel, Yes/No and Retry/Cancel groups, four icons and available default
buttons are supported. Escape/window close cannot dismiss groups without Cancel
(except OK-only); Cancel groups return vbCancel. Prompt text is plain text.

Positional omitted buttons/title are distinguished from an explicitly empty title.
The default title is LinAccess. The runtime checks argument counts, Null arguments,
Long conversion/overflow and returned button values. Windows system-modal/Help
styles and Helpfile/Context arguments fail explicitly with 3251. Named arguments,
Windows Help integration, complete integer subtype identity and a Windows Access
visual/behavioral oracle remain open. Right-to-left/right-aligned text and the
foreground request are translated to native Qt options; compositor behavior varies.
An invalid/unavailable default button currently raises 5 rather than claiming an
unverified Windows fallback. A headless host must implement styled dialogs; it
cannot manufacture Yes/No responses through a logging-only message callback.

The original Kassa module with SHA-256
`fa72dad21f4875398da7de601750d963c1d5d1292a36f3db6fbc240ed66959ba`
executes unchanged in a native control fixture. Its missing-input branch displays
the original message and focuses the actual input widget after OK. The test uses
a generic non-credential selection and a Null second input, not authentication
records. It does not claim the full login workflow or original form layout passes.
All original bytes are hash checked and neither messages nor input values are logged.

That regression exposed absent resolution of unqualified form controls. The native
host now resolves existing control names as values or object references, and bare
assignments update those controls. Existing local variables and procedure parameters
have priority. Unknown reads still fail; controls are not invented. This is a bounded
class-member bridge, not complete module/class scope, Option Explicit enforcement,
forward declaration binding, typed assignment or a replacement for the planned AST.

`tests/message_box_tests.cpp` verifies native button results, SQL branching on
vbYes, statement/Call forms, omitted arguments, plain-text display, icons/defaults,
non-dismissable Yes/No, local/parameter shadowing, object references and the original
handler. `tests/core_tests.cpp` additionally checks headless delegation, optional
arguments and error cases. Normal DAO SQL execution adds no LinAccess consent.
DoCmd.SetWarnings and complete Access action-query warning semantics remain open.

Validation for this stage: all 18 CTests passed (54.38 seconds), the separate
GUI-disabled core build/test passed, and a fresh Debian package passed content,
permission and executable Build ID checks. All four original fixture SHA-256 checks
passed unchanged. This is local validation, not a claim that remote distribution
CI or Microsoft Access reference runs have completed.
