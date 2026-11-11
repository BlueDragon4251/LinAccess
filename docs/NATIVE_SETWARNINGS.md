# Native SetWarnings state and record confirmation

DoCmd.SetWarnings(WarningsOn) is available through actual standard-module and
form DoCmd objects, including named argument binding and Variant Boolean coercion.
Arguments are converted and owner lifetime checked before application state
changes. Invalid arity/Null do not change it. The application defaults to warnings
on. VBA changes persist across procedure errors until explicitly changed again.

A native macro SetWarnings action shares that state. An invocation that executed
this action restores warnings on when it exits, including StopMacro/StopAllMacros
and errors. Nested/repeated invocation equivalence to Windows has not been measured.
Record-delete confirmation uses the flag while retaining native deletion events,
Cancel and transaction/rollback paths. Explicit VBA MsgBox/InputBox and runtime
errors retain their existing paths. The project requirement of no additional
per-SQL consent prompts is preserved.

Contract sources (read2026-11-11):
- https://learn.microsoft.com/en-us/office/vba/api/access.docmd.setwarnings
- https://learn.microsoft.com/en-us/office/client-developer/access/desktop-database-reference/setwarnings-macro-action

The docs distinguish persistent VBA state from automatic macro restoration.
This is a bounded native implementation, not proof of all Access system-dialog,
GetOption/confirmation preferences, nested macro or Windows2003 equivalence.
F059 remains open; fixed456/386/70. App/helper/core/GUI built; existing core/GUI bundle2/2 in19.11s
(0.52/18.58) passed. Actual named VBA call, persistent state after Err.Raise,
invalid arguments without mutation, real record deletion, macro completion/stop/
error restoration, RunSQL without extra consent and explicit MsgBox are covered.
All seven protected hashes unchanged. No new CTest ID or Windows oracle.
