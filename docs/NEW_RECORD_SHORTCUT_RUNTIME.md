# Native new-record shortcut and navigation commands

Native forms now handle Ctrl+Plus (including keypad modifier normalization) and
Ctrl+Shift+Equals through the existing active-document navigation command. Native
SendKeys "^{+}" uses the same event path. KeyPreview can cancel the command.
A pending field is validated/saved before a permitted move; cancelled validation
retains the draft and current record. Disallowed additions/read-only or unavailable
new-record sources reject before saving. Entering the empty new row does not
INSERT a record. Existing navigation/default/undo/lifetime contracts apply.

DoCmd.RunCommand now supports the documented AcCommand navigation constants:
New=28, Next=65, Previous=66, First=67, Last=68. Common unqualified constant names
are available in VBA. Incorrect assumed aliases 24-27 are explicitly not accepted.
All commands use the existing bounded GoToRecord implementation and active native
document context; this is not full navigation/source/MDI/query parity. Exact
Windows key-layout/keypad event codes and complete keyboard behavior remain open.

The fixture-independent Qt/VBA/SQL test verifies KeyPreview cancellation, disabled
additions with both clean and pending records, validation rejection/retry, actual
save of the prior record before allowed navigation, SendKeys, both key variants,
new-row/undo without INSERT and all five native RunCommand actions. Rejected
wrong command numbers preserve position/data. Targeted core/arithmetic/shortcut
tests: 3/3 (1.58 s). Full validation is recorded in PROGRESS.md. Original business
execution coverage is unchanged.

References:
- https://support.microsoft.com/en-us/office/keyboard-shortcuts-for-access-70a673e4-4f7b-4300-b8e5-3320fa6606e2
- https://learn.microsoft.com/en-us/office/vba/api/access.accommand
