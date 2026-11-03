# Native command button keyboard policies

Imported CommandButton.Default and Cancel (including type templates) configure
native button keyboard policies. Enter uses a focused command button first,
otherwise the enabled/visible default after control validation. Escape invokes
the enabled/visible cancel button's actual Click binding; it first undoes the
focused supported TextBox buffer. Without a cancel button, the existing Escape
undo behavior remains. A cancel button without a closing handler does not close
the form automatically.

VBA can read/write Default and Cancel. Setting one button True clears that
policy on the other buttons; both policies can belong to the same button.
Boolean conversion is checked, Null rejected. Unavailable native command-button
widgets fail explicitly instead of dereferencing a placeholder.

Existing native GUI regression sends real Enter/Escape/Alt keys, executes VBA
policy switching and verifies SQL Click effects, discarded textbox input and
focused-button priority. Desktop build and form_current_record pass (0.56 s).
No new target. Full multiline/subform focus, duplicate mnemonics, every control
validation/Windows event order and Access 2003 rendering remain open.

References:
- https://learn.microsoft.com/en-us/office/vba/api/access.commandbutton.default
- https://learn.microsoft.com/en-us/office/vba/api/access.commandbutton.cancel
