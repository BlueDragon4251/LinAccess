# Native Text/Value editing buffer

The TextBox and editable ComboBox adapters expose Text on their actual focused
Qt line editor. Unfocused reads/writes raise 2185; unsupported control types
raise 438, Null Text assignments 94 and read-only editor writes 2046. Text
assignment replaces native editor contents and runs the real Change path.

User typing and SelText/Text edits retain the prior typed Value until control
update. Change sees the current Text and prior Value. Bound mirrors and
calculated controls continue to use Value. Control BeforeUpdate sees the new
value; Cancel/error retains the draft for correction or retry. After successful
validation, Value/mirrors/calculations update and AfterUpdate runs. Focus-loss
callbacks retain Text visibility through this control-update phase despite Qt's
earlier native focus transition. Record saving separately performs the existing
form validation and actual SQL write.

Programmatic Value assignments replace pending editor input without Change or
control update events, refresh the edit baseline and preserve normal record
dirty/save behavior. Escape restores the current control edit baseline; form
Undo reloads the record. Null new-record values remain Null while text is being
edited. Initial combo row-source loading preserves a modified draft separately
from its prior bound Value.

## Evidence

Native Qt/VBA/SQL regressions check distinct Text/Value visibility during Change
and BeforeUpdate/AfterUpdate, unchanged SQL until record save, bound mirrors and
calculated Currency fields, cancellation/retry, programmatic overrides including
Error Variants, Escape/form Undo, focus errors, new rows, editable combos and
finite cascading Change handlers. Existing real-fixture regressions remain
mandatory. No direct Text-property member calls occur in the restored original
VBA token inventory; no additional original-handler execution is claimed.

The behavior is based on Microsoft's documented
[TextBox.Text](https://learn.microsoft.com/en-us/office/vba/api/access.textbox.text)
and [BeforeUpdate](https://learn.microsoft.com/en-us/office/vba/api/access.textbox.beforeupdate-event)
contracts, with native regression evidence. A Windows Access reference trace
has not established exact event timing.

## Remaining scope

This is a bounded single-line editing adapter. Full BeforeInsert timing before
the first input, failed focus-transition details, self-modification restrictions
inside BeforeUpdate, arbitrary reentrant control updates, OldValue/control Undo,
formatted/multiline/rich Text, IME, combo auto-expansion/LimitToList/NotInList and
complete popup/validation ordering remain open. Imported field-only references,
non-direct writable sources and every keyboard path are not complete Access
parity. No original full application workflow or visual parity is claimed.
