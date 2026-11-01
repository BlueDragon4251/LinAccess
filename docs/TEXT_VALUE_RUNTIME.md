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
the first input, failed focus-transition details, arbitrary reentrant control updates, broader Undo events/control kinds,
formatted/multiline/rich Text, IME, combo auto-expansion/LimitToList/NotInList and
complete popup/validation ordering remain open. Imported field-only references,
non-direct writable sources and every keyboard path are not complete Access
parity. No original full application workflow or visual parity is claimed.

## OldValue and control Undo

Bound data controls expose read-only OldValue from the original record buffer.
It retains the previous field value during control update and the form's
AfterUpdate handler, then advances after successful record save. New unsaved
records expose Null. It uses the existing field converter: supported numeric,
Boolean, text and Null representation is inherited; this does not add exact
Currency/Decimal/Date Variant storage. Unbound/calculated OldValue is explicitly
unsupported (3251), reports are rejected (438), assignment raises 383.

TextBox/editable ComboBox Undo restores the pending control-edit baseline,
synchronizes its bound mirrors and recalculates fields without Change. Other
record edits remain dirty. It works inside Change or BeforeUpdate; an undone
control update does not dispatch AfterUpdate. The same operation backs the
first Escape action. Already updated controls have no pending edit to undo.
Other control-kind Undo adapters/events remain unsupported.

Self-modification of Value/Text/SelText inside the same control's BeforeUpdate
raises 2115, including a focus-loss update. Normal AfterUpdate Value assignment
remains allowed. A reentrant same-control commit is stopped; drafts created by
AfterUpdate are retained and record save cancels with 2501 for explicit retry,
rather than saving the earlier value and silently discarding newer input. This
is a bounded recovery path, not general Access reentrant-update parity.

Native tests cover real SQL writes, control/form event traces, prior-value
visibility through form AfterUpdate, text/numeric/Null fields, programmatic
changes, preservation of other fields, Change/BeforeUpdate Undo, both Escape
levels, focus-loss rejection, AfterUpdate assignment and draft retry. Original
restored modules contain no OldValue or control-qualified Undo tokens; no
additional original-handler execution is claimed. Reference contracts:
[OldValue](https://learn.microsoft.com/en-us/office/vba/api/access.textbox.oldvalue),
[Control.Undo](https://learn.microsoft.com/en-us/office/vba/api/access.control.undo),
[AfterUpdate](https://learn.microsoft.com/en-us/office/vba/api/access.textbox.afterupdate-event).
