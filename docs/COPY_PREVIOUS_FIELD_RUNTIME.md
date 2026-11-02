# Native Ctrl+apostrophe previous-field copy

Ctrl+apostrophe now copies the preceding displayed record's same bound column
into a focused writable native TextBox. New records use the last displayed row.
The native action replaces the field's edit text, retains focus and the original
committed Value until validation, and does not use or change the clipboard.
Form/control KeyPreview dispatch can suppress the key before the action.

The action uses the existing user edit buffer and native Dirty/Change/BeforeUpdate/
AfterUpdate/Undo/save paths. NULL and empty text remain distinct. Identical values
are a no-op; no preceding row, read-only/non-TextBox/unsupported sources reject
without changes. Binary values and text exceeding the editor capacity explicitly
reject rather than converting/truncating. Native execution lifetime guards protect
callback processing. Record navigation/filter ordering uses the existing record
snapshot; full concurrent refresh/source/type/format semantics remain open.

The fixture-independent copy_previous_field CTest uses actual Qt key events,
focused controls, native VBA handlers and SQL records. It verifies Unicode and
apostrophes, old Value versus pending Text, explicit save, exact Change/Before/
Before/After sequence across cancelled/retried validation, KeyPreview cancellation,
read-only/capacity/binary guards, NULL/empty persistence, new-row copy without
implicit INSERT, Escape undo and the same path through SendKeys "^'". Targeted
regression: 1/1 (0.26 s). The complete GUI suite now has 24 CTests. Full validation
is recorded in PROGRESS.md. Original business execution coverage is unchanged.

This is bounded field-value replacement in writable native TextBoxes. Exact
Access selection/edit-mode insertion behavior, native Windows event/error timing,
ComboBox/ListBox adapters, all masks/types, query updates and complete keyboard
parity remain open. These are not implied by the shortcut's presence.

Reference: https://support.microsoft.com/en-us/office/keyboard-shortcuts-for-access-70a673e4-4f7b-4300-b8e5-3320fa6606e2
Microsoft describes Ctrl+apostrophe as inserting the value from the same field
in the previous record; exact Access 2003 mode/selection behavior still needs
a Windows comparison.
