# Imported control state and tab traversal

SaveAsText NotDefault means the opposite of the supported property's built-in
default: Visible/Enabled/TabStop/AllowAdditions/AllowDeletions become false;
Locked/Modal/KeyPreview/FilterOnLoad become true. Explicit Boolean words/numbers
are also parsed; invalid values fail. Control-default template blocks remain
separate from instantiated controls. TabIndex is validated as a nonnegative
Integer. Original source bytes are unchanged.

Native controls retain imported Enabled/Visible across record changes. VBA
assignments update the same instance model; the legacy Me.control.Visible path
now shares it instead of being overwritten by the next state update. A Null
visibility assignment raises 94. SetFocus uses the actual visible/enabled control
check (2110), including the legacy direct statement path.

TextBox Locked maps to QLineEdit read-only state while keeping focus/selection.
VBA Value assignments remain possible on writable sources and follow normal
SQL save rules. Locked on other kinds is explicitly unsupported (3251); imported
locked unsupported kinds show a diagnostic instead of an editable substitute.

Read-only data sources preserve configured Enabled. Line/combo editors prevent
input; checkbox/radio/checkable-button transitions are suppressed without
pretending that Enabled was set false. Read-only combo dropdown attempts are
rejected (2448). Source writability and designer Locked/Enabled are distinct;
this does not implement every Access read-only-control interaction.

TabStop false removes native Tab focus while retaining click/programmatic focus.
Exported TabIndex orders controls within header/detail/footer scopes. Native
Tab/Shift-Tab skips hidden, disabled and TabStop-false controls. Dynamic TabStop
changes the live focus policy. Dynamic TabIndex in a flat section moves a real
control and renumbers other positions without dirtying SQL data; it validates
0..count-1 (5), Integer overflow (6) and Null (94). Nested groups, tabs, subforms
and custom controls reject dynamic reorder with 3251. Duplicate/default indices
retain stable source order; complete LoadFromText index normalization is open.

## Evidence and corrected original boundaries

Native Qt/VBA/SQL tests verify imported flags, both tab directions independent
of construction order, live reorder/renumber, error preflight, unchanged focus
and record data, locked-user input versus programmatic writes, persisted states
across navigation and read-only checkbox/combo/option-group input protection.
A hash-verified original flat layout is traversed in both directions with original
metadata and no original VBA execution. The unchanged original initializer still
performs its six actual Enabled/Visible assignments.

The 47 original forms retain 879 parsed controls; 81 command buttons are declared
visible before VBA, correcting the earlier count of 86 that ignored NotDefault.
Hidden controls remain instantiated for VBA. The original calendar's exported
Enabled flag now makes the native control disabled. Its selection regression
first verifies this state, then explicitly enables it through the actual property
adapter. CALENDAR_RUNTIME.md qualifies those original event tests accordingly.
Windows OCX activation/persistence has not been compared with an Access oracle.

## Remaining scope

This is not full Access keyboard or visual parity. General nested/page cycle behavior (bounded flat Cycle is now implemented in
FORM_CYCLE_RUNTIME.md), mouse focus validation, Enter/Shift-Enter behavior,
label access keys, container/subform tab scopes, disabling a focused control,
all Locked kinds, default-property inheritance and exact Windows focus/rendering
traces remain open. Qt rendering is not an Access 2003 pixel reference.

References: [TabIndex](https://learn.microsoft.com/en-us/office/vba/api/access.textbox.tabindex),
[TabStop](https://learn.microsoft.com/en-us/office/vba/api/access.textbox.tabstop),
[Locked](https://learn.microsoft.com/en-us/office/vba/api/access.textbox.locked).
