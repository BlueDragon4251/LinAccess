# Native Form.AllowEdits and click policy

## Implemented scope

SaveAsText AllowEdits defaults to true; NotDefault decodes to false. Native form
VBA getters/setters expose the Boolean and preserve it in instance state. Null
raises 94, invalid scalar conversion 13, report access 438 and changes during
record commands 3251 before mutation. Object scalar defaults resolve once.

On a clean existing record, AllowEdits=false blocks user edits to bound controls.
Text boxes remain enabled/focusable for reading/copy; editable combo editors and
popup/navigation input are read-only, and checkbox/option-group state changes
are suppressed. Unbound search fields remain editable. New-record input follows
AllowAdditions independently. Existing-record deletion commands require both
AllowEdits and AllowDeletions; blocked deletion stops with 2046 before SQL or a
confirmation dialog. Direct DAO/SQL operations retain their separate contracts.

Per the Microsoft contract, a programmatic bound Value assignment dirties the
current record and permits its user editing even while AllowEdits remains false.
Saving or record Undo restores the restriction; cancelled validation retains the
editable draft. Changing AllowEdits while dirty neither saves nor discards the
record. Backend writability, Locked, calculated controls, Enabled and additions
remain independent constraints. Record loading clears dirty state before applying
input policy, so Undo/navigation do not retain the preceding record's exemption.

Native standalone checkbox Click bindings now execute after supported control
update handling; a read-only checkbox can still run its Click code without a
user state change. That code can perform a programmatic Value override. Click
on an option-group member dispatches only the group's binding, as Microsoft
documents; its individual OnClick is not fired. An older synthetic regression
incorrectly expected a child event and now checks the group event instead.
Empty bindings are absence of handlers, not unsupported macro expressions.
Callbacks retain document lifetime across nested VBA/dialog execution and stop
following events once a document closes. Reports remain noninteractive previews
for these input/click paths.

## Evidence

Actual native Qt user input and SQLite checks cover blocked existing-record text/
combo/checkbox/group changes, unbound search, an unbound checkbox switching the
policy through VBA, checkbox/group Click code overriding values, saving/restoring
restrictions, existing dirty input during policy change, record validation cancel,
Undo, navigation, new-row insertion, disabled additions, font changes without a
data exemption, Locked independence, Null/type/scalar preflight and report 438.
SQL trace rows prove the group event runs and its child event does not. Changing
policy in Form_Delete fails 3251 before mutation/deletion. A dynamically opened
checkbox form closes from AfterUpdate, enters a real MsgBox, survives its nested
loop, then is deleted. Its row saves once and a following Click produces no SQL.
Targeted layout/original-heavy regressions pass (37.28 s).

All five hash-verified original forms declaring disabled edits expose false
through the native property, without supplied CodeBehind handlers or data edits.
Their current native query sources already remain read-only: zero writable-source
policy transitions were observed. Thus the original probe establishes flag/lookup
retention, not original writable-query or business-workflow compatibility. The
actual editable-to-read-only SQL transitions above use synthetic writable tables.
No additional original business-handler execution count is claimed.

References: [Form.AllowEdits](https://learn.microsoft.com/en-us/office/vba/api/access.form.allowedits),
[CheckBox.Click](https://learn.microsoft.com/en-us/office/vba/api/access.checkbox.click).
These document the programmatic dirty-record exception and group-only Click.
No Windows Access 2003 runtime differential was available.

## Remaining scope

Full event ordering, mouse/focus/keyboard parity, programmatic Text/selection
editing on read-only editors, Dirty=true, original updatable query sources,
DataEntry/session filtering, OpenForm DataMode overrides, RecordsetType/locking
and arbitrary reentrancy remain open. This is input policy, not a full Access
record lifecycle or database permission implementation.

Validation: host 21/21 (83.77 s), Debian 12 21/21 (68.34 s), Ubuntu 24.04
21/21 (62.50 s), GUI-disabled core 1/1 (0.88 s). Fresh host/distro DEBs pass
exact payload/backend/permissions/fixture-exclusion checks; isolated installed
Debian/Ubuntu packages pass backend/counter execution and GUI startup. All four
original fixture hashes pass. Five original flag-retention probes and zero
original writable-source policy transitions are recorded separately.
