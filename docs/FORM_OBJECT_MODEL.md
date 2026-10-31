# Imported form object model

This is an executable C++/Qt subset for imported AccessFormPreview dialogs, not a
complete Access class engine. Form/control references target actual widgets and
SQLite project records. Original MDBs are never opened for writing.

Supported paths:

- `Set frm = Me`, `Forms(name/index)`, `Forms!name`, `Forms.Item(...)`,
  `Forms.Count` and For Each over open imported forms in opening order.
- The Forms collection is restricted to the same QSqlDatabase connection. Other
  project dialogs are excluded. Closed/cancelled forms are excluded.
- Form `Name`, `RecordSource`, `Dirty`, `NewRecord`, `Requery`, `SetFocus` and
  named control-value reads/writes. Name/RecordSource/NewRecord setters currently
  fail explicitly; dynamic record-source changes remain open.
- `Me.Controls`, `Me.Controls(name/index)`, Controls.Item/Count and enumeration
  of instantiated named controls, including explicitly hidden controls.
- Control Name, ControlType, typed Value for supported data controls,
  Visible/Enabled, SetFocus and bounded ComboBox/ListBox Requery. The latter
  executes table/SQL/saved-query row sources through the shared query adapter,
  uses the first result column, and blocks synthetic dirty signals during reload.
- Access acLabel/acTextBox/acCommandButton/acComboBox/acListBox/acCheckBox/
  acOptionButton/acToggleButton/acSubform/acRectangle/acLine constants.

Live wrapper identities are cached weakly so `Is` compares the same form/control
consistently without retaining the Qt dialog. QPointer guards destroyed targets;
closed or destroyed targets fail with error 2450. No stale raw widget pointer is
used after destruction. A Forms wrapper is currently anchored to its originating
form context; application-global/module-lived collection ownership remains open.

Bang and dot lookup are distinct: `Forms!example!Name` reads the control, while
`Forms!example.Name` reads the form property. `Me!Name` and `Me.Name` follow the
same separation in imported form hosts. DAO also separates `rs!EOF` from `rs.EOF`
and `rs!Fields` from `rs.Fields`. Assignment finds the final member separator
outside quoted arguments/bracket names, including controls named `Name.Field`.

## Evidence

The GUI regression clicks synthetic VBA handlers to enumerate controls with
Access constants, compares form/control identity, writes and saves a different
open form, verifies real SQLite changes, tests bang/property name collisions,
updates row-source data and requeries a real combo, and checks closed/deleted
references and separation between project connections.

All 47 original Kassa form artifacts are instantiated without their original VBA
executing during this structure test. The collection Count, names and control
types of their supported instantiated widgets are inspected. Original handler
execution remains separately measured as eleven keypad buttons/helpers and one
DAO insertion button; this does not demonstrate the complete application.

## Remaining limits

Only named controls with a supported widget implementation appear in Controls;
unsupported Access controls are not fabricated. Default control-object coercion
and `Set ctl = Me!control` are incomplete; use Controls(name) for explicit object
references. Full property/event namespaces, dynamic control creation, module/class
lifetimes, custom form methods, native-designed forms, Reports, embedded subforms,
combo bound columns/multicolumn rendering, control arrays, timers and Access
window-activation semantics remain open. Requery on unsupported control kinds
raises an explicit error. No COM/ActiveX implementation is implied.
