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
- Control Name, ControlType, Caption for supported label/button controls, typed Value for supported data controls,
  Visible/Enabled, SetFocus and bounded ComboBox/ListBox Requery. The latter
  executes table/SQL/saved-query row sources through the shared query adapter,
  uses the native multicolumn model and configured BoundColumn, and blocks
  synthetic dirty signals during reload.
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

Reports is now a separate project-scoped collection, with bounded read-only
record binding and Open/Close/NoData event support (REPORT_EVENTS.md).

Only named controls with a supported widget implementation appear in Controls;
unsupported Access controls are not fabricated. Default control-object coercion
and `Set ctl = Me!control` are incomplete; use Controls(name) for explicit object
references. Full property/event namespaces, dynamic control creation, module/class
lifetimes, custom form methods, native-designed forms, complete report processing, embedded subforms,
combo bound columns/multicolumn rendering, control arrays, timers and Access
window-activation semantics remain open. Requery on unsupported control kinds
raises an explicit error. No COM/ActiveX implementation is implied.

Print/PDF and the scoped Application.CommandBars adapter are documented in
PRINT_PREVIEW_ADAPTER.md. They remain distinct from a complete Reports engine.

## Value versus object reference contexts

The evaluator distinguishes scalar reads from object-reference reads for Set and
property-assignment owners. `Me.control.Enabled = ...` must retain the control
object rather than try to assign through its current scalar Value. Both dot and
bang paths, parenthesized references and Forms collection paths are covered.
Reserved form properties such as Name/Caption keep their property meaning;
Set on a scalar raises error 13 and an unset assignment owner raises error 91.

`Me("control" & suffix)` resolves the actual Controls collection item; explicit
properties such as Caption/Visible/Value work on that reference. Implicit scalar
coercion of arbitrary objects and all default-member signatures remain open.
DAO bang references similarly retain a Field object for Set or a following
.Value, while ordinary `rs!field` reads keep their scalar behavior. Collisions
with Recordset EOF/Fields properties are exercised separately.

A hash-verified original Form_Load initializer now runs automatically and after
resetting actual Qt state: six Enabled/Visible assignments are checked. The
original module and form bytes are unchanged. This is control/event execution;
it does not claim that this form's complete record-source or business workflow
is compatible. The other two original Form_Load handlers remain separate work.

DoCmd.Requery ControlName resolves the supported control on the active native
form; Me.control.Requery resolves the caller's instance. A bound non-list data
control reloads the form record source after normal save/validation. Native
regressions check external SQL changes, separate caller/target combo models and
BeforeUpdate cancellation (2501). Unsupported control types still raise errors.

## Native text selection

SelStart, SelLength and SelText use the actual focused QLineEdit or editable
combo editor. Selection changes do not dirty records; SelText replaces the
selection (or inserts at the cursor), invokes the native Change path and can
be saved through normal form validation. No selection returns Null for SelText.
Unfocused access raises 2185, unsupported control types 438, negative positions
5, Null assignment 94 and Integer overflow 6. Large valid positions/lengths are
clamped to the native text bounds. UTF-16 selection is exercised with a surrogate
pair, replacement and independently read SQL values.

This native single-line selection adapter shares the bounded Text/Value buffer
in TEXT_VALUE_RUNTIME.md; formatted/multiline text, IME and rich-text parity
remain open. The original VBA
fixtures contain no selection-property calls, so evidence is the documented
Microsoft selection contract and native Qt/VBA/SQL tests, not original-handler
execution or a Windows Access reference comparison.

Focused Text and bounded Text/Value buffering are now implemented for native
text/ editable combo editors; see TEXT_VALUE_RUNTIME.md for update phases,
SQL/validation regressions and the remaining focus/input/reentrancy limits.

Bound OldValue, native text/combo Undo and BeforeUpdate self-change guards are
documented in TEXT_VALUE_RUNTIME.md. Original record values persist through the
form AfterUpdate callback; supported field conversion limits are inherited.
