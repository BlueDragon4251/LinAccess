# Native form record fields without controls

Forms expose actual loaded record-source fields even when no control is bound to
that field. A field reference uses its actual canonical column name and the
owning form's lifetime; it exposes Name, Value and Parent. It is a native form
field reference, with no fabricated DAO metadata or widget. Controls retain
priority on a name collision. A field which already has a bound control uses
that actual control and its existing shared value path; Controls.Count is
unchanged.

Supported direct-table form fields have a typed per-record draft buffer. Reading
through Forms!Name!Field, calculated controls and the field's default Value sees
the draft. VBA assignment and SetValue use the same checked setter. Conversion
succeeds before the buffer changes, and scalar object evaluation rechecks owner
liveness. Generated AutoNumber fields refuse2448. Reports and noneditable query
sources gain no additional write permission.

SaveRecord merges these fields into the existing parameterized record write;
other bound controls are saved together. Native SQL constraints remain active.
Failed saves retain the draft. Form.Undo and record loading clear it, while
Control.Undo retains unrelated record-field drafts and Dirty. New records read
actual table literal defaults for fields without controls; explicitly assigned
values participate in the same insert and generated identity handling. The
existing record validation and close/navigation paths apply.

This does not implement writable joins, full DAO Field metadata, arbitrary
RecordSource mutation, report-field writes, or complete Access validation/event
and locale semantics. The fixed456-point catalog is unchanged.

The existing mainwindow_documents check passes (8.50s). Actual VBA and both
macro hosts change real SQL-backed fields, preserve the original SQL values
until save, merge a visible aliased control, recalculate from the hidden draft,
retain a failed CHECK-constrained draft, undo without writing, insert a new row
from actual defaults and refuse AutoNumber/readonly/stale writes. Control.Undo
uses actual keyboard input and preserves the independent hidden-field draft.
The existing layout check passes (11.32s). No new CTest target is added.
