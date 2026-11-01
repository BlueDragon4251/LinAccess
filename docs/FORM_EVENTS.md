# Form record lifecycle subset

The Access layout reader now retains form and control event properties. A handler
is invoked only for an exported `[Event Procedure]` binding; merely having a
same-named Sub in a module is not an event binding.

Implemented and checked with actual Qt keyboard edits and SQLite state assertions:

- Bound forms initially select the first existing record. Open, Load and Current
  run in order; Cancel in Open prevents Load/Current.
- The first bound edit marks the record Dirty and invokes Form_Dirty(Cancel).
  A cancelled Dirty event restores the previous record values.
- Control BeforeUpdate(Cancel) and AfterUpdate run at commit. Validation failures
  retain the edit and prevent a record save/navigation. Unbound/programmatic
  assignments do not fabricate control AfterUpdate events.
- Form BeforeUpdate(Cancel) runs before SQL persistence, Form AfterUpdate after a
  successful write. Cancel preserves unsaved values and leaves the database
  unchanged. A successful save of a new record selects its assigned identity.
- Record navigation saves pending changes or restores the old selection on
  validation cancellation, then runs Current for a completed move.
- Controls bound to the same field stay synchronized. Loaded NULL and empty text
  are kept distinct. A missing/externally deleted record is not reported saved.
- Unload(Cancel) and Close are honored by window close, accept and reject paths.
  Pending changes save before Unload. Cancelling Unload leaves the window open.
- CurrentDb.Execute no longer implicitly requeries a form and discards its edits.
  Explicit Requery validates/saves pending changes first, reloads filtered data
  and dispatches Current for a bound record source. Unbound layout refresh does
  not recursively manufacture Current. See [native filters](NATIVE_FILTERS.md).

Regression coverage lives in tests/form_event_tests.cpp and is included in the
real-layout runner. It exercises Qt text entry/focus commits, persisted values,
event trace order, duplicate bindings, cancelled controls/forms/Dirty/Open/Unload,
new-record identity and a handler deliberately lacking an OnClick binding.
Existing original keypad and DAO button regressions continue to run separately.
This is not an instrumented Microsoft Access reference trace.

Still open: complete BeforeInsert/AfterInsert/Delete sequences, key/mouse/timer
semantics, full activation/focus variants, nested form/report events, editable joins, general property
objects, module lifetimes, record locking/conflict resolution and virtualized
navigation beyond the explicit 10,000-row display limit. Query sources are not
yet generally editable and complete Access rendering/event compatibility is not
claimed.

## Stored event procedure prefixes

Control EventProcPrefix is preserved from SaveAsText and exposed as a read/write
VBA property. An explicit prefix selects `prefix_Event`, independently of the
current control Name, for bound Click/DblClick, KeyDown/KeyPress/KeyUp and control
BeforeUpdate/AfterUpdate. The native MSCAL source-interface Click lookup uses the
same prefix. Existing event-result keys and value/control lookup retain current
control names; a prefix does not create another control alias. Without an
explicit prefix, the existing generated control-name prefix remains in use.

The 47 original form layouts contain 68 stored prefixes. The layout regression
checks their parsed metadata and property values on rendered controls. Synthetic
native GUI/SQL regressions prove renamed-button dispatch, cancellation of
DblClick, keyboard order and ASCII replacement, actual record writes, validation
cancellation, VBA-driven prefix reassignment and default-prefix behavior. An
unbound event stays unbound, and an absent explicitly prefixed procedure remains
ProcedureMissing even if a procedure using the current Name exists.

The in-memory setter accepts 1–64-character letter/underscore-led identifiers
and rejects Null (94) or invalid identifiers (5) without changing the prior value.
Design-view source refactoring/persistence and exact Access 2003 edge-case name
validation are not established by these tests. Form/report lifecycle procedure
names remain their existing Form_/Report_ bindings. The two previously diagnosed
renamed original KeyPress controls have no EventProcPrefix override and remain
missing bindings; this change does not guess their older procedure names.

Property contract reference (current Microsoft documentation; no Access 2003
execution oracle): https://learn.microsoft.com/en-us/office/vba/api/access.textbox.eventprocprefix

## Native control Change increment

TextBox/ComboBox OnChange bindings now dispatch with the imported EventProcPrefix
from native edits. A native Qt/SQL trace verifies KeyDown → KeyPress → BeforeInsert
→ Change → KeyUp, an existing-record edit without BeforeInsert, a combo edit,
KeyPress cancellation and visible handler errors. A handler's VBA Value assignment
to another control does not fire its Change; loading/requery and synchronized
bound mirrors also suppress Change. Actual record persistence is checked.
Bounded focused Text/Value buffering and finite cascading Change are now tested
(TEXT_VALUE_RUNTIME.md). LimitToList/NotInList and complete Access input/locale
behavior remain open. See INSERT_EVENT_RUNTIME.md for pre-input timing limits.
