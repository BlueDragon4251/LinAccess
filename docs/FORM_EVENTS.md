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
  Explicit Requery validates/saves pending changes first.

Regression coverage lives in tests/form_event_tests.cpp and is included in the
real-layout runner. It exercises Qt text entry/focus commits, persisted values,
event trace order, duplicate bindings, cancelled controls/forms/Dirty/Open/Unload,
new-record identity and a handler deliberately lacking an OnClick binding.
Existing original keypad and DAO button regressions continue to run separately.
This is not an instrumented Microsoft Access reference trace.

Still open: complete BeforeInsert/AfterInsert/Delete sequences, key/mouse/timer
semantics, activation, nested form/report events, editable joins, general property
objects, module lifetimes, record locking/conflict resolution and virtualized
navigation beyond the explicit 10,000-row display limit. Query sources are not
yet generally editable and complete Access rendering/event compatibility is not
claimed.
