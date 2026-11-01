# Bounded native insertion events

The native form adapter distinguishes changes through Qt controls from VBA
field assignments. The first control change on a new direct-table record invokes
Form_BeforeInsert before Dirty. Cancel or handler failure reloads the unsaved row;
Undo inside BeforeInsert also cancels it. Programmatic VBA assignments do not
trigger BeforeInsert/AfterInsert. A BeforeInsert handler can initialize another
field without recursively firing insertion/dirty handlers.

Saving preserves the existing BeforeUpdate validation. A successful actual SQL
INSERT invokes AfterUpdate followed by AfterInsert for a user-started insertion.
A validation/constraint failure keeps dirty input and the insertion state for
retry, without firing AfterInsert. Undo/reload resets the insertion state.
Rejected combo/option/checkbox changes must not be recommitted by the remaining
activation/click callback.

Native Qt tests cover text editing and checkbox activation, exact event traces,
actual inserted values, handler initialization, Cancel, raised errors, Undo both
during and after BeforeInsert, programmatic assignment, validation retry and
UNIQUE failure followed by successful retry. No private source text is copied.

## Limits

BeforeInsert currently runs from native value-change handling: the editor has
already changed, before the adapter marks the record dirty. This does not prove
Access's precise BeforeInsert/Change/Text/Value visibility and input ordering; complete
Text/Value buffering and Access reference traces remain necessary. Native
OnChange now follows BeforeInsert and precedes KeyUp in the tested key path;
VBA Value assignment and bound-mirror updates do not fire Change. Complete
combo/popup/IME/context-menu timing, macros and keyboard cancellation interactions
are not established. The separate PasteAppend path still rejects insertion-event
bindings with 3251 before inserting; it does not reuse this UI event sequence.
Reports, non-direct record sources and general Jet form updateability remain
outside this insertion adapter. This is not full Access event parity.
