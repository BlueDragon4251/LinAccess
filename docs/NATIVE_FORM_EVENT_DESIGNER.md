# Native form event bindings

The native form designer's **Ereignisse zuordnen...** dialog connects supported
form and selected field events to actual saved macros without writing VBA.
Binding choices are automatic VBA discovery, a macro name, a function expression,
an explicit `[Event Procedure]`, or no binding. The macro picker lists actual
native project macros and accepts a name to be created later. A missing target
fails at execution rather than being invented. Function expressions use the
existing bounded native event-expression engine.

The subdialog edits copies. Cancel discards all its changes; OK validates and
returns them to the form designer draft. The parent form Save stores them.
Automatic leaves the generator's existing VBA procedure discovery in place;
no binding explicitly suppresses discovery for that event. Original VBA bytes
remain handled by the existing form editor preservation path.

Form events: Open/Load/Current/Activate/Deactivate/Unload/Close,
BeforeInsert/AfterInsert/BeforeUpdate/AfterUpdate/Dirty/Timer. Field events:
Click/DblClick/Change/BeforeUpdate/AfterUpdate/MouseDown/MouseUp. Existing automatic
keyboard procedure discovery remains separate; this dialog does not claim
keyboard macro arguments or complete event scheduling.

The native definition adds `events`, `field_events` and `timer_interval`.
Event names are validated and normalized; duplicate case-insensitive event or
field names, unknown events/fields and malformed text are refused. Bindings are
limited to4096 UTF16 units, with no NUL/invalid UTF16. Macro names retain their
existing64-unit/group/reserved-prefix bounds. TimerInterval is a checked signed
32-bit nonnegative millisecond count;0 disables delivery. Absent overrides preserve
legacy discovery and default timer behavior. No macro or expression is evaluated
while the definition is validated/saved.

Existing mainwindow_documents/access_2003_layouts2/2 (18.42s) and app build pass.
Actual nested dialog Cancel/reopen/OK, stored JSON/layout, Load SQL effects,
BeforeUpdate cancellation with unchanged source SQL, explicit suppression of an
existing Form_Load procedure and unknown event/field/negative timer refusals are
covered. In the same package macro lookup/save conflict checking uses bounded
Qt case-insensitive name matching, including Ä/ä; duplicate names fail before
macro effects. This is native matching, not full Jet/locale collation parity.
No new CTest target or Windows Access2003 observation. The earlier original corpus
run1/1 (114.46s) precedes this designer extension; no fresh full-suite or distribution
matrix is claimed. Full designer/macros/events/Access equality remains open under
F080/F088/F050/F057/F130/F148;456/384/72 unchanged.
