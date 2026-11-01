# Native form Cycle and validated Tab movement

Imported Cycle defaults to All Records (0); Current Record (1) is also supported
for standalone flat forms. VBA Me.Cycle and form-object property paths update
the live policy without dirtying SQL data. Null raises 94, invalid modes 5;
Current Page (2), nested/group/subform/custom-control modes reject dynamic
assignment with 3251. Reports reject the form property with 438.

Only actual native Tab/Shift-Tab key events enter this policy; programmatic
hide/disable focus changes do not trigger a record-cycle save. Tab and Shift-Tab
validate the edited control before transferring focus.
A cancelled/error control BeforeUpdate retains the draft and focus. Current
Record wraps among eligible fields in the same row: control values update but
record SQL is not saved just by wrapping. All Records validates/saves through
the existing native record command, moves to the next/previous record and
focuses its first/last eligible field. Record validation cancellation retains
focus and unsaved data. Hidden, disabled and TabStop-false fields are skipped;
existing tab positions remain ordered. Toolbar focus keeps its native path.

Forward traversal on an untouched empty new row wraps without inserting an
empty record. Backward traversal returns to the prior existing row. A completed
new row is really inserted before preparing the next empty row. Disabled
additions/missing previous records stop with native navigation errors (2046/2105).
Read-only sources do not acquire a writable new row; incomplete materialized
data stops record-cycle navigation with 3035. Unbound forms wrap controls
without creating SQL records.

A VBA control AfterUpdate handler's explicit focus change is respected. The
focus-navigation scope retains closing documents across nested native VBA/
message-box loops, then permits deferred deletion after unwinding.

## Evidence

Native Qt/VBA/SQL tests assert actual focus, record selector and persisted rows: 
Current Record wrap without SQL save; All Records forward/backward; control
validation cancellation/retry; record validation cancellation/retry; new-row
insertion without empty extras; disabled additions; handler focus override;
Null/invalid/unsupported mode preflight; unbound control wrap; hiding/disabling
a dirty focused control without a record save or navigation. A dynamically
opened form closes from its control AfterUpdate, opens a real MsgBox, remains
alive during the nested loop, then is deleted after returning. Its SQL value
is saved once without inserting another row.

The original export contains no explicit Cycle settings; no new original
business-handler execution count is claimed. Existing original layout and
calendar/DAO regressions remain mandatory. Current documentation is the contract
reference, not a Windows Access runtime oracle:
[Form.Cycle](https://learn.microsoft.com/en-us/office/vba/api/access.form.cycle).

## Remaining scope

Embedded/nested controls retain their existing Qt path; page/section-cycle
semantics, Ctrl+Tab subform escape, mouse/Enter focus transitions, complete
Exit/Enter/LostFocus/GotFocus event order and Windows focus/rendering parity
remain open. This does not establish general Access keyboard compatibility.
