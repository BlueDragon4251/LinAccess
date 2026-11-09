# Native Form.RecordsetClone

RecordsetClone returns an independent existing native DAO snapshot cursor over
the form's actually loaded filtered/sorted record source. It works for direct
tables and read-only SELECT forms. Moving/searching the clone does not navigate
the form; assigning its Bookmark to the form uses the normal save/cancel and
Current-event navigation path. Separate property reads return separate cursors.

The property takes no arguments and is read-only (450 / 383). Reports, unavailable
sources, truncated results and closed forms retain the existing explicit errors.
Clones are read-only: this does not implement editable Form.Recordset or writable
query support. Supported existing Clone/Find/Fields/Move/Bookmark behavior and DAO
transaction restrictions apply. A Requery starts a new bookmark family; already
held clones keep their prior snapshot. Independently held clones can outlive the
form, while new property access through the closed form fails 2450.

The existing navigation fixture now exercises actual VBA Set rows=Me.RecordsetClone,
independent table cursors, cancellation/save navigation, read-only sorted SELECT
clones, empty-query BOF/EOF, stale-bookmark rejection and closure. One bundled
existing layout/GUI check passed 2/2 (12.40 / 14.29 seconds). No new test target,
full-suite/Windows oracle or refreshed installer proof is claimed. This extends
F054/F160; their complete requirements and catalog 456/386/70 stay unchanged.
