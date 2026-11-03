# Native direct-table DataEntry

SaveAsText DataEntry metadata and native VBA Me.DataEntry/unqualified DataEntry
now select a real direct-table entry session. Existing rows are excluded by a
bound SQL rowid predicate, not a display-only flag. Rows inserted by this form
remain in its session; successful key changes update their identity and committed
native deletion removes the session identity. The session is bounded to 10000
saved rows, matching the native materialized-form limit.

DataEntry has an effect only with AllowAdditions=True. Changing AllowAdditions
rebuilds the effective view after checked saving. Switching DataEntry=False saves
the current record, removes active filter/sort state and restores all source
rows; a new True transition starts a fresh entry session. Complex OR filters are
grouped before combining the session predicate. VBA property routes preserve
checked Boolean/Null handling; cancelled validation retains the previous mode.
acFormEdit overrides imported DataEntry to False. acFormAdd remains unsupported.

Existing native GUI regression starts with pre-existing rows, performs real
keyboard inserts and a primary-key update, checks the session recordset against
physical SQL counts, switches through native VBA, removes filter/sort and checks
AllowAdditions dependency. It also covers OR-filter isolation and unqualified
VBA DataEntry reads. Desktop build and form_current_record pass (0.60 s). No new
CTest target, full-suite or package rerun for this increment.

Writable query/join entry sources, arbitrary external SQL/rollback/rowid reuse,
all requery/filter/event/window combinations and exact Access 2003 session
reference behavior remain open. Unsupported active non-table sources fail
explicitly before native opening. This is a UI/data-session policy, not database
authorization. The last grouped 48/48/package evidence covers 992d082.

Reference: https://learn.microsoft.com/en-us/office/vba/api/access.form.dataentry
