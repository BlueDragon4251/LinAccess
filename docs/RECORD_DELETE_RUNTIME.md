# Native single-record selection and deletion

Bounded native `RunCommand acCmdSelectRecord` (50),
`RunCommand acCmdDeleteRecord` (223) and bounded Access 95 menu compatibility
`DoMenuItem acFormBar, acEditMenu, 8/6, , acMenuVer70`. The latter preserves
omitted positional arguments; other menu contexts/versions/actions fail explicitly.
The default historical menu version is not silently interpreted as version 70.

Selection saves/validates the current record, focuses and highlights its native
record selector, and exposes bounded read-only `Form.SelHeight` (0 or 1). Editing
a control or loading another record clears selection. Delete on a selected native
record routes through the same command path as VBA. Text-editor Delete retains
its normal behavior when no record is selected.

Deletion currently supports a single saved row of a direct table form. Queries,
reports, new/empty records and reentrant record commands are rejected. A pending
edit must pass normal control/record validation before deletion. Imported and
in-memory VBA `AllowDeletions=False` prevents deletion before SQL or Delete events.
Null policy assignments are rejected. The bound
`Form_Delete(Cancel)` can prevent the deletion before SQL. SQLite savepoints then
hold the real DELETE until `BeforeDelConfirm(Cancel, Response)` and the native
confirmation dialog complete. The pending next record and Current event are
visible before confirmation. Cancel restores the deleted row and original
position; `AfterDelConfirm(Status)` receives success/programmatic/user-cancel
status. `Response=acDataErrContinue` suppresses the confirmation dialog. Foreign
key rejection reports 3200. Cancellation reports 2501 to DoCmd callers.

Existing saved form snapshot clones retain surviving bookmarks after commit;
deleted-row bookmarks expire. Cancellation restores the prior snapshot family.
Navigation, saves and project close are prevented while a deletion is pending.
Native command lifetime guards cover the complete callback/modal operation.

Targeted regressions cover programmatic Delete cancellation, BeforeDelConfirm
cancellation, actual Yes/No dialog input, response suppression/invalid/Null values,
pending-edit validation, forbidden callback edits, foreign-key protection,
surviving/deleted bookmarks, empty/new records and Delete-key delivery. Modal
tests inspect the pending SQL deletion and verify nested deletion, navigation
and close are rejected. A second document proves Application.DoCmd operates on
the active target while subsequent Me access remains on the calling form.

The unchanged module SHA-256
`dbd3623b9404e490a814b70de6318d2c0877969b7feef081696d3e235ae3b427`
and its unchanged original layout execute one original deletion procedure twice:
one explicit call and one actual bound button click. Two prepared synthetic rows
are independently counted before, during and after native confirmation. The UI
button is made visible/enabled through runtime properties for this probe. No
original rows/source/credentials are copied into the regression output. The
original Copy/PasteAppend procedure remains unsupported at runtime.

Boundaries: multi-record selection/deletion, Copy/PasteAppend, datasheet editing,
full Access focus/event scheduling, all writable query sources, global warning
options, callback edits during the guarded operation, concurrent writers/locks,
tentative snapshot visibility and a Windows
Access 2003 comparison remain open. This is not full DoMenuItem/RunCommand support.

Contract references (documentation, not an Access 2003 runtime oracle):

- https://learn.microsoft.com/en-us/office/vba/api/access.docmd.domenuitem
- https://learn.microsoft.com/en-us/office/vba/api/access.accommand
- https://learn.microsoft.com/en-us/office/vba/api/access.form.delete
- https://learn.microsoft.com/en-us/office/vba/api/access.form.beforedelconfirm(even)
- https://learn.microsoft.com/en-us/office/vba/api/access.form.afterdelconfirm(even)
