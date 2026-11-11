# Native PageDown across fetched table ranges

Before calculating a downward record page, the native view loads enough real
source rows for the current viewport height. Actual row heights and hidden rows
are counted; fetching remains bounded to one million rows. Each fetch validates
the same model and persistent current record before navigation proceeds. Failed
or nonprogressing fetches refuse explicitly.

The resulting Qt page destination uses the existing record navigation handler.
Current-record validation/selective Save, unrelated drafts, key/error guards and
PageUp behavior retain their existing policy. Fetching does not itself commit
SQL or create an Access addition row. Full focus/selection/Access event and
concurrent source equivalence remain open under F368/F148/F178.

Existing mainwindow_documents/currency_field_documents passed together2/2 in
16.24s (15.70/0.54s), app built. Actual Delete/context menu, NULL versus stored
values, clipboard preservation, Undo/Redo, binary/protected/readonly/discontiguous
targets and no SQL before Save were checked. PageDown crossed the initial
fetched-row boundary in the actual mainwindow while preserving another draft;
existing current-row Save/failure/PageUp cases also passed. No new CTest ID.
