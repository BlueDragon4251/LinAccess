# Native deletion of selected table records

Ctrl+Minus and the existing native DeleteRecord command now accept multiple
fully selected stored rows, including noncontiguous rows. Whole-row selection
is required; additional partial-cell areas are refused rather than interpreted
as requests to delete other records. Existing single-record behavior remains.

The bounded multi-record path requires a bounded manual-submit model, a stable
keyed table source and non-NULL primary keys on saved rows. Existing updates, insert
default state and unrelated deletions are captured before staging and preserved.
Already-deleted records are refused. Selected pending inserts are removed from
the cache; mixed selections are staged in descending physical order. Their
supplied values and omitted defaults are retained in the pre-action snapshot. Persistent row anchors preserve the selected
records through validation. Reentrant deletion and disappeared/replaced views
are rejected. Rows are staged in descending order. A staging failure restores
the captured pre-action cache.

Deletion follows the existing table Save workflow: it is first marked pending;
Speichern submits the pending set in the model's existing transaction. A later
foreign-key failure rolls back earlier SQL deletes, leaving the pending choices
for the user to correct or discard. This is distinct from Replace All's direct
batch save. The table guidance now describes ordinary cell-save behavior without
claiming that every command waits for the Save button.

No complete Access confirmation/SetWarnings, BeforeDelConfirm/AfterDelConfirm,
external clipboard formats, concurrency or Windows
event/selection equivalence is claimed. F168/F148 remain open.

The existing mainwindow target drives Ctrl+Minus for noncontiguous rows, protects
partial-cell selection, checks no early SQL delete, forces a later foreign-key
Save failure with all rows retained, then saves the selected deletions successfully
with only the unselected record remaining. Existing native VBA SelectAllRecords/
DeleteRecord now stages the real selection and can be discarded. Target1/1
(2.80 s), full current desktop48/48 (84.19 s), updated Debian/build-ID/pinned
SQLite3.54.1 and four original hashes pass. Core runtime unchanged from fresh
GUI-free payment-components1/1 (1.79 s). No new target/ID; fixed456/384/72.

Pending multi-record Undo/Redo is now implemented within the bounded keyed
path, including preserved drafts; TABLE_DELETE_HISTORY.md supersedes that earlier limitation. Full Access
confirmation/event/selection and committed-deletion undo remain unproved.
