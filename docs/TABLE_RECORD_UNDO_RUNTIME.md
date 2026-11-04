# Discard pending edits in the current native datasheet record

Esc outside a live cell editor discards pending changes in the current native
manual-submit model row. A context action, Datensatzänderungen verwerfen,
exposes the same operation. Inside a live editor, the existing first Esc only
cancels the uncommitted editor input; the next Esc in the view can discard the
record's cached changes. This does not navigate or save unrelated pending rows.

The actual AccessTableModel checks pending state per record and uses its
existing revertRow path, preserving other dirty/deleted row anchors. Undo during
model mutation/save is rejected. A cancelled insertion removes that pending row
and restores a remaining current position. Model/view ownership checks prevent
restoring an index into a replaced model. Saved SQL is never reversed by this
operation. Readonly query views gain no write/history behavior.

The existing main-window target verifies two-stage editor/record Esc, preservation
of a separate pending row and unchanged SQL, cancelling an insertion while
retaining another draft, discarding the last pending row and leaving a committed
Save intact. No extra CTest target.

This is current-record draft discard. General action history, committed SQL
undo/redo, acCmdUndo outside text editors, full delete-selection reversal,
Access form Undo events and Access 2003 error/focus/timing equivalence remain
open. It does not claim complete Access Undo behavior.


## Current-record pending-edit discard with Esc (2026-11-04)

Native datasheet Esc/context action uses actual per-row manual-submit revert
without saving or discarding other drafts. First Esc inside the cell editor
retains prior cached drafts; subsequent record Esc discards that row, including
a pending insertion. Committed SQL remains unchanged. TABLE_RECORD_UNDO_RUNTIME.md
defines remaining history/deletion/form/oracle boundaries. Existing main-window
regression only. Full desktop 48/48 (79.77 s), GUI-free build confirmed and core
1/1 (1.74 s), targeted GUI/SQL 1/1 (2.10 s), Debian payload/permissions/Build-ID/
pinned SQLite and all four original fixture hashes pass. Scoped checklist
374 checked / 72 open; not complete Access 2003 parity.
