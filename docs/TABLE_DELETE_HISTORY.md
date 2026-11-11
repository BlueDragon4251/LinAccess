# Pending multi-record deletion history

The clean keyed table single-record and multi-selection paths stage selected stored rows as
one model history action. Ctrl+Z restores exactly those pending deletions in the
existing cache; Ctrl+Y stages the same rows again. Neither action executes SQL,
reselects the source, evaluates insert defaults or manufactures replacement rows.
The existing selection and non-NULL primary-key validation remain in force.
Duplicate/out-of-range row requests, existing drafts, read-only sources and
reentrant mutation are refused before staging. A failed stage reverts only the
rows staged by that action.

A transactional Save failure retains the deletion action, so Undo/Redo remains
available after a later foreign-key rejection. A successful Save retires history
through the existing refresh. Ordinary subsequent edits, row changes, explicit
record cancellation and source reload invalidate history as for range paste.
Deletion participates in the bounded multi-action native history (at most 64
actions and a conservative 256 MiB combined payload estimate); see
[TYPED_GRID_CLIPBOARD.md](TYPED_GRID_CLIPBOARD.md). A later native range action
can be undone first, followed by the deletion. Ordinary edits still invalidate
the stacks. This does not provide an unlimited Access undo stack.

This scope does not undo committed SQL deletion, add dirty-cache single-record
history, restore the original selection/focus/event sequence, or implement all
Access confirmation and deletion events. F168, F148 and F340 remain open; the
456 catalog IDs and 386/70 partial/full counts are unchanged.

The existing mainwindow_documents case exercises noncontiguous rows using actual
Ctrl+Minus, Ctrl+Z and Ctrl+Y; verifies saved rows remain untouched; verifies the
exact pending rows after Redo; then forces a late foreign-key Save failure and
checks Undo/Redo before the existing successful transactional deletion. Existing
currency_field_documents covers the shared paste history. No new CTest target.

Clean saved single-record deletion uses the same history action. Dirty-cache and
keyless single-row paths retain their existing behavior. Existing GUI coverage
checks singleton Undo/Redo, invalidation by subsequent cell edits, duplicate and
out-of-range requests and read-only refusal without modifying saved data.
