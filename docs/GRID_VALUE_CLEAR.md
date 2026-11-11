# Native selected field-value clearing

Outside a cell editor, unmodified Delete clears a contiguous visible cell range
to pending SQL NULL without publishing or changing clipboard contents. The
context menu exposes the same explicit field-value action. Complete row
selections use the existing staged record deletion path, including mixed saved
and inserted records. Editor Delete and Shift+Delete retain their existing text
editing and Cut behavior.

Cut and value clearing share the existing validated range preparation, protected
field/source checks, pending-cache snapshot bound and typed NULL mutation path.
Hidden/projected field mappings are retained. Binary values can be cleared;
generated fields, readonly/deleted targets and discontiguous visible selections
are refused before mutation. Partial writes and implicit SQL saves are not added.
Normal Save handles constraints transactionally. Native range Undo/Redo restores
the pre-clear drafts and can replay the clear; other edits retain their existing
history invalidation rules.

This does not establish Windows event/selection/confirmation equivalence or
external clipboard interoperability. F148/F168/F340 remain open; the fixed456
catalog and386 confirmed partial scopes/70 open full requirements are unchanged.

Existing mainwindow_documents/currency_field_documents passed together2/2 in
16.24s (15.70/0.54s), app built. Actual Delete/context menu, NULL versus stored
values, clipboard preservation, Undo/Redo, binary/protected/readonly/discontiguous
targets and no SQL before Save were checked. PageDown crossed the initial
fetched-row boundary in the actual mainwindow while preserving another draft;
existing current-row Save/failure/PageUp cases also passed. No new CTest ID.
