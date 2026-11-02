# Typed native table and Access SQL editors

The application table view now uses AccessTableModel, a real QSqlTableModel
subclass with existing Access source-type metadata and checked assignment
conversion. Manual edits remain pending until Save; the existing application
transaction wraps submission. Currency edits use exact text EditRole/line editors,
then checked four-decimal/bounded storage conversion. Source Byte/Integer/Long
ranges, Boolean and canonical date conversion use the shared runtime helpers.
An invalid edit returns false, preserves the old cell and reports a typed generic
message through the application status bar. Invalid stored typed values are
shown as errors and blocked from accidental editing; reads never repair them.

Verified AutoNumber columns are read-only and pending rows display (AutoWert).
Inserts omit generated columns so SQLite native IDs and managed counter defaults
allocate them. Binary columns remain read-only; no OLE extraction/round-trip
claim. AccessTableDelegate removes Qt's 32767-character line-editor limit for
strings, preventing long Memo truncation on editing. This does not add a complete
Access datasheet/format mask/lookup/constraint/collation/undo implementation.

The custom SQL editor now uses AccessQueryModel and the existing Access compiler
rather than raw SQLite query execution. Declared parameter prompts feed checked
binding. SELECT displays native results; action queries execute through DAO
QueryDef.Execute and show RecordsAffected. Previous SELECT results are hidden for
action execution. SELECT and action definitions can be saved. Unsupported SQL
retains the existing diagnosed runtime boundaries; no universal Jet claim.

Focused coverage extends existing tests, without new targets. The native field
integration test drives actual Qt Currency/Memo editors, maximum Currency display,
rounding/rejected-value retention, source Integer overflow, date edits, 50000+1
Memo characters and SQLite/managed AutoNumber insertion. The main-window test
opens the actual query dialog, runs CCur/DateDiff, inserts via DAO, completes a
Currency parameter prompt and saves an action definition. Tests use temporary
synthetic projects only. Full Access 2003 visual/editor parity remains open.

Header sorting is blocked while manual changes or a live cell editor are pending;
Qt's default sort/select would otherwise clear them. The previous header indicator
is restored and the user is asked to finish/save/discard the edit. The existing
field integration test reproduces the cache-loss bug before the fix, then checks
both cached edits and still-open editor text without resets or writes.

## Checkbox and keyboard cell completion

Verified Boolean columns now expose native Qt checkbox state, with Space toggling
and checked persistence as -1/0. Existing stored Null is displayed partially
checked; this does not claim complete Jet Yes/No defaults/Null constraints.
Checkbox edits remain pending until Save.

Ctrl+S in the app's table view invokes the existing manual-save transaction.
The view first reads and validates any live cell editor through its actual index
and editor value property. Invalid input stays focused and is not discarded or
written; valid input enters the checked cache before editor close/submission.
The existing native field/main-window tests check real Space and Ctrl+S events,
old database values before save, final values after save, invalid Currency editor
retention and action-created table opening through the actual explorer.
Manual-save batching remains the existing app convention, not complete Access
per-record event/undo/keyboard parity. No new CTest target is added.

## Pending datasheets participate in close preflight

Native table tabs refuse closing while checked cache changes or a changed live
editor remain. The app's project/application close participant checks all table
pages before native forms unload. This covers tab/window close, project switch
and DoCmd.Quit; rejected Quit reports 2501 and leaves the calling form open.
Unchanged live editor values do not block closing. Finishing/saving or reverting
the pending changes allows normal close. The preflight does not write or discard
the pending data. The small interface keeps the shared form runtime independent
of the main-window implementation.

The existing main-window test checks a modified live editor, cached changes,
refused tab/window/project close, native VBA Quit with retained caller, unchanged
editor allowance and successful tab close after revert. Full Access automatic
record-save/events/undo/crash recovery and complete window semantics remain open.
