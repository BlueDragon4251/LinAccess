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
