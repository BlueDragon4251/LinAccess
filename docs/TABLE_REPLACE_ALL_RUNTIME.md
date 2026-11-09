# Native datasheet Replace All

The existing Ctrl+H replacement dialog now has a real Alle ersetzen action.
It replaces literal whole-field text matches in one visible non-key text column
of a writable native table and saves the batch through the existing transactional
model path. All rows of the current filtered/sorted view are fetched, up to the
native1,000,000-row limit. Search direction must be All; the button is disabled
for Up/Down. Qt case-sensitive/insensitive comparison matches the existing native
Find behavior, including Unicode, rather than silently using SQLite ASCII NOCASE.
SQL NULL values remain untouched and distinct from empty text.

Existing active-editor/current-record validation runs first. Unrelated pending
record drafts refuse the action without saving or discarding them. The model must
be clean when batch staging starts. Existing physical schema constraints remain
enforced. The actual SQL updates share one transaction; a later constraint error
rolls back earlier updates and discards only this operation's replacement drafts.
Qualified Qt row rollback avoids the saving guard blocking virtual revertRow.
As with ordinary Save, a committed batch followed by a view-refresh error is
reported as saved with a refresh error; it is not claimed to have rolled back.

Success reloads the view and restores the current record by its primary key,
including when sorting moves it. A record no longer visible under the filter is
reported explicitly. The dialog reports the actual matched-row count and saved
state. No placeholder updates, per-SQL consent prompts or raw SQL construction
from search/replacement strings were added.

Bounded scope: one non-key text column, literal complete field values, stable
source and a keyed writable table. Readonly/query/binary/generated/key-column,
partial/wildcard/all-field/locale-specific matching, nested workspace transactions,
concurrent-writer/trigger/event equivalence and complete Access Undo behavior are
not claimed. F373/F160/F168 remain open. The existing single-value Ersetzen action
keeps its current-record draft behavior; Alle ersetzen performs the batch save.

The existing mainwindow target drives the actual button, replaces Unicode case
variants, preserves NULL and the sorted current key, provokes a second-record
CHECK failure to verify full SQL rollback/clean cache, and preserves an unrelated
unsaved draft. Final1/1 passes (2.92 s). Updated Debian payload/build-ID/pinned
SQLite3.54.1 and four original fixture hashes pass. Last broader full suite is
the preceding label milestone48/48 (86.70 s); core runtime is unchanged from fresh
GUI-free payment-components1/1 (1.79 s). No new target/ID; fixed456/384/72.

Native literal substring/prefix replacement now extends the earlier whole-field
boundary; see LITERAL_REPLACE_RUNTIME.md. Full Access semantics remain open.
