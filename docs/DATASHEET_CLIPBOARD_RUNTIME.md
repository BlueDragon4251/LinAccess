# Native datasheet selection and text clipboard

Shift+Space selects the current row, Ctrl+Space the current column and Ctrl+A
all cells. Column/all selection fetches the remaining Qt query batches first;
it refuses an unsupported size instead of silently selecting only the initial
fetch batch. The current limit is one million cells.

Ctrl+C and the main-window Copy menu command copy a rectangular native table
selection to the actual Qt text clipboard. Rows use CRLF, columns use tabs;
tabs/newlines/quotes inside a value use doubled-quote TSV quoting. No header row
is added. The native table model provides checked values, retaining Currency
precision, finite numbers with round-trip digits, ISO dates and Boolean text.
Null and empty text both appear empty in this text format.

An active editor must pass normal cell validation into the manual-submit cache
before copying; copying does not submit database changes. Binary/non-NULL BLOB,
NUL-containing text, hidden cells, disjoint selections and oversized text are
refused before replacing the clipboard. Text is capped at four million UTF-16
units. Cut/Delete for whole ranges remain disabled. Paste is enabled only for the
bounded clean-model case described below.

The existing main-window/Currency targets cover actual selection/copy keys and
menu dispatch, SQL-backed cells, Currency/date/Boolean values, invalid editor
preservation, quote handling, full selection beyond the first fetch batch and
binary/disjoint refusal. Typed clipboard MIME, record append, paste over preexisting drafts, headers,
regional format masks and exact Access/Windows clipboard representation remain open.


## Checked range paste into existing cells

Ctrl+V and the Paste menu action accept strict, headerless quoted TSV, anchored
at the current cell. Tabs, embedded newlines and doubled quotes round-trip
through the shared native delimiter scanner; legacy CSV import behavior stays
separate. The input and cell budgets match the bounded clipboard scope. Invalid
Unicode, malformed quotes, ragged rows, read-only targets and out-of-range
rectangles fail before range values are changed. Existing rows are fetched as
needed; this path does not create records or columns.

The table model must use manual submit and start with a clean cache. An active
edit is validated first; existing pending changes are preserved and require
save/discard before a range paste. Every target and converted field value is
prepared before applying the block. Application signals are deferred until the
block has been applied; an unexpected cache-write failure reverts only this
block because the cache was clean on entry. Existing SQL constraints are still
checked by the normal transactional Save path. Paste itself does not submit SQL.

The existing regressions exercise a bad second Currency row without partial
changes, valid typed rows and actual SQL saving, dirty-cache refusal, native
Paste menu dispatch and quoted text through clipboard/cache/database. This is
not complete Access/Windows clipboard or all editing-mode/undo parity.

## Transactional save and retry

Native table Save writes the complete pending batch inside one database
transaction using the model's protected SQL workers. Qt cache entries are not
marked submitted before commit. A constraint failure rolls back SQL while
retaining every pending update, insertion and deletion, including rows written
before the failing statement. The user can correct a value and save the whole
batch again. Generated AutoNumber fields and database defaults retain the
existing insert behavior. Persistent row anchors retain draft bookkeeping when
an unsaved inserted row is reverted or removed.

This avoids Qt submitAll marking earlier successful rows submitted even when a
later statement fails and the surrounding database transaction rolls back.
The application table Save path uses submitTransactionally; direct external
calls to the inherited Qt submitAll do not gain this guarantee. Successful
commit refreshes the model. Database constraints remain authoritative; this
is not a claim of Access locking, multiuser conflict or undo parity.

Existing Currency integration cases reproduce a second-row unique constraint
failure, retain both updates and retry successfully. A mixed update/insert/
delete batch verifies failed-save preservation, AutoNumber/default values and
row shifts after reverting an inserted row. These and main-window Save checks
pass 2/2 (2.10 s). The shared TSV scanner also passed the GUI-free core check
(1/1, 2.63 s). The rebuilt full desktop suite passes 48/48 (86.21 s). Debian payload,
permissions, matching Build-ID and pinned SQLite checks pass; all four original
fixture hashes match.


## Commit succeeded but refresh failed (2026-11-03)

After a successful database commit, native table Save retires the submitted
drafts before reloading metadata/query rows. A refresh failure therefore cannot
re-execute an already committed insertion on the next Save. The first failure
message explicitly states that the changes were saved but the view could not
be refreshed. Before-commit failures still roll back and retain the entire
draft batch for correction/retry. A failed refresh can leave the old query view
unavailable or stale; correcting the metadata and refreshing restores the data.

A synthetic metadata mismatch introduced during beforeInsert reproduces the
previous duplicate-insertion bug. The regression commits one row, refuses the
refresh, retries Save without a second insert, repairs metadata and reloads the
one persisted row. No original data or fixtures are changed.

## Checked floating-point editors and nullable fields (2026-11-03)

REAL/FLOAT/DOUBLE-backed cells now expose checked round-trip decimal text to
the native editor, using Double max_digits10. Qt's default two-decimal spin
editor previously changed an SQL value just by opening, closing and saving
the cell; this was reproduced before the fix. Untouched and explicitly edited
finite values now retain the actual supported Double precision. Unsupported
numeric storage refuses editing instead of fabricating a checked value.

Non-Boolean/non-BLOB NULL fields expose a typed-null text EditRole. Native
editors can therefore open and close without materializing a spin/date widget's
default value. REAL, Currency and Date NULL cells remain NULL after Save. This
uses the existing shared editor-change check; explicit user edits remain typed
and validated. Exact Decimal, Windows regional masks and complete Access field
formatting remain open. Main-window/Currency checks pass 2/2 (2.60 s); the rebuilt
complete suite passes 48/48 (88.41 s). Debian payload/permissions/Build-ID/pinned
SQLite checks and all four original fixture hashes pass. GUI-free core remains
checked 1/1 (1.53 s), with subsequent Qt-only changes.
