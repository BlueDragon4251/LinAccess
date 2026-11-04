# Scoped native SendKeys

Global SendKeys and DoCmd.SendKeys(Keys, Wait) share the real Qt dispatcher.
Named and omitted method arguments are supported; omitted Wait is False.
Malformed sequences are rejected before emitting keys; Null raises 94.

Each key resolves actual focus again. Active table cells, their modeless Find
and modal Zoom dialogs participate, using explicit widget parent traversal
across Qt window boundaries. The originating database connection and database
identity stay fixed. Switching to another project cannot redirect old keys.
Database explorer SendKeys remains unsupported (3251); a missing eligible
receiver fails with 2110. Asynchronous failures appear in the requesting
form status. Closing the requester or stopping/Quit cancels remaining keys.

Batches retain FIFO order. Continuations are scheduled before native delivery,
allowing subsequent keys to run inside a key handler's modal event loop.
Wait=True waits for handler completion; Wait=False queues delivery. DoEvents
drains queued keys through the interactive Qt event loop. Transient activation
may settle for up to 500 ms. Limits: 4,096 keys per sequence, 1,000 repeats,
16 parser nesting levels, 65,536 queued keys and 32 nested native key handlers.
These are explicit resource boundaries, not measured Access limits.

Real form TextBox typing changes Text first. Value remains the original bound
value until validation/commit; Save then reaches actual SQL. Table typing uses
its checked draft cache and normal Save pipeline. No OS-wide key injection,
external application targeting, parentless table context menus or root file
picker routing is claimed. Full Access keyboard grammar/timing equivalence
remains open.

Existing main-window, layout and original Kassa event targets exercise ordered
queued batches, DoEvents, focus changes, dialog continuation, named method calls,
checked typed edits, requester isolation and SQL after Save. The original
Wait=False probe yields until the native queue is drained rather than assuming
one Qt processEvents call delivers a complete sequence; its content, receiver
and event assertions remain intact. No additional CTest targets were added.


## Validated vertical record keys (2026-11-04)

When the native table view receives unmodified Up/Down (including keypad keys),
a destination in the loaded/fetchable record set uses the existing native
previous/next record dispatcher. Live cell input is completed and the owning
project/page is activated before the existing checked keyed Save/navigation
path. SQL failure retains the current row and pending model value instead of
allowing Qt's default cursor movement. Successful movement retains the current
column. Readonly native tables use the same dispatcher without introducing writes.

The view keeps a weak-owner navigation callback; closed pages and cancelled
activation refuse. Existing row-fetch limits, unrelated-draft, pending-insert/
delete, current-key and context guards remain active. The last-row addition
placeholder, general mouse/Tab/Page/modified-key movement, broad filtered-query
refresh and Access 2003 editor/focus equivalence are not added by this increment.
Editors may consume their own keys before the view receives them. This is not
complete Access datasheet keyboard parity.

The existing main-window GUI target executes actual Down/Up keys, verifies SQL
Save before a successful row change, verifies a real UNIQUE failure preserving
the current row/cache and unchanged SQL, then uses Esc and retries navigation.
No extra CTest target.


## Validated Up/Down datasheet record movement (2026-11-04)

Bare vertical record keys reuse actual native current-record Save/navigation,
retaining the current column and stopping on SQL failure with drafts/selection
intact. TABLE_KEYBOARD_RUNTIME.md records destination, edit, unrelated-draft,
addition-row and general movement/oracle boundaries. Existing main-window target
checks actual keys, saved SQL, UNIQUE failure, Esc and retry. Full desktop 48/48
(77.41 s), GUI-free build confirmed and core 1/1 (1.79 s), targeted GUI/SQL
1/1 (2.30 s), Debian payload/permissions/Build-ID/pinned SQLite and all four
original hashes pass. Scoped checklist 380 checked / 72 open; no extra CTest
target or complete Access 2003 parity claim.


## Validated adjacent-row Tab/Backtab (2026-11-04)

MoveNext/MovePrevious cursor requests that cross to an adjacent existing row
now reuse the same checked record Save/navigation callback. This covers Tab and
Backtab in the view and the native delegate's editor-completion path. The source
record is saved before movement; the actual post-save row and Qt-proposed
logical column form the destination, preserving visible-field wrapping. Failed
Save returns the current index and retains the pending row. Reentrant cursor
requests do not recursively run record navigation. Weak model/view ownership
checks reject stale destinations after callbacks.

Same-row field movement and existing endpoint behavior remain Qt's existing
path. General wraparound/addition rows, arbitrary hidden-row skips, modified
commands, mouse selection and full Access editor/focus/error equivalence remain
open. This does not claim complete Tab/datasheet parity.

Existing main-window cases verify real view Tab and Shift+Tab positions, actual
SQL Save, editor text committed and saved before Tab movement, and a real UNIQUE
failure that retains the original record/cache/SQL, followed by targeted Esc.
Earlier vertical-key and SendKeys cases remain in the same target. No new CTest
target.


## Validated adjacent-row Tab/Backtab movement (2026-11-04)

View/delegate cursor requests to adjacent existing rows reuse checked native
Save/navigation, restore the proposed logical column after Save and preserve
current drafts on failure. Existing GUI cases execute view Tab, Backtab, editor
Tab, actual SQL persistence and UNIQUE failure/Esc recovery. Endpoint/addition/
hidden-row/general movement and full Access oracle limits remain in
TABLE_KEYBOARD_RUNTIME.md. Full desktop 48/48 (78.54 s), GUI-free build confirmed
and core 1/1 (2.07 s), targeted GUI/SQL 1/1 (2.48 s), Debian payload/permissions/
Build-ID/pinned SQLite and all four original hashes pass. Scoped checklist
381 checked / 72 open, no extra CTest target or complete Access parity claim.
