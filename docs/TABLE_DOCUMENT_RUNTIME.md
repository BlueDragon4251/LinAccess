# Native project table documents

These operations use the actual Qt table view/model and native project SQL
storage, including working copies imported from Access. They do not add original
MDB/ACCDB write support. Source fixture bytes remain unchanged.

## OpenTable

DoCmd.OpenTable(TableName, View, DataMode) supports the normal datasheet view
(acViewNormal=0), ordinary editing (acEdit=1, the default) and a read-only view
(acReadOnly=2). Name lookup uses the current project's real table metadata;
new views join the existing Window menu and project document ring. Default and
named arguments run through the VBA object method binding. Positional holes and
named omitted positions are Missing; explicit Empty is coerced to numeric zero
for enum arguments rather than substituted with the default data mode.

Unsupported Design/Print Preview views and acAdd/append-only mode fail with
3251 before creating a page. Required/malformed/Null arguments retain checked
runtime handling. An unavailable production UI host returns 3251; an active
modal context returns 2046; a missing project table returns 7874. Complete Access
error precedence and library-database fallback remain open.

References: [Access 2003 OpenTable](https://learn.microsoft.com/en-us/previous-versions/office/developer/office-2003/aa220411%28v%3Doffice.11%29),
[AcOpenDataMode](https://learn.microsoft.com/en-us/office/vba/api/access.acopendatamode).

## View-level read-only policy

The native table model rejects data writes, insert/delete, block paste,
previous-record value assignment and transactional submission in read-only mode.
Its flags remove editing and user checkbox mutation. Add/Delete/Save/Discard
buttons are disabled and the page identifies the read-only state. Copy, selection,
checked readonly Zoom and query sorting retain their existing read paths.
This is a table-view policy; DAO/SQL access elsewhere keeps its own contracts.
It is not a database file lock or a claim of full Access permission semantics.

Changing the policy while cached drafts are dirty is refused without discarding
those drafts. Reopening an existing table with an explicitly conflicting data
mode returns 3251; existing-mode transition equivalence has not been established.
An ordinary explorer reopen preserves the existing view's mode and cache.

## Existing view selection and lifetime

A repeated table open reuses the same page/model. DoCmd.SelectObject acTable,
name,False can activate an actually open table with source-field validation and
modal guards. A closed/missing table in SelectObject remains an explicit 3251
boundary; exact closed-object errors are not inferred from OpenTable behavior.
The caller's Me stays bound to its own live form, and closing that caller does
not close the table page owned by MainWindow. Table activation clears the old
active form context. Other unimplemented table DoCmd commands, Screen.ActiveDatasheet,
query/editor-window execution and full Access window/format/error parity remain
open; commands must not silently operate on the old form caller.

## Verification

Existing main-window/Currency/layout targets pass 3/3 (16.70 s), exercising
interpreted named/default OpenTable calls, real views, read-only cell/checkbox/
insert/delete refusal, unchanged readonly SQL values, editable SQL persistence,
dirty policy-change refusal and unsupported positional-hole append mode.
GUI-free rebuild/core passes 1/1 (1.68 s). Explicit Empty DataMode is independently rejected as zero/unsupported append
rather than treated as the editable default (main-window 1/1, 1.91 s).
Full rebuilt desktop suite passes 48/48 (92.13 s). Debian payload/permissions/
Build-ID/pinned SQLite checks and all four fixture hashes pass. No extra targets.


## SaveRecord on the active table (2026-11-03)

DoCmd.RunCommand acCmdSaveRecord and Application.RunCommand now submit the
actual active native table's cache. The specialized interpreter Save statement
and ordinary object-method dispatch share this route. The existing table Save
button uses the same transactional helper, including field validation, rollback/
retry behavior and primary-key selection restoration. Read-only Save is an
unavailable command (2046); other unimplemented table dispatcher choices are
3251 instead of acting on the old form caller. Native errors/descriptions return
to VBA without adding an SQL consent/error dialog on the programmatic path.

A never-activated form retains its initialization command context unless it
explicitly opened/selected a table. That explicit switch is recorded separately
from real form activation events. A constructor Current save was checked with a
dirty background table: its cache and SQL stay unchanged. Me remains the caller
while a subsequent explicit table Save commits the actual table.

Existing main-window/Currency/layout targets pass 3/3 (16.13 s); final focused
initialization check passes 1/1 (1.93 s). Actual DoCmd and named Application calls
persist two different SQL values, readonly mutation is refused, and existing
UI/save/selection cases still pass. Last complete suite/package baseline remains
local e054910 / private API 0328119, 48/48 (92.13 s), GUI-free core 1/1 (1.68 s).
This Qt controller increment is focused-verified, with no new test targets.


## Close on native table documents (2026-11-03)

DoCmd.Close acTable,name,Save and default DoCmd.Close resolve real table pages
through MainWindow, preserving the calling form's Me context. Default Close
requires the active table context; forms/reports retain their existing route.
Save values 0..2 are validated. acSaveNo does not discard record data: the Save
argument concerns object design. Pending fields are validated and dirty table
records use the same transactional submission as SaveRecord before page removal.
Field and SQL failures return to VBA with the editor/cache and table retained;
programmatic SQL failure does not open an error dialog. Missing/closed named
tables remain 3251, pending unavailable/modal contexts 2046. Table design-save
prompts and exact Access error precedence remain unverified.

A parent-hidden active cell editor remains discoverable and checked before
closing its table. This prevents hidden invalid input from being silently lost.
A table's open native context menu exits when its owning view is destroyed,
including reentrant interpreted default Close during its nested event loop.
Existing main-window/Currency/layout targets pass 3/3 (15.85 s): rejected numeric
input, SQL CHECK rollback/draft retention, successful corrected SQL persistence
with acSaveNo, and default Close from a live popup are covered without new targets.

Full rebuilt desktop suite passes 48/48 (91.15 s), GUI-free core 1/1 (2.01 s),
Debian payload/permissions/Build-ID/pinned SQLite checks and all four original
fixture hashes pass. This also covers the preceding SaveRecord increment.


## Native table selection and Copy commands (2026-11-03)

DoCmd.RunCommand acCmdSelectRecord selects the current native datasheet row.
DoCmd/Application RunCommand acCmdCopy routes the actual table selection to the
existing checked rectangular TSV clipboard path. Both use field validation;
read-only views retain selection/copy access. Copy keeps dirty manual-submit
record drafts and does not write SQL. Caller Me remains bound to the form.
Clipboard text format remains the documented LinAccess TSV contract; no claim of
all Access clipboard formats, OLE or external application parity is made.

Existing main-window regression passes 1/1 (1.97 s), including invalid numeric
field refusal with unchanged clipboard. Last full desktop/package baseline is
local 2dd59b2 / private API 32cb556, 48/48 (91.15 s), GUI-free core 1/1 (2.01 s).
This controller increment is focused-verified; no new targets.


## ZoomBox on the active native table (2026-11-03)

DoCmd/Application RunCommand acCmdZoomBox uses the actual current table cell and
the shared native Qt Zoom dialog. Typed clipboard-text conversion feeds the
editor; accepted edits use checked model conversion and remain cached until a
normal Save. Read-only cells display in the same dialog with a Close button.
Cancel and readonly Close are normal completed commands with no mutation.
Invalid typed input still returns the model error; unsupported binary values
remain explicitly refused. Dialog-owner/model/index lifetime is checked before
applying any result. This does not add Screen.ActiveDatasheet or Access OLE Zoom.

SelectRecord preserves the previously active column while selecting the whole
row. Qt's ordinary selectRow moves the current cell to column zero; the native
command restores the persistent current cell without replacing row selection.
This was reproduced by a subsequent Zoom and fixed in the same path.
Existing main-window/Currency targets pass 2/2 (2.47 s), covering readonly Close,
accepted text edits, Cancel retaining the prior draft and unchanged SQL. Last
full suite/package remains 2dd59b2 / API 32cb556, 48/48 (91.15 s). No new targets.


## Bounded native table record navigation (2026-11-03)

DoCmd.GoToRecord can address an open named native table (acTable) or the active
native table by default. First, Last, Previous, Next and positive one-based GoTo
offsets use real view/model row ordering. Native RunCommand First/Previous/Next/
Last uses this same route. Named targets activate their actual page with existing
source-field/modal validation; Me remains the calling form.

Last/GoTo fetch actual SQL rows past Qt's initial batch, bounded to one million
rows. Checked invalid offsets/unavailable existing rows return 2105. Moving past
the last existing record does not synthesize a new-record placeholder. Explicit
acNewRec now uses the checked native record-creation path described below. These
bounds do not claim complete Access error precedence or addition-row behavior.

An existing current record update is transactionally saved before navigation
only when it is the sole dirty record in the cache. The same checked SaveRecord
helper restores its primary-key position before calculating the relative target.
SQL failure prevents navigation and preserves the draft for correction/retry.
Unrelated dirty rows, pending inserts/deletes and unsafe current-record identity
remain an explicit 3251 boundary; no unrelated draft is committed or discarded.
Explicit SaveRecord then clean navigation remains available for multirow drafts.

Full rebuilt desktop suite passes 48/48 (91.33 s), GUI-free core 1/1 (2.49 s),
Debian payload/permissions/Build-ID/pinned SQLite checks and all four original
fixture hashes pass. Existing focused trio passes 3/3 (16.79 s), exercising
300 actual SQL rows, named/default/relative targets, sole-record SQL persistence,
constraint failure/retry and untouched multirow drafts. Fixtures explicitly sort
ascending; production navigation follows the current view order. No new targets.


## Native table sort commands (2026-11-03)

DoCmd/Application RunCommand acCmdSortAscending/acCmdSortDescending sort the
actual native table by its active column through Qt/SQL and update the real
header indicator. A sole changed existing current record can be checked/saved
first; unrelated dirty rows and insert/delete batches remain explicit 3251.
The active field and primary-key record are restored after SQL sorting, fetching
rows as necessary within the bounded native search budget. Tables without a
primary key do not claim stable current-record identity across the sort.
Collation and NULL ordering use the existing native project SQL contract; this
is not full Jet/Windows locale or multi-column Access sort parity.

Existing main-window regression passes 1/1 (2.08 s): real ascending/descending
SQL row order, readonly sorting, primary-key/current-field retention past the
initial batch, checked sole-record persistence and refusal of unrelated drafts.
Last complete suite/package baseline remains e4e0900 / private API 714a96e,
48/48 (91.33 s), GUI-free core 1/1 (2.49 s). This controller change is
focused-verified, without adding test targets.


## Native table Find and checked view errors (2026-11-03)

Real table Find/Ctrl+F/FindNext, literal text modes, SQL fetch limits, cached
record saving, owning-page activation and dialog lifetime are specified in
[TABLE_FIND_RUNTIME.md](TABLE_FIND_RUNTIME.md). View refusals record the current
native error for VBA; unsupported binary Copy retains the actual 3251 boundary.

## Native SelectAllRecords (2026-11-03)

DoCmd/Application.RunCommand acCmdSelectAllRecords (109) uses the same checked
selection as Ctrl+A on the actual native table view. It validates an active cell
editor, fetches remaining SQL batches within the existing 1,000,000-cell selection
boundary and selects the rows without submitting SQL or changing the current
field. Readonly access remains available; unrelated cached drafts are preserved.
The constant is taken from Microsoft's AcCommand enumeration:
https://learn.microsoft.com/en-us/office/vba/api/access.accommand
General multi-record clipboard/deletion and form/report selection parity remain
open. This native table command does not imply those operations are available.

## Native table Paste command (2026-11-03)

DoCmd/Application.RunCommand acCmdPaste (191, Microsoft's AcCommand enumeration)
uses the real table Ctrl+V rectangular TSV path. The full block is parsed and
type-checked before cache mutation. An invalid later cell leaves earlier cells
unchanged; readonly targets reject 2448. Existing drafts or an editor that
produces a draft retain the 2501 boundary. Missing text MIME now sets a fresh
2046 native error instead of leaving an unrelated previous model error visible.
Successful paste changes only the manual-submit cache; normal Save reaches SQL.
Record append, general paste-over-draft merging, general multiline editor routing,
binary MIME, locale masks and Access-wide Paste equivalence remain open.

## Active table Requery (2026-11-03)

DoCmd.Requery with no control name and Shift+F9 on an actual table page rerun
the real SQL model query. Source changes and newly selected SQL rows are loaded
again; the native view starts at its first result row and retains the active
column. Readonly policy and existing sort remain attached to the same model.
The active field is validated first. A sole changed existing current record
with a primary key is saved through the transactional native Save path before
reloading; unrelated drafts, insertions and deletions refuse 3251 without being
discarded. Save failure retains its checked error and draft. Modal dialogs block
2046. Caller Me stays with the originating form; initialization without an
explicit table context keeps the existing form Requery path. Named-control
requery on tables, general multirow draft handling and exact Access record/
selection/error-priority equivalence remain open.

Microsoft documents the no-name Requery/Shift+F9 relationship and source reload:
https://learn.microsoft.com/en-us/office/vba/api/access.docmd.requery

## Native new/delete record commands (2026-11-03)

DoCmd/Application.RunCommand acCmdRecordsGoToNew (28) creates a real model
insert draft in the active native table. An active field is validated first;
a sole dirty existing current record is saved through the checked transaction
path before insertion. Existing insert/delete/unrelated drafts refuse 3251.
An empty table can receive its first draft without a current cell.

acCmdDeleteRecord (223) uses the native table's checked single-record deletion.
Selections spanning more than the current row now refuse 3251 before validation
or mutation, shared with the delete button and Ctrl+Minus. Readonly mutation
retains 2448. Both native operations use the manual-submit cache: subsequent
Save performs actual transactional INSERT/DELETE. This staged table behavior
is explicit; immediate Access commit, confirmation/SetWarnings, deletion events,
general multirecord deletion, addition-row UI and complete error/timing parity
remain open. Source enum values are documented in Microsoft's AcCommand table.

## Generated primary-key selection after Save (2026-11-03)

Transactional native INSERT captures the actual persisted primary key from the
inserted SQLite row on the same connection before commit, using an unshadowed
rowid alias and last_insert_rowid(). Captured keys are published only after a
successful commit, by original draft row. Save then restores the actual keyed
record and its current column after Qt model reset, including generated keys.
No MAX(), guessed row number or fabricated counter value is used. Failed writes
publish no keys; existing rollback/retry and postcommit-refresh protections stay.
WITHOUT ROWID tables or tables shadowing all rowid aliases have no captured
insert bookmark; successful SQL is retained without inventing a selection key.
This is native SQLite view-state handling, not general Jet bookmark parity.

A captured insert bookmark additionally requires SQLite changes()=1. A BEFORE
trigger that uses RAISE(IGNORE) cannot reuse a previous row's last-insert ID.
The existing native integration target exercises actual ignored INSERT and
queries the retained SQL rows independently.

## Explicit new-record navigation and complete source fetching (2026-11-03)

Named/default table DoCmd.GoToRecord acNewRec now shares acCmdRecordsGoToNew's
actual insert-draft path. A named target is activated and validated before
mutation; caller Me and unrelated documents remain unchanged. Readonly targets
reject 2448 and existing insert/unrelated drafts reject 3251. The actual first
row in an empty table can be created and saved through this route.

Native table insertion now fetches the remaining existing SQL rows before
appending its draft, shared by the button, Ctrl+Plus and VBA routes. A partially
loaded 300-row source gains its draft at row 301 while SQL still contains 300
rows. Total native rows including the draft stay bounded to 1,000,000; a stalled
fetch refuses 3251. The explicit creation route does not add implicit Next-past-
last behavior or general navigation away from unsaved insert/delete drafts.

## Focused native Paste (2026-11-03)

acCmdPaste now checks actual QApplication focus before grid TSV paste. An owned
QLineEdit in a table cell editor or its modeless Find dialog receives normal
Qt paste at the real text selection/caret. Cell text remains an editor buffer
until normal field validation; it does not become a grid range or write SQL.
A Find query remains editable for a readonly data source. Disabled/readonly
text targets refuse 2448, missing text MIME 2046, and the existing 4 MiB/NUL
boundary remains explicit. Parent ownership is checked across Qt dialog windows;
other forms, projects and external applications are not fallback targets.
General multiline/rich editor command routing and full Access equivalence remain
open; the grid's checked TSV and dirty-draft rules remain unchanged.

## Focused native Copy and Cut (2026-11-03)

acCmdCopy/190 and acCmdCut/189 share actual owned QLineEdit focus resolution
with Paste, including native cell editors and modeless Find. Copy transfers
only selected text, Cut removes that selection from the editor buffer; normal
field validation/Save still control model/SQL changes. A readonly data source
does not make the independent Find query text readonly. No selection, hidden/
disabled targets or masked echo modes refuse 2046; readonly text Cut refuses
2448. Copy can read a selected readonly text field. The 4 MiB/NUL boundary is
checked before replacing the clipboard or cutting text.

Without a focused owned QLineEdit, Copy retains the checked grid TSV route;
range Cut remains explicitly unsupported (3251) rather than deleting a row or
clearing an arbitrary field. General multiline/rich/form command dispatch and
Access-wide selection/clipboard/timing equivalence remain open.

## Focused native text Undo/Redo (2026-11-03)

acCmdUndo/292 and acCmdRedo/199 operate on the actual owned QLineEdit's native
undo/redo stack, the same buffer used by Qt shortcuts. Missing history refuses
2046, readonly text 2448; hidden/disabled targets are not receivers. Undo/Redo
changes buffered text without prematurely validating the field or writing SQL.
Grid-record/range undo and general rich/multiline/form dispatch remain explicit
3251/open boundaries; no blanket Access undo semantics are claimed. The enum
values are independently documented in Microsoft's AcCommand enumeration.

Focused Copy likewise reads selected invalid numeric editor text without trying
to commit it. SelectRecord/Save still reject that invalid field and retain it.
Copy with no text selection raises 2046 and preserves the clipboard. The legacy
validation regression now checks this distinction instead of requiring Copy
itself to validate and close a text editor.

Native table range clipboard and bounded pasted-block Undo/Redo now extend the
older text-only boundaries above; see TYPED_GRID_CLIPBOARD.md. General Access
undo/redo, range Cut and unrelated pending draft merging remain open.
