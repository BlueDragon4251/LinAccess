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
the last existing record does not synthesize a new-record placeholder; new-record
navigation remains unsupported (3251). These bounds do not claim complete Access
navigation error precedence or addition-row behavior.

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
