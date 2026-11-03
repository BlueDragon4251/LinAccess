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
active form context. General table DoCmd commands, Screen.ActiveDatasheet,
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
