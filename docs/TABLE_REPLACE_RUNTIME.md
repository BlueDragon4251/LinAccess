# Native whole-text-field replacement

Editable native table datasheets offer a modeless Suchen und Ersetzen dialog
through Ctrl+H (including child cell editors), Bearbeiten > Ersetzen and native
acCmdReplace/29 routing. The shared Find dialog retains normal search behavior;
the replacement variant selects one visible text field and whole-field matching
with optional case sensitivity. Weitersuchen uses the existing real bounded
search/navigation path.

Ersetzen checks the current field value and writes the new String through the
actual manual-submit AccessTableModel. It leaves a normal pending edit; Save or
validated navigation uses the existing transactional keyed SQL path. Existing
validation can save prior drafts before a subsequent search/replacement action.
Unique/SQL failure retains the draft for correction and leaves SQL unchanged.
Readonly, hidden/nontext, unsupported matching, missing/null primary keys and
new unsaved records refuse explicitly. Search/replacement text is bounded to
one million characters. This operation never changes original MDB files.

The action captures the current stored primary key and source field before
validation and verifies them afterward. A changed current record/source cancels
2501 instead of editing another row. Nested search/replacement is guarded.
Dialogs follow their actual native view lifetime. Supported editable query
datasheets share this view/model path, but the new regression specifically
executes a table-backed flow; general query updateability is not expanded.

The existing main-window target exercises actual Ctrl+H, dialog search/replace,
SQL before/after Save, a real UNIQUE failure retaining pending values, readonly
refusal, menu activation, dialog destruction and a reentrant record-switch
refusal. No new CTest target or original source/data is used.

Partial/first-occurrence and wildcard replacement, Replace All, all-field and
typed nontext replacement, original form event timing, complete undo/confirmation
behavior and an Access 2003 search/replace oracle remain open. This is bounded
whole-field replacement, not complete Access Find/Replace parity.

Current AcCommand numeric reference, not an Access 2003 execution oracle:
https://learn.microsoft.com/en-us/office/vba/api/access.accommand


## Native whole-text-field datasheet replacement (2026-11-04)

Ctrl+H, native edit menu and acCmdReplace/29 route to an actual modeless
whole-field text replacement path. It edits the real manual-submit model,
uses keyed transactional Save and retains drafts on SQL failure. Captured
primary-key/source/field checks cancel reentrant record changes before mutation.
TABLE_REPLACE_RUNTIME.md defines search/text/readonly/new-record limits and
remaining partial/bulk/form/undo/oracle work. Existing main-window target only.
Full desktop 48/48 (77.04 s), rebuilt GUI-free 1/1 (1.36 s), targeted GUI/SQL
1/1 (2.15 s), Debian payload/permissions/Build-ID/pinned SQLite and all four
original fixture hashes pass. Scoped checklist 373 checked / 72 open, not
complete Access 2003 parity.
