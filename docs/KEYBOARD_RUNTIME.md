# Native Access keyboard and event subset

This is executable C++20/Qt behavior, not complete Access 2003 keyboard or UI
parity. The pre-change regression reproduced Escape closing the native form via
QDialog instead of cancelling edits. Tests use actual Qt key events and independently
inspect SQLite state. No Windows Access 2003 event or rendering oracle is available.

## Implemented native document keys

| Input | Native effect |
|---|---|
| Escape in a form | Restore the pending focused field to its typed focus-entry value; a further Escape discards remaining record edits |
| Escape in the report layout preview | Close that preview through its normal lifecycle |
| Shift+Enter, including keypad Enter | Save the current document record, honoring validation/cancellation |
| F9 | Recalculate displayed calculated controls without reloading records |
| Shift+F9 | Requery the current native form source with normal save/validation handling |
| F2 in a line editor | Toggle whole-field selection and editing at the end of the text |
| Enter in a line editor | Commit pending controls and advance focus; do not activate an unrelated default dialog button |
| Ctrl+W / Ctrl+F4 | Close the containing document window, honoring Unload cancellation |

Undo retains NULL versus empty text, restores duplicate bindings to the same field,
and can discard invalid numeric text before conversion. Uncommitted new records
are not inserted by Escape. Form.Undo discards the pending record; Form.Recalc
recomputes displayed expressions without a source reload or save. These are real
native form methods. Report invocation is explicitly unsupported.

This does not implement every Access Undo event, record conflict, edit/navigation
mode, TabIndex, option, keyboard layout, menu accelerator or designer shortcut.
Other controls retain their existing native Qt behavior; complete Access focus,
selection and UI parity still requires reference tests.

## VBA key events

SaveAsText OnKeyDown, OnKeyPress, OnKeyUp and form KeyPreview are retained.
Native keyboard delivery invokes only bound event procedures with actual ByRef
KeyCode/KeyAscii arguments and Shift bit masks (Shift=1, Control=2, Alt=4).
KeyPreview chooses form-before-control versus control-before-form delivery for each
stage. Embedded documents own their events; the parent's KeyPreview does not
intercept the child's input. Cancelling KeyCode suppresses subsequent/default key
processing. KeyAscii can cancel or replace a supported character before Qt inserts
it. Handler failures stop default processing and appear in native diagnostics.

The bounded key map covers the documented common navigation keys, ASCII letters/
digits, function keys and common punctuation; it is not full Windows virtual-key,
keyboard-layout, IME or locale emulation. KeyPress conversion/replacement currently
supports ASCII; native Unicode text without a KeyPress handler remains supported.
Non-ASCII bound KeyPress conversion, virtual-key replacement, macro/function event
bindings and complete Ctrl/Alt character-event behavior remain explicitly limited.
VBA vbKey constants and acShiftMask/acCtrlMask/acAltMask are available for the
implemented native event paths. No desktop-global input injection is introduced.

`tests/keyboard_commands_tests.cpp` checks saved SQL values, cancelled saves/closes,
NULL restoration, aliases, invalid numeric input, uncommitted new records, F2/focus,
record-count distinction between Recalc/Requery, real method execution and key-event
SQL traces. KeyPreview order changes the actually inserted character. KeyCode and
KeyAscii cancellation, typed modifier masks and a handler raising error 11 are
checked independently of default Qt behavior.

## Original source boundaries, not invented bindings

All original artifacts remain hash-verified and unchanged. The three existing
KeyPress procedures in modules
`6bde4373d4584a1626a6a7f5edbb14d072f3f01083bd797faf9fee914f04f7f8` and
`6e8e0a0e41e1a1768cff4a0b79a07205c7b0f799eb8b465af5397843ce048ed8`
are separately tested as explicit original calls. Only the latter procedure has a
matching bound control in the exported layout. The former layout binds two renamed
controls whose handler procedures are absent. Native dispatch records those as
ProcedureMissing and does not substitute the older similarly named procedures.
These checks do not establish a fully working original sale workflow.

The complete original login-layout probe also reproduces a source-column boundary:
its combo declares ColumnCount=5, while its unchanged handler requests Column(5).
The native handler fails with error 5 at source line 34, before authentication
comparison. The regression uses deliberately synthetic non-credential input and
never reads or logs original passwords. It does not claim successful login.
Microsoft's current [Column documentation](https://learn.microsoft.com/en-us/office/vba/api/access.combobox.column)
limits the zero-based index to ColumnCount-1; an actual Access 2003 comparison is
still needed to classify the supplied original application's behavior. Neither
layout/source bytes nor column limits are silently altered to make this pass.

Validation: Release build, 18/18 CTests (49.12 s), GUI-disabled Release core 1/1
(0.37 s), fresh .deb executable Build ID/payload/permissions, and four unchanged
original fixture SHA-256 checks passed. Remote distribution CI and Access 2003
reference behavior remain unverified.


## Native form date/time insertion keys (2026-11-03)

Ctrl+semicolon inserts the current local ISO date; Ctrl+Shift+semicolon/colon
inserts the current local HH:mm:ss into the focused native form TextBox,
replacing its selection or inserting at the caret. Existing control buffers,
Change events, BeforeUpdate cancellation and record saving are retained.
Read-only/disabled/calculated/unsupported bound contexts and capacity overflow
refuse mutation. The existing layout target passes 1/1 (13.58 s), exercising
actual keys, Change, cancelled save, SQL Date/time persistence and read-only
refusal. App builds. Regional Access date masks and datasheet keys remain open
in this increment; full-suite/package baseline remains 8dc6afa.


## Datasheet clock keys and keyed cell restoration (2026-11-03)

Forms and native editable datasheet text cells share one Qt clock-key mapping.
The table view starts editing the selected cell; an already open delegate editor
preserves the caret/selection. Ctrl+S retains existing validation/manual-submit
behavior and persists the actual inserted value. Read-only/non-text editors do
not mutate. ISO date/whole-second time formatting remains bounded.
Saving a table now restores the selected existing record/column by its complete
primary key after model reset, including when sort order moves that record.
Search is bounded to one million fetched rows; absent/NULL/new-generated keys
and complete large-table/no-key selection semantics remain open.
Main-window regression passes 1/1 (1.65 s), proving real date/time keys, persisted
SQL values, partial replacement and keyed restoration as sorting moves the same
record twice. The shared-helper form layout target also passed in the preceding
integration run. App builds; existing targets only. Last full suite/package/hash
baseline remains 8dc6afa; no complete datasheet/regional-format parity claim.
