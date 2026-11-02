# Native literal form search

Ctrl+F and documented RunCommand acCmdFind (30) open a real modeless Qt search
dialog for complete native direct-table forms with declared text fields. It
searches a selected text column or all text columns, with any-part/whole-field/
field-start and case options, advances/wraps through the displayed record order
and navigates actual bound records. Text is explicitly searched literally; the
UI does not imply wildcard support. Empty queries disable Find Next; NULL/binary
values do not become false text matches. Long query/editor text is not truncated
by Qt's default capacity.

Field validation applies before opening/reusing/searching. Accepted current
drafts participate without an implicit record save when matching the current
record or finding nothing. Moving to another match uses native save/navigation
and re-identifies the target by its row ID. The dialog is reused, stays open on
Unload cancellation and closes with its source document. Native lifetime guards
protect callbacks. A background VBA caller targets the active document rather
than its own hidden form. Masked editor selections are not prefilled visibly.

The fixture-independent find_form_documents CTest verifies Unicode/case/modes,
all-field advance, literal wildcard text, long text, Return/default-button search,
reuse, active Screen/RunCommand context, masked-prefill protection, both open-time
and already-open validation cancellation, unsaved draft/no-match behavior, real
SQL save on navigation and parent Unload/close lifetime. Targeted copy/delete/
long-text/find tests pass 4/4 (1.02 s). A synthetic-only private screenshot was
visually inspected; no original project data/source was exported. GUI suite now
has 27 CTests. Full validation is in PROGRESS.md. Original business execution
coverage is unchanged.

This is literal text search, not complete Access Find/Replace parity. Wildcards,
replace, formatted numeric/date fields, complete query/report/datasheet sources,
Jet locale/selection/highlighting/search-direction semantics, concurrency and all
callback mutations, global utility/MDI keys and pixel-identical Access 2003 UI
remain open. Truncated or unsupported sources explicitly diagnose unavailable
search rather than claiming complete results.

References:
- https://support.microsoft.com/en-us/office/keyboard-shortcuts-for-access-70a673e4-4f7b-4300-b8e5-3320fa6606e2
- https://learn.microsoft.com/en-us/office/vba/api/access.accommand
