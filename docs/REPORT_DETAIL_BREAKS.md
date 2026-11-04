# Native report body ForceNewPage

Imported report headers, details and report footers retain ForceNewPage 0 (none),
1 (before), 2 (after) and 3 (before and after). Invalid values fail decoding.
The native plan handles repeated detail records, a report-header after-break
for a title page and a report-footer before-break for a final summary page.
Concurrent pending/before breaks coalesce on an already fresh body page.
Page headers/footers repeat and Page/Pages uses the actual resulting plan.

The report starts on a fresh page, so before the report header needs no empty
leading page. An after-break with no following body content does not manufacture
an empty trailing page. Nonzero header/footer breaks on a zero-height section
are refused rather than inventing their exact Access blank-page behavior.
Group sections, section events and Windows trailing/zero-height page equivalence
remain open. Nonzero page-header/footer settings remain unsupported; that
property does not apply to page sections in the documented Access model.

Owned report design provides all four choices for detail and report header,
persisting detail_force_new_page and header_force_new_page. The adapter validates
integral 0–3 values and emits actual section metadata. Missing values retain none.
The current automatic owned layout has no report-footer designer.

The [native Section object](NATIVE_SECTION_OBJECT.md) exposes bounded reads and
runtime writes for all three supported report body sections. Existing native
previews rebuild after a valid write; mutation during rendering is refused.
This does not implement full page setup, grouping or the complete event schedule.

Existing SQL/PDF cases check all nonzero detail choices, title/final pages both
separately and together, and extracted row/summary placement. Actual typed VBA
sets/resets header/detail/footer breaks; real mainwindow design/save/reopen/preview
checks both exposed selections. No additional CTest target or catalog point.

Reference: [Microsoft Section.ForceNewPage](https://learn.microsoft.com/en-us/office/vba/api/access.section.forcenewpage).
