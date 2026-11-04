# Native detail ForceNewPage

The imported detail section retains ForceNewPage 0 (none), 1 (before), 2 (after)
and 3 (before and after). Invalid values fail decoding. The native page plan
separates actual detail records and honors a pending after-break before a report
footer. It does not create an empty trailing page merely to satisfy a pending
break. Repeated page headers/footers and Page/Pages use the resulting plan.
A before-break separates the first detail from an existing report header; an
already fresh detail page does not gain a redundant empty page.

Owned report design provides all four choices and persists detail_force_new_page.
The adapter validates its integral 0–3 value and emits actual section metadata;
opening and PDF use the same native renderer. Missing values retain legacy none.
This is detail-only: report/group header/footer ForceNewPage, page setup, section
VBA objects, runtime changes and the complete formatting/event schedule remain
open. Nonzero unsupported-section settings still refuse rendering. No Windows
Access output oracle or universal trailing-page equivalence is claimed.

The existing report case verifies all three nonzero settings with real SQL rows,
actual native page counts and extracted row values on the expected pages. The
existing mainwindow case chooses, saves, reopens and previews before/after breaks
through the actual report editor. No additional CTest target or catalog point.

Reference: [Microsoft Section.ForceNewPage](https://learn.microsoft.com/en-us/office/vba/api/access.section.forcenewpage)
documents the four settings and exclusion of page headers/footers.
