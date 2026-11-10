# Report PageHeader and PageFooter settings

The layout importer retains the actual report-level PageHeader/PageFooter mode
(0-3), rejecting invalid values. Missing mode uses the documented All Pages0
default. Native report definitions and the real report designer persist the same
settings through page_header/page_footer, including reopen/edit/preview/PDF.

Mode1 omits the respective page band on a page containing the report header.
Mode2 omits it with the report footer and forces that footer onto a new page,
as documented by Access. Mode3 combines those rules. Either mode's footer bit
forces the footer page; the other page band still follows its own setting.
A missing report footer creates no artificial footer page. The native empty /
zero-height report-header interpretation uses actual nonzero rendered sections.
This is native output behavior, not a supplied Windows zero-height oracle.

Pagination uses the actual active page-band heights, including growing/split
details. Report header/footer space checks and error guards remain active. The
render pass restores the original record/control layout, and selected/ranged /
reversed/multiple-copy output continues to use recorded source pages. Layout
view retains the source design; page previews show the actual printed bands.
No report Format/Print event, grouped output, runtime VBA property mutation or
complete Access renderer equivalence is introduced by this scope.

App/GUI/report targets built. The existing report/layout check passed1/1 in10.72s:
actual PDFs cover all four modes and mixed header/footer settings, real first/
last markers, source page counts, one occurrence per data row, restored record,
and invalid source modes. GUI creation/reopen uses the new modes with explicit
ForceNewPage0 to prove the footer page comes from the new setting. The initial
GUI case retained an old expected ForceNewPage1; only that existing GUI check is
repeated after correcting the fixture and reusing its existing reopen assertion.
The corrected existing GUI check passed1/1 in12.07s with the unchanged20s
watchdog. The duplicate reopen added during implementation was removed; the
existing reopen check verifies the persisted new settings. No new CTest/catalog identities.

Documented public contract (not an Access execution oracle):
https://learn.microsoft.com/en-us/office/vba/api/access.report.pageheader
https://learn.microsoft.com/en-us/office/vba/api/access.report.pagefooter
