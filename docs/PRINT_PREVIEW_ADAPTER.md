# Native print-preview command bar subset

Imported Report layout previews now have a real Qt toolbar with current-layout
preview, printing, PDF output and close actions. Application.CommandBars exposes
that toolbar as the standard Print Preview bar. This adapts a small Office API
surface to native Qt; it neither starts COM nor implements the Office object model.

Supported:

- Application.CommandBars(name/index), the bare global CommandBars, Count and
  enumeration. The current Print Preview toolbar and actual toolbars of the
  enclosing Qt project window are exposed. Office indices start at 1.
- CommandBar.Name/Enabled/Visible and Controls(name/index), Count/enumeration.
- Command control Caption/Enabled/Visible and Type=msoControlButton (1), backed
  by actual QAction state. Unsupported Office bars, members and IDs fail.
- Weakly cached wrappers and QPointer guards for deleted toolbars/actions/dialogs.
- Qt separators and widget actions are excluded from this button-only adapter;
  they are not misreported as Office command buttons.
- Caption reads/writes for native label/button/check controls and form captions.
  Textboxes and shapes do not acquire a fabricated Caption property.

The print preview uses QPrintPreviewDialog and QPrinter. PDF writes are generated
in a temporary directory and committed with QSaveFile, require a .pdf destination,
and do not replace MDB/ZIP fixtures. The displayed canvas is scaled onto one page.
The UI explicitly calls this the current layout. It does not run Access group
formatting, multi-record pagination, subreports, aggregates or page events.

## Imported section geometry

The layout importer preserves section heights and offsets. Report layout previews
stack FormHeader (report header), PageHeader, Section (detail), FormFooter (report
footer), and PageFooter in that order. Page sections no longer overlap the detail
because their heights were discarded. Forms keep their existing header/detail/
footer arrangement. Negative heights and coordinate overflow are rejected.

The existing report-event test checks the preserved heights, native Qt positions,
and a detail textbox populated from a real SQL result. This fixes the displayed
single-detail layout only. It does not implement repeated detail bands, page
breaks, CanGrow, grouping, or Format/Print/Page event scheduling. The fixed work
catalog remains 456 entries; the broader report tasks stay open.

Report-footer placement follows the print-preview ordering described in
[Microsoft's report guide](https://support.microsoft.com/en-gb/access/guide-to-designing-reports).

## Fixed-band report PDF output

A separate **Bericht als PDF** action and `exportReportPdf` path now paginate the
loaded SQL rows. Report header is emitted once, page header/footer on each page,
detail once per row, and report footer after the last row (on another page when
necessary). The renderer uses physical TWIP-derived dimensions on QPrinter's
printable page, keeps detail bands whole, and restores the previously displayed
record after output. The current-layout actions retain their one-page behavior.

Calculated fields use the existing native Access expression engine, including
Sum/Count over fields and supported per-row expressions in the actual loaded
record source. Every
row is checked before opening the output device, so a calculation error refuses
the output instead of silently printing error placeholders. Compound arguments bind fields to each source row without moving the displayed
record, including arithmetic, IIf and NULL handling. Nested aggregates and
compound arguments on empty sources fail explicitly; group scope, running totals
and exact full Access aggregate typing remain open. Existing row expressions and the report-wide sum are
checked in the PDF as well as on restored native controls.

During each paginated render, at most 512 scalar aggregate results are cached.
This avoids rescanning the same row set for every printed detail. The cache is
cleared after output and on Requery; it is not a persistent module/procedure cache
and cannot retain object, DAO or widget references. A real SQL insertion followed
by native Requery and a second export checks that the total changes from 31 to 52.
The original corpus independently compares three available aggregate expressions
with SQL. The six original compound calls remain unexecuted in this prepared
corpus because their report-source context is unavailable.

Reference: [Microsoft's Sum function documentation](https://support.microsoft.com/en-gb/access/functions/sum-function)
describes expression arguments and NULL omission. These checks use native SQL as
an independent calculation, not a full Access-2003 output oracle.

Native Qt controls are recorded per section with QPicture and replayed as vector
commands. Recording only that section prevents duplicate controls; compensating
QPicture's device-DPI conversion avoids applying the print scale twice. Output
uses the existing temporary-file/QSaveFile path and does not overwrite the
destination after an unsupported or failed rendering attempt.

This is a fixed-band subset: Label, TextBox, Rectangle and Line, with direct field
bindings, supported calculated expressions, or unbound values. Truncated row sets, missing fields, other controls,
failed calculated expressions, overflowing controls/bands, groups, declared growing/
shrinking/repeated sections, running sums, hidden duplicates, multiple columns,
forced page breaks and section/page print events are rejected. It does not
implement Access page setup, grouping, expression pagination, CanGrow, subreports,
Format/Print/Page scheduling, or DoCmd's direct-print view. The current-layout toolbar
commands retain their one-page behavior. Separate Berichtsvorschau and Bericht
drucken actions use the same paginated renderer. The direct print dialog offers
the complete report; selection/page-range options are disabled there. Physical
printer output has not been validated in this environment.

The existing report-event target creates two pages from two real SQL rows and
checks that the displayed record is restored. Where `pdftotext` is installed,
it also verifies both values occur exactly once in order (Unicode ligatures are
normalized). It opens the real QPrintPreviewDialog and checks that its preview widget holds
two generated pages, then closes it and verifies the original displayed record.
It checks that overflowing and CanGrow sources cannot replace a previously
generated PDF. No additional CTest target or catalog item was added.

## Original source execution evidence

Three hash-fixed original Kassa Report_Open procedures execute in full through
the same Qt VbaHost path as application handlers, using the original report layout
and original unmodified module. They maximize the window, enable/configure the
Print Preview bar, enumerate its actual native actions and hide actions according
to Caption. The regression asserts that only Close remains visible, resets the
actions and reruns each complete handler. Their unused Cancel parameter is zero.

The test also asserts automatic bound Report_Open and Report_Close execution
through the navigator's shared preview path; explicit runtime entry calls rerun
the original opening body for action-state checks. Report data/event coverage is
detailed in REPORT_EVENTS.md. This does not prove complete original report
execution or print equivalence with Microsoft Access.

Each original layout also produces a temporary one-page PDF, checked for PDF
format and page count, then deleted. Tests verify that a non-PDF sentinel file is
unchanged on rejected export. No source data, VBA literals, PDF output or raw
SQL are emitted in test summaries or packaged artifacts. Additional synthetic
checks cover one-based indices, action caption lookup/identity, actual toolbar
state and access after action deletion.

Project-toolbar regression additionally exercises bare CommandBars enumeration,
real QAction caption writes, identity across global/Application lookups, toolbar
visibility and deletion guards. The large original Form_Load now completes with native embedded forms.
The native ActiveMenuBar adapter and mixed-Variant comparisons also complete
its OnOpen; complete startup is not
claimed (NATIVE_FORM_RUNTIME.md). Its diagnostic test rolls back changes on an isolated
MDB import before the report inventory runs. No original toolbar names or
complete Office menu hierarchy are fabricated.
