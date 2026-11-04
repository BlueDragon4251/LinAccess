# Bounded native CanGrow

Imported TextBox CanGrow values and templates are retained. Section CanGrow
metadata is retained too. The native getter returns the flag; assignment in a
runtime view fails explicitly because the property is design-only there. Owned
report design offers **Text vertikal erweitern (CanGrow)** and saves its setting.

The renderer measures the actual plain/masked text with the same Qt layout used
for painting. For growing direct-field/unbound textboxes it increases field heights and moves
lower controls. Growers with identical original top/bottom form an aligned group
whose greatest delta controls movement below it. Vertically disjoint groups add
their deltas, including for later growers. The detail expands across its width. Variable detail heights feed
the existing page plan. Other sections are hidden only for their captured band;
a scoped background override prevents later section colors leaking into a grown
detail. Control geometry/visibility, canvas size/background, record/page state
and aggregate cache are restored on success and failure.

The subset refuses partially staggered/overlapping vertical growth groups, overlapping controls,
calculated-source growers, growing page/report headers/footers, contradictory
explicit section settings and details beyond the existing 3500-pixel canvas bound.
Details fitting a printable page stay whole. Taller details are divided between
actual native glyph/decorative-line bounds, repeating page headers/footers and
retaining the row context. Fixed controls remain indivisible. If columns/control
spans leave no safe boundary, output is refused. This is bounded long-detail
splitting, not full KeepTogether/group scheduling or arbitrary-height streaming.
CanShrink and full multi-control flow remain open. A read-only getter plus bounded flow is not full Access report or
Windows font/geometry equivalence. No original CanGrow result oracle was run.

Qt's normal single-line display string replaces paragraph characters. Unmasked
normal report fields therefore use their original text; password/other input-mask
fields keep the authoritative display string. The native raw value is unchanged.
An earlier multiline probe happened to wrap where its paragraph break was; the
long-flow probe exposed and corrected this normalization issue. Masked-output
checks remain active.

The existing report case uses real SQL multiline rows that occupy one page with
fixed fields and two pages with growth. Additional parallel/stacked cases extract
both complete fields on the expected pages, checking max-versus-sum flow. PDF extraction independently checks both
final text markers and the calculated values below them. It checks restored
geometry/value/canvas and refusal beyond the canvas bound without replacing an
existing PDF. A single SQL row with 125 numbered paragraphs, then the same data
as automatically wrapped underlined text, spans pages; all tokens occur exactly
once in order and Page/Pages contains the real total. The painter emits only
lines with visible ink in the slice: clipping a full QTextDocument alone left
invisible duplicate text commands in PDF extraction. Fragment clipping and
line selection are scoped to recording and cleared before restoring the view. The existing mainwindow case persists the design setting and reads
it through actual Report VBA. No new CTest target or catalog ID.

Reference: [Microsoft TextBox.CanGrow](https://learn.microsoft.com/en-us/office/vba/api/Access.TextBox.CanGrow)
documents vertical growth, movement of lower controls, full-width section growth,
overlap limits and design-view-only assignment.
