# Native detail RunningSum output

Access TextBox.RunningSum uses 0 (No), 1 (Over Group) and 2 (Over All):
https://learn.microsoft.com/en-us/office/vba/api/access.textbox.runningsum

The source decoder preserves textbox/template settings, rejects invalid values,
and exposes the report textbox property as native Byte. Runtime mutation remains
unsupported. Actual report preview, printing and PDF support Over All and bounded Over Group on
Detail textboxes bound to non-Null numeric fields, or the literal source =1 for
record numbering. Existing grouping and report-event safety bounds remain.

Each actual ordered source record contributes exactly once to a per-control
prefix. Existing checked numeric/Currency/Decimal sum state is reused. Page
planning, rendering repeated slices, refresh and repeated export read these
prefixes rather than accumulating again. Bound values, formats and calculated
controls referring to the running control see the prefix while rendering. The
original current record and control values are restored afterward. Export still
uses the existing temporary-file/atomic-target path.

Running aggregates in headers/footers, arbitrary calculated running
sources, Null source values and running CanGrow controls remain refused. This
is not full RunningSum, grouping, event order, Access numeric-promotion or Windows
rendering parity. F080/F085/F087 remain open; catalog totals stay 456/386/70.

Application and existing access_2003_layouts target built; its single bundled run
passed in 13.23 seconds. Actual multi-page PDFs contain exact four-place Currency
prefixes, =1 numbering and dependent calculated values. PDF text is compared
across repeated exports; control restoration and original PDF preservation after
a refused Null/group output are checked. No new CTest IDs were introduced.

## Native report authoring

The report designer now stores an optional running_sum_field alongside the
existing source, VBA, page setup, footer and display properties. It validates
the field against actual prepare-only result metadata without stepping rows or
executing user procedures. Missing fields fail 3265; unproven/non-numeric fields
fail 3251. The generated detail textbox receives RunningSum=2 by default; CanGrow remains
available on other fields. Reopening/editing retains the setting, including
through page-setup metadata merges. Existing designs without it retain ordinary
field output. Prefix allocation is bounded to one million cells per rendering.

The application and existing GUI/layout targets built. One bundled run passed
2/2 in 32.05 seconds (19.09/12.96). The actual native report editor saves and
restores the setting, opens the generated report and exports through the same
renderer. Existing PDF assertions additionally verify aligned record ordinals
1/2/3 alongside the Currency prefixes and dependent calculations. No new CTest targets
or catalog IDs were introduced. The complete F080/F085/F087 requirements remain
open. The checked 8fd1d50 package predates both running-sum scopes.

## Native Over Group authoring and output (2026-11-10)

The actual report designer now offers the persisted running_sum_mode1/2 next to
the numeric field. Existing definitions default to2. Invalid stored/new mode
types or values are rejected. Mode1 with a selected field requires an actual
group header/footer; missing group context refuses3251. Saving/reopening and
metadata-only page setup preserve the selected mode.

For the checked explicit native group contract, mode1 resets its precomputed
prefix at the beginning of the innermost level with a header/footer. Sort-only
levels are not treated as group sections. Each sorted source row contributes
once, so pagination/repainting/copies do not double-count. Existing non-Null
numeric/source=1, detail-only, no-running-CanGrow bounds remain; imported missing
group metadata and advanced group behavior remain guarded. This is no claim of
complete Windows RunningSum/event/layout equivalence.

App/GUI/layout built. The existing GUI bundle passed in12.75s with the unchanged
20s watchdog; only the layout bundle was repeated after correcting its PDF
regex to stay within a line (it had also captured the following page number).
It passed1/1 in10.19s. Actual native creation/reopening persists mode1; the
nested-group PDF contains all eight exact Currency prefixes across inner/outer
and Null-key groups, while group and overall totals remain correct. Restored
source order/current record and prior-file refusal checks remain active. No new
CTest/catalog identities;456/386/70. Latest checked package is app62d00929 plus
packaging correction5ec76beb, before SQL-CStr and this mode; original116.96s,
inventory49.42s, seven protected hashes unchanged.
