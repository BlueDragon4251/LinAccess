# Native detail RunningSum output

Access TextBox.RunningSum uses 0 (No), 1 (Over Group) and 2 (Over All):
https://learn.microsoft.com/en-us/office/vba/api/access.textbox.runningsum

The source decoder preserves textbox/template settings, rejects invalid values,
and exposes the report textbox property as native Byte. Runtime mutation remains
unsupported. Actual report preview, printing and PDF now support Over All on
Detail textboxes bound to non-Null numeric fields, or the literal source =1 for
record numbering. Existing grouping and report-event safety bounds remain.

Each actual ordered source record contributes exactly once to a per-control
prefix. Existing checked numeric/Currency/Decimal sum state is reused. Page
planning, rendering repeated slices, refresh and repeated export read these
prefixes rather than accumulating again. Bound values, formats and calculated
controls referring to the running control see the prefix while rendering. The
original current record and control values are restored afterward. Export still
uses the existing temporary-file/atomic-target path.

Over Group, running aggregates in headers/footers, arbitrary calculated running
sources, Null source values and running CanGrow controls remain refused. This
is not full RunningSum, grouping, event order, Access numeric-promotion or Windows
rendering parity. F080/F085/F087 remain open; catalog totals stay 456/386/70.

Application and existing access_2003_layouts target built; its single bundled run
passed in 13.23 seconds. Actual multi-page PDFs contain exact four-place Currency
prefixes, =1 numbering and dependent calculated values. PDF text is compared
across repeated exports; control restoration and original PDF preservation after
a refused Null/group output are checked. No new CTest IDs were introduced.
