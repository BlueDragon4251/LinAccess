# Native list column headings

SaveAsText ColumnHeads is retained as a Boolean, including per-control templates
and explicit false overrides. Native ComboBox and single-selection ListBox expose
its VBA getter/setter. Runtime changes rebuild the real row model and header
view, retaining the bound value, pending combo draft and form dirty state through
the shared loader. The new Boolean setter resolves scalar defaults once; native
Null/type failures use 94/13 and preserve the previous model/metadata. Runtime
report assignment rejects with 3251; static report headers/read access work.

Table/Query sources use the actual projected SQL field names/aliases as headers.
Value Lists use their first ColumnCount-sized data row as headings, removing it
from selectable model rows. Header labels are displayed by the real QHeaderView,
including hidden-column behavior; they are not selectable data items. BoundColumn
zero remains the zero-based data ordinal in this bounded native model.

The documented ListCount rule includes a heading row: a Table/Query with N rows
returns N+1 when headings are enabled, including an empty query returning 1.
A Value List uses its existing first row, so turning headings on leaves its total
ListCount unchanged while reducing selectable rows by one. An empty Value List
has no supplied heading row and returns zero. Actual model data rows and the
visual heading count are kept separately; no fake data row is inserted.

Column/ItemData/ListIndex retain their documented bounded native data-index
contracts. In particular invalid explicit indices still use the existing native
5; exact Windows behavior at ListCount's extra heading boundary is unverified.
Field List heading semantics explicitly reject with 3251, preserving prior state.
Field captions beyond projected SQL names, complete Access header sizing,
Windows report mutation/printing/error order and all header/index edge behavior
remain open. These are scoped native contracts, not a Windows Access oracle.

Actual GUI/VBA/SQL tests cover both ComboBox popup and persistent ListBox headers,
SQL alias labels, header clicks without selection, hidden columns, VBA toggling,
scalar-once and rollback errors, unchanged SQL despite dirty form state,
Table/Query and Value List counts including empty sources, BoundColumn zero,
Column/ComboBox ItemData and ListIndex data access, template overrides and static
report/read-only runtime context. Initial targeted layout test: 1/1 (9.01 s).
Complete validation is in PROGRESS.md; no new original business execution claim.

References:
- https://learn.microsoft.com/en-us/office/vba/api/access.listbox.columnheads
- https://learn.microsoft.com/en-us/office/vba/api/access.combobox.listcount
- https://learn.microsoft.com/en-us/office/vba/api/access.combobox.column
