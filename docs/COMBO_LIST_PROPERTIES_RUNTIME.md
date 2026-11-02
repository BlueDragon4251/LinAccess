# Native ComboBox ListIndex and ItemData

ComboBox.ListIndex reads the selected zero-based row of the actual native model;
an empty or unmatched selection returns -1. ComboBox.ItemData(index) reads the
specified row's bound-column QVariant, independently of the displayed column or
current selection. BoundColumn=0 returns the row ordinal; numeric/text/NULL and
zero-length text values retain the existing native Variant conversion scope.
The first read lazily loads the real row source. Both properties are read-only
(383 for native assignment); unknown control adapters use 438. ListIndex accepts
no arguments; ItemData requires one index (450 otherwise). The index scalar is
resolved once, rounded with the existing integer conversion, checked as signed
Long (6 overflow), with Null 94 and incompatible text 13. Invalid native row
indices use 5; exact Windows invalid-index error identity is not established.

ListIndex is limited to native Form view; report view rejects with 3251. Its
value during an uncommitted editable text draft also rejects with 3251, because
exact Windows selection timing has not been established. ItemData can still
inspect stored row data during that draft without committing it. This does not
implement the differently documented ListBox property or multi-selection API.
Full declared Long/Variant subtype identity remains open.

The regression exposed an older null-selection bug: string matching could pick
a NULL-key list row when assigning Null. Typed load/assignment/mirroring/requery
now pass null state into the widget value setter, clearing combo selection and
text for Null. A non-null empty string can still select its own empty bound row;
NULL rows are not matched merely because their string conversion is empty.
This is a bounded native normalization, not a claim about all Windows duplicate
or NULL-key selection/event timing.

Actual GUI/VBA/SQL tests cover lazy loading, selected/absent indices, numeric,
text and NULL ItemData, BoundColumn changes including zero, scalar evaluation,
rounding/error/read-only paths, report limits, pending drafts, row removal and
requery. Null and empty text yield distinct selected indices despite identical
empty displays. SQL records remain unchanged. Targeted main-window/layout/
original-heavy trio: 3/3 (40.11 s). No new original business execution is claimed.

Public references:
- https://learn.microsoft.com/en-us/office/vba/api/access.combobox.listindex
- https://learn.microsoft.com/en-us/office/vba/api/access.combobox.itemdata

ColumnHeads is now implemented within COLUMN_HEADS_RUNTIME.md. Its documented
ListCount includes a visual header row independently of selectable model data;
exact Windows behavior at that extra index boundary remains unverified.
