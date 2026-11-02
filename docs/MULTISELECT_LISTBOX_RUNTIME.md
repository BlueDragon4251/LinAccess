# Native ListBox selection API and multiple selection

SaveAsText MultiSelect=0/1/2 now creates actual single/Simple/Extended QTreeView
selection modes. Simple mouse clicks and Space toggle rows; arrow movement changes
focus without changing selection. Extended supports ordinary, Ctrl and Shift
mouse selection and Shift-arrow range extension. The existing native keyboard
route still runs before the view handles input. Native Locked/report/read-only
input guards retain their current scope.

ListBox.Selected(row) reads/writes real selected state. Programmatic writes do
not synthesize user control events. For single selection, writes use the existing
typed Value/dirty path; multiple selection changes only view state. ItemData(row)
reads typed bound-row values; ListIndex reads the current/focused native data row.
Its runtime writer remains unsupported (3251), rather than generalizing the
separate ComboBox read-only contract. MultiSelect is readable but runtime mutation
rejects with 3251, following the documented design-view-only setting restriction.

ItemsSelected is a live weakly owned collection of numeric row indices. Count,
Item/default indexing and For Each operate on actual selected rows in ascending
native data-row order. Enumeration snapshots indices; changes during enumeration
are not claimed to match COM collection behavior. Collection writes reject with
383, unknown members with 438, and references reject with 2450 after form close.
Native row indices resolve scalar defaults once, use existing Long conversion/
rounding, Null/type/overflow errors 94/13/6, arity 450 and actual-data bounds 5.
Exact Windows boundary/ordering/subtype identity remains unverified.

VBA object assignment now supports indexed named members such as
Me.pick.Selected(CLng(index))=True, including nested argument expressions,
bracketed names and With. A balanced target scanner distinguishes strings,
brackets and nested calls. The object, indices and right value evaluate once;
assignment uses an explicit setter protocol and never calls a getter pretending
to assign. Unsupported object setters reject with 438. Invalid assignment targets
do not fall through to a getter call; ordinary method-call parsing retains its
existing scope. This does not implement VBA arrays, classes or a full AST.

Multi-selection Value/OldValue stay Null even when the control is bound. Selection
does not dirty a form or allocate a new record. Record loading, dirty comparison,
control mirroring and record persistence exclude these non-value controls, so
saving another field cannot replace their underlying bound field with Null.
A second value control can explicitly update that same field without changing
selected rows. Direct multi Value assignment rejects with 3251; use Selected.
User selection still uses the native control BeforeUpdate/AfterUpdate/Click path,
including cancellation. Exact Windows multi-control event timing remains open.

Extended Requery replaces the model and clears selections as documented. For
Simple mode with selected rows, Requery/source replacement currently rejects
with 3251 while preserving the previous model/selection, because exact preservation
semantics are not established. Clear Selected first to reload. Multi Column
requires an explicit data-row argument (3251 otherwise); explicit Column and
ItemData work independently of the Null Value. Selection APIs are Form-view scoped;
reports retain read-only display/Null Value and pure ItemData access, while
selection-state access rejects. Header/index contracts retain COLUMN_HEADS_RUNTIME.md
limits, including the unverified extra ListCount header boundary.

Actual GUI/VBA/SQL regressions cover Simple and Extended mouse/keyboard behavior,
programmatic indexed setters, scalar-once, collection/default access and For Each,
read-only/bounds/context errors, cancelled validation, Extended clearing/Simple
requery guard, weak lifetime, single-selection setters, unchanged records and
zero inserts on a new row. Both modes save other fields and an explicit same-field
editor without corrupting bound values or selected state. Core tests cover nested
indexed targets, With/bracket/quoted-equals syntax, single evaluation, invalid
lvalues and default unsupported setters. Targeted core/main-window/layout/original-
heavy set: 4/4 (40.66 s). Full validation is in PROGRESS.md. Original business
execution coverage is unchanged.

References:
- https://learn.microsoft.com/en-us/office/vba/api/access.listbox.multiselect
- https://learn.microsoft.com/en-us/office/vba/api/access.listbox.selected
- https://learn.microsoft.com/en-us/office/vba/api/access.listbox.itemsselected
