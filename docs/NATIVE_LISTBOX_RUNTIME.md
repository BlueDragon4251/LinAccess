# Persistent native single-selection ListBox

Imported single-selection ListBox controls now use an actual persistent QTreeView,
not a combo popup or an editable text replacement. The visible view consumes the
same real typed QStandardItemModel row-source loader as ComboBox. Table/query,
Value List and Field List sources, BoundColumn, ColumnCount, hidden ColumnWidths,
Column, ListCount, dynamic source properties and bound/unbound values use that
shared SQL/model path. List rows load before Open so the persistent view is
populated without a preceding Dropdown call. Failed initial sources show the
existing native error, retain the control and can recover through F9/Requery.

Single row selection works through actual mouse and arrow-key input. The native
control buffering/BeforeUpdate/AfterUpdate path validates selection; Click uses
the imported native event binding. Explicit record saving persists the typed
bound value. Row refresh and source changes do not implicitly save records.
Value assignments and Undo synchronize the visible selection, including clearing
selection for Null. Font/foreground/focus/tab/visible/enabled state uses the
existing native widget paths. Locked is available for this ListBox adapter;
locked/report/read-only rows block user value changes while retaining display.
Old owned models are released on replacement, and obsolete Qt selection models
are deferred for disposal through Qt ownership.

SaveAsText MultiSelect is retained, including control templates and explicit
zero overrides. Native Simple/Extended selection, Selected/ItemsSelected and
ListBox ItemData/ListIndex getters are now in MULTISELECT_LISTBOX_RUNTIME.md.
Their value/persistence, requery and event limits remain explicit; the separate
ComboBox read-only property contract is not generalized blindly.

Actual GUI/VBA/SQL regressions verify persistent geometry/model identity with
no popup, hidden/visible widths, typed NULL columns, mouse/arrow selection,
control events/cancelled saving, SQL persistence, Locked, F9/source replacement,
failed-source recovery, Null/Undo, unbound defaulted Value Lists, Field Lists,
report read-only behavior and MultiSelect modes. Metadata tests cover
template inheritance and zero override. Targeted main-window/layout/original-
heavy trio: 3/3 (40.87 s); final complete validation is in PROGRESS.md.
Original business execution coverage is unchanged.

This is bounded single-selection presentation, not complete Access ListBox or
visual parity. Bounded ColumnHeads is now in COLUMN_HEADS_RUNTIME.md. Field
captions and complete multi-selection/VBA behavior,
exact Windows click/validation/focus order, all row-limit/locale semantics,
background styling, accessibility and real window-manager/print fidelity remain
open. Native reports share the visible read-only view; this does not prove full
Access report pagination or printing compatibility.
