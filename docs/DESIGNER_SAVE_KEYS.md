# Native designer Ctrl+S and classic menu organization

Ctrl+S in owned query/form/report editors uses their actual validated save path
and keeps the editor open. Successful saves show the saved object name and update
the native editor context. The existing explicit save button retains its prior
save-and-close behavior. The shortcut is scoped to the dialog and its child widgets;
no global Ctrl+S handler intercepts unrelated/modal file pickers or runtime forms.
Runtime record saving and project Ctrl+N/Ctrl+O guards remain their existing paths.

After an initial save, subsequent saves merge from the current saved definition.
Query editing now preserves unrelated JSON members; report editing also retains
page geometry saved from a nested preview after the editor's first Ctrl+S save.
Controlled UI fields still provide SQL/VBA/geometry/caption values through the
existing generators. This is not complete concurrent property-sheet merging or a
full Access design-view save/dirty/rename/undo pipeline.

Main menus now order Datei, Bearbeiten, Ansicht, Einfügen, Fenster. Existing object
creation actions remain connected to the real builders in Einfügen; form/report
edit actions move to Ansicht with selected-object enablement updated there.
New/Open actions expose actual WindowShortcut standard key sequences, while
existing preview KeyPreview, modal and project-switch guards remain in force.
Missing Access menus/actions are not filled with nonfunctional placeholders;
full Access2003 menu/context/pixel/keyboard parity still requires further work.

Existing real UI cases send Ctrl+S to activated SQL/VBA text editors, independently
read saved database definitions and verify each editor remains open. Query SQL is
changed and saved again, an externally added archive marker survives the merge,
and Execute produces the actual new value43 in the Qt result model. Existing
project/window/KeyPreview/modal-switch and report page-geometry/event cases remain.
No new CTest target or fixed-catalog item. Activating the modal window before key
injection is necessary in the offscreen fixture; production code does not bypass
keyboard or validation to satisfy the check.

## Actual native editor modified state

Owned query/form/report editors now track changes after initialization: text,
source choices, checkboxes, spin values and field-list/table model mutations mark
the actual Qt window modified. The saved-status label changes to
`Nicht gespeicherte Änderungen`; the title contains the native Qt modified marker.
A new unsaved editor starts modified. Successful actual saves reset modified/new
state and show the persisted object name. Failed validation/save does not execute
the successful-reset path. Query result models are not design-input models and do
not mark the editor modified merely because Execute refreshes results. Existing
form geometry selection uses blocked signals and does not invent a geometry edit.

This is change tracking, not baseline equality after arbitrary undo, complete
Access object rename/new-name reservation, design locking, shutdown save prompting
or a full design-view Undo pipeline. Existing core data-record dirty state is
separate. Native editor state is tagged for future checked API consumption.
Existing mainwindow1/1 (2.71 s), current Debian/build-ID/SQLite3.54.1 and four original
hashes pass. Latest full baseline48/48 (83.85 s), headless runtime unchanged1/1
(1.75 s); no extra target or catalog item. F080 remains open;456/384/72.
