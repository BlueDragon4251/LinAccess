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
