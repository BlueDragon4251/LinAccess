# Native lookup keyboard refresh

F9 now reloads the actual row model of a focused native combo adapter before
recalculating form expressions. It works with focus on either the combo or its
editable text field. Other focused controls retain the existing F9 recalculation
behavior; Shift+F9 keeps its separate underlying form/subform requery path.
Native KeyPreview/KeyDown runs before these actions. The F9 path retains native
execution depth through callbacks and skips recalculation on a closed document.

Row refresh uses the existing real SQL/value/field-list model loader. It does
not save or validate an unfinished typed label as a numeric bound value. A
successful refresh preserves pending text, its modified flag, focus and dirty
state. Query failure preserves the previous model and draft, reporting the
existing native status error; a later F9 retries and recovers. RowSource/property
contracts retain their documented limits. F4 popup opening uses the existing
Qt combo behavior after native dispatch, with cancellation tested explicitly.

A synthetic SQL-backed regression checks cancelled F4/F9, successful F4 opening,
external label/row changes, model replacement, retained bound ID, pending draft,
failed refresh after dropping the choices table, recreation/recovery and Escape
undo. Actual stored SQL remains unchanged until an explicit save. Targeted
main-window/layout/original-heavy trio: 3/3 (41.19 s). No additional original
business execution is claimed.

The current ListBox adapter still uses the bounded combo-backed renderer; this
does not implement an Access multi-selection list box. Popup-focused keyboard
routing, exact Windows cursor/selection/event timing, every lookup mode and
complete visual/keyboard parity remain open.

Reference: archived Access 2003 keyboard help,
https://documentation.help/MS-Office-Access-2003/AccessKeyboardShortcuts.htm
