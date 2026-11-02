# Native ComboBox ListRows

ListRows is retained from SaveAsText control templates and explicit instances,
with the documented default of 16 and valid range 1–255. Actual native ComboBox
widgets receive that maximum popup-item setting at construction. VBA reads the
real Qt setting; writes update the retained definition and widget without loading
or replacing the row model, validating pending text or saving a record. Changing
the cap applies on the next popup opening. Reports retain static/read access;
runtime report mutation explicitly rejects with 3251.

The setter resolves scalar defaults once, uses the existing rounded integer
conversion and checks signed Integer overflow (6 outside -32768..32767), then the
property range (5 outside 1..255). Null/type failures use 94/13 and preserve state;
getter arguments use 450 and unsupported control adapters 438. Numeric results
retain the existing generic Number scope, not full declared Integer subtype
identity. Invalid imported values fail before widget construction.

The GUI regression checks actual popup viewport row geometry at caps 3 and 5,
vertical scrolling through a 40-row SQL model and real selection/persistence of
a row beyond the initial viewport. It also verifies template/default metadata,
VBA writes, unchanged model identity, retained pending text/dirty state, conversion/
error/scalar-once behavior and report/adapter limits. Initial targeted layout
suite: 1/1 (9.70 s). Full validation is in PROGRESS.md. No new original business
execution coverage is claimed.

This is the native Qt data-row cap. Exact Windows heading-row interaction,
screen-edge clamping, an already-open popup's resize timing, alternate themes,
completion popups, pixel/scrollbar fidelity and report context identity remain
open. ListRows does not truncate the SQL model or claim full combo compatibility.

Reference: https://learn.microsoft.com/en-us/office/vba/api/access.combobox.listrows
