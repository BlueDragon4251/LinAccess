# Editable physical query field projections

The native query datasheet edit path now accepts explicit physical field lists
from one real keyed main table, including a table alias, per-field AS/implicit
aliases, a bounded physical WHERE predicate and up to ten source sort fields.
The caller requests an actual projection mapping. Existing consumers that only
ask for a full-table source keep their previous star-only contract; this does
not silently widen report/form/DAO updateability.

Each visible field resolves once to a real physical model column. Repeated
physical columns, ambiguous output names, expressions, DISTINCT/ALL, joins,
parameters, virtual/TEMP-shadowed/generated sources and unsafe keys stay outside
this edit path. Native SQLite identifier matching is ASCII-case-insensitive;
other Unicode characters must match exactly. Ambiguous Qt-folded physical names
refuse3251 instead of selecting another column. The checked original SELECT is
prepared without stepping it; no user row/expression is evaluated for metadata.

The physical model retains its real key even when omitted from SELECT. UI edits
save through that key; unselected fields are hidden and retain their values.
Visible order/current field follow the actual projection, and explicit different
aliases override the header. Compiler-added same-name aliases retain the native
field caption. Unqualified ORDERBY output names map to their physical field;
a qualified sort name denotes the source field. Header/visibility bindings are
reapplied after model reset/sort/Requery; aliases are cleared when changing the
underlying model table. Visible field search uses header order.

The subsequent native clipboard implementation maps visible columns for copy,
cut, paste and Undo/Redo, including reordered nonadjacent fields; see
PROJECTED_GRID_CLIPBOARD.md. Hidden keys/omitted data stay out of projected row
copy. Other native clipboard/format/limit guards remain.
Full joined/projection/form-recordset/Access caption/identifier equivalence is
not claimed; F079/F175/F340/F373 remain open.

App built. Existing GUI target passed18.42s with actual stored query opening,
reordered aliased subset, hidden key/other fields, visible search, real SQL save
by hidden key, unchanged omitted values, Requery/layout retention and protected
reordered paste. The small metadata/model target exposed compiler-added AS id;
normalization then passed0.10s. Additional native Unicode mapping/collision guards
passed the affected model target0.11s. Only that target was repeated after those
metadata changes; no new CTest IDs/Original/Windows/full-suite/package run.
Fixed catalog456/386/70 is unchanged.
