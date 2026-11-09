# Native typed grid clipboard

Native table Copy publishes a bounded application/x-linaccess-grid-v1 JSON
rectangle alongside its ordinary quoted TSV representation. The native payload
retains SQL NULL independently of empty text, canonical signed 64-bit integers,
finite real values, booleans and binary byte arrays. Binary TSV uses hexadecimal
0x text; this is a presentation for other applications, not an implicit binary
input parser. Existing typed text/Currency/Date clipboard presentations remain.

Paste prefers the native payload when present. It checks version, shape, cell
kinds, canonical integer/base64 encodings, size limits and target field types
before changing a manual-submit cache. Existing pending updates to physical
records can be preserved, including values outside the target rectangle. Pending record insertions/deletions can also be captured and restored; a deleted
row itself cannot be a writable range target. A malformed native payload fails
instead of falling back to its text representation. Native binary input is
permitted only into binary columns; ordinary text cannot write binary fields.
Generated columns, readonly models, hidden target cells and unknown metadata
remain protected. Boolean values use the existing native logical conversion;
negative-zero real payloads are refused rather than silently losing their sign.

The same pending-range implementation supports existing rows, extra rows,
empty tables and explicit append. SQL Save keeps the existing transaction and
counter generation path. This adds no Binary Variant to VBA and no original MDB
binary extractor. Native form-record MIME remains a separate protocol. Full
Windows/Access clipboard interoperability, general pending-record merges, regional parsing and complete range undo remain
open under F340/F168/F016.

Tables with no SQL defaults do not reconstruct unrelated DAO column definitions
just to evaluate an empty default set. This allows supplied native grid values
in generic SQLite BIGINT/BLOB tables without pretending those types are native
DAO field declarations. Tables requiring actual default evaluation retain the
existing supported-schema checks.

An unchanged pending pasted block has one native Undo/Redo step. Ctrl+Z/Ctrl+Y,
the existing Edit menu and RunCommand Undo(292)/Redo(199) share that step outside
a cell editor. Undo changes the cache only, including removal of added rows;
Redo restores checked values and the captured native Date defaults. Focused text
editors retain their own existing text undo. Any unrelated cell/row modification,
select/requery or successful Save invalidates block history. Readonly models and
missing/invalidated history refuse the operation. A failed SQL Save retains the
pending block; this is not a general SQL undo or a multi-operation Access stack.

Native grid payloads also record selected physical field names. Explicit append
maps supplied known names to target columns, independently of their order and
including partial named records. Unknown, duplicate or ambiguous names refuse.
Omitted fields retain SQL defaults; supplied NULL remains NULL even when the
field has a default. Only proven generated fields are ignored and obtain new
keys on normal Save; supplied manual keys retain their constraints. The supplied-
field/default/generated policy survives block Undo/Redo. Hidden omitted/generated
fields are not clipboard write targets; supplied hidden fields remain protected.
Normal range paste and older payloads without names retain positional semantics.
Form-record MIME and general external clipboard interoperability remain separate.

Clean native table datasheets can Cut a contiguous visible cell rectangle with
Ctrl+X, Shift+Delete, the Edit/context menus or RunCommand Cut(189). Copy validates
and publishes the typed values before the cells become pending SQL NULL values.
Generated fields/readonly sources are protected; an unrelated dirty cache is
refused. SQL remains unchanged until normal transactional Save. The same bounded
block Undo/Redo restores or reapplies the pending clear. This clears cells; it
is not record deletion or a claimed full Access Cut/clipboard event protocol.

For range paste over pending physical-record updates, history captures the exact
preexisting raw values and changed-field positions. Undo restores those pending
updates rather than discarding them; Redo reapplies the checked range on that
state. A failed range mutation uses the same cache restoration path. No SQL
transaction is used to simulate cache changes, and no automatic Save is added.
The capture is bounded to one million loaded cells. It preserves inserted
records, their supplied/default-field state and prior deletion markers. Cut retains its clean-cache
requirement. General Access multi-operation undo remains open.

Cache restoration recreates pending inserted rows at their actual positions,
then restores raw values and reestablishes pending deletions. Proven generated
fields are regenerated on Save; omitted fields retain their SQL default behavior.
Undo/Redo restores a valid current cell when row reconstruction invalidates Qt
indexes. Record events, arbitrary external QModelIndex ownership, concurrency
and general Windows/Access undo equivalence remain outside this bounded scope.
