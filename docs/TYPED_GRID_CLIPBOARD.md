# Native typed grid clipboard

Native table Copy publishes a bounded application/x-linaccess-grid-v1 JSON
rectangle alongside its ordinary quoted TSV representation. The native payload
retains SQL NULL independently of empty text, canonical signed 64-bit integers,
finite real values, booleans and binary byte arrays. Binary TSV uses hexadecimal
0x text; this is a presentation for other applications, not an implicit binary
input parser. Existing typed text/Currency/Date clipboard presentations remain.

Paste prefers the native payload when present. It checks version, shape, cell
kinds, canonical integer/base64 encodings, size limits and target field types
before changing a clean manual-submit cache. A malformed native payload fails
instead of falling back to its text representation. Native binary input is
permitted only into binary columns; ordinary text cannot write binary fields.
Generated columns, readonly models, hidden target cells and unknown metadata
remain protected. Boolean values use the existing native logical conversion;
negative-zero real payloads are refused rather than silently losing their sign.

The same pending-range implementation supports existing rows, extra rows,
empty tables and explicit append. SQL Save keeps the existing transaction and
counter generation path. This adds no Binary Variant to VBA and no original MDB
binary extractor. Native form-record MIME remains a separate protocol. Full
Windows/Access clipboard interoperability, pasting over unrelated drafts,
regional parsing and complete range undo remain open under F340/F168/F016.

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

Native grid payloads also record the selected physical field names. Explicit
append recognizes a full row only when its field names and order match the full
target schema. Such rows append from the first physical column, even with an
AutoWert selection, and ignore only proven generated columns. Save obtains actual
new keys through the existing counter/insert path; supplied manual keys retain
the normal database constraints. Full-row shapes/names are checked before drafts
change. The generated-field policy survives pending block Undo/Redo. Older grid
payloads without field names retain positional range semantics. Form-record MIME
remains separate; differently ordered/partial named record mappings remain open.

Clean native table datasheets can Cut a contiguous visible cell rectangle with
Ctrl+X, Shift+Delete, the Edit/context menus or RunCommand Cut(189). Copy validates
and publishes the typed values before the cells become pending SQL NULL values.
Generated fields/readonly sources are protected; an unrelated dirty cache is
refused. SQL remains unchanged until normal transactional Save. The same bounded
block Undo/Redo restores or reapplies the pending clear. This clears cells; it
is not record deletion or a claimed full Access Cut/clipboard event protocol.
