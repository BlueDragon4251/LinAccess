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

Successful native paste, Cut and staged deletion actions share bounded multi-action
Undo/Redo. Ctrl+Z/Ctrl+Y, the existing Edit menu and RunCommand Undo(292)/Redo(199)
use this history outside
a cell editor. Undo changes the cache only, including removal of added rows;
Redo restores checked values and the captured native Date defaults. Focused text
editors retain their own existing text undo. Any unrelated cell/row modification,
select/requery or successful Save invalidates block history. Readonly models and
missing/invalidated history refuse the operation. A failed SQL Save retains the
pending history; committed SQL and arbitrary Access operations are outside this scope.

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

Native table datasheets with supported pending caches can Cut a contiguous visible cell rectangle with
Ctrl+X, Shift+Delete, the Edit/context menus or RunCommand Cut(189). Copy validates
and publishes the typed values before the cells become pending SQL NULL values.
Generated fields/readonly sources are protected; a pending cache beyond the supported snapshot bound is
refused. SQL remains unchanged until normal transactional Save. The same bounded
block Undo/Redo restores or reapplies the pending clear. This clears cells; it
is not record deletion or a claimed full Access Cut/clipboard event protocol.

For range paste over pending physical-record updates, history captures the exact
preexisting raw values and changed-field positions. Undo restores those pending
updates rather than discarding them; Redo reapplies the checked range on that
state. A failed range mutation uses the same cache restoration path. No SQL
transaction is used to simulate cache changes, and no automatic Save is added.
The capture is bounded to one million loaded cells. It preserves inserted
records, their supplied/default-field state and prior deletion markers. Cut uses the same bounded cache restoration
requirement. General Access multi-operation undo remains open.

Cache restoration recreates pending inserted rows at their actual positions,
then restores raw values and reestablishes pending deletions. Proven generated
fields are regenerated on Save; omitted fields retain their SQL default behavior.
Undo/Redo restores a valid current cell when row reconstruction invalidates Qt
indexes. Record events, arbitrary external QModelIndex ownership, concurrency
and general Windows/Access undo equivalence remain outside this bounded scope.


Readonly-Abfragedatenblätter veröffentlichen jetzt neben TSV das vorhandene typisierte
Grid-MIME mit Feldnamen. Tatsächliche rohe Querywerte erhalten Binärbytes, NULL,
Leertext/leeres BLOB und volle SQLite-Ganzzahlen; TSV verwendet Hex-Binärtext und
exakte Ganzzahldarstellung. Proveniente CURRENCY_TICKS-Ergebnisse erhalten ein
getrenntes Currency-Tag mit kanonischen 64-Bit-Ticks, keine Umdeutung als Ganzzahl.
Paste wandelt Currency vor Cacheänderung in exakten Text oder verlustfrei begrenzte
bestehende Currency/Double-Speicherung um; andere Zieltypen werden verweigert.
Keine physische CURRENCY_TICKS-Unterstützung oder externe Access-MIME-Parität.
Vorhandene GUI/Feld-Prüfungen gemeinsam2/2 (6,42s), App gebaut: tatsächliches
Query→anders geordnetes Tabellenziel, SQL-Default, volle Integer/Binär-/NULL-
Unterscheidung, beide Currency-64-Bit-Grenzen nach Text, kleine Currency nach REAL,
Undo/Redo/Save und vollständige Verweigerung großer Currency→REAL ohne Entwurf/
SQL-Änderung. Fehlendes QMimeData-Include beim Build gezielt korrigiert. Keine neue
CTest-ID/Windows-Beobachtung. TYPED_GRID_CLIPBOARD.md; F340/F168/F173/F014 und
456/384/72 unverändert. Direkt weiterimplementieren.


Cut over pending native drafts (2026-11-11)
------------------------------------------
Native table Cut uses the existing bounded pending-range snapshot instead of
requiring an entirely clean cache. Earlier updates, inserted records/default
state and unrelated deletion markers remain pending. The target rectangle must
still contain writable, non-deleted records and no generated/readonly columns.
Typed MIME, visible field mapping and binary/NULL distinctions are unchanged.
Undo restores the exact pre-cut drafts; Redo reapplies the clear on that state.
Normal transactional Save remains the only persistence step. Cut participates in
the bounded native range history described above.

Clipboard publication can synchronously call application code. Cut checks the
same model, selection, visible mapping, row positions, field identities and raw
value types/values after publication before clearing. A changed source is refused
without clearing its new values. The clipboard may contain the successful copy;
no SQL write or forced draft save occurs. Existing range/source/memory/error
bounds remain; no Windows/Access event-equivalence claim is added.

Bounded multi-action history (2026-11-11)
---------------------------------------
Up to 64 native range actions are retained within a conservative combined payload
estimate of 256 MiB. Oldest actions are evicted first; a single larger snapshot
is refused before mutation. Stack transfers are prepared before cache changes.
A successful new action discards the redo branch; a refused conversion preserves
it. Save failure preserves both stacks; successful Save and explicit source reload
retire them. Fetch/topology boundaries cannot replay an older incompatible cache.
Ordinary cell/row edits still invalidate history. This does not add general cell
edit undo, committed SQL reversal or full Access event/history equivalence.

Existing mainwindow_documents passed in 15.28 s; the affected field target passed
in 0.58 s after correcting a regression expectation to the preceding pending
value. Coverage includes three actions backwards/forwards, append removal/default
preservation, branch replacement, failed conversion/Save, successful Save and
delete/Cut interleaving. No new CTest ID or Windows oracle. F340/F168 remain open.
