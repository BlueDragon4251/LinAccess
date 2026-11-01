# Bounded native record clipboard

RunCommand Copy (190), PasteAppend (38) and Access-95 Edit-menu commands 2/5 use
the real system clipboard. Ctrl+C on a selected record copies its complete row;
ordinary text editors retain native text-copy behavior. The versioned MIME data
preserves Unicode, multiline text, NULL versus empty text, exact 64-bit integers,
finite reals and binary bytes. A quoted TSV representation is also provided.

PasteAppend validates field shape/names/types/version and 8 MiB/255-column
limits. External plain text is not guessed to be a record. Actual AUTOINCREMENT
columns, verified through SQLite metadata, generate destination keys; manual keys
remain unchanged and conflicts report 3022. Imported counters without native
generation fail with 3251. SQL uses a savepoint and execution budget. Failure or
exclusion by the current form filter rolls back. Snapshot clones retain surviving
bookmarks and gain the appended row. DoCmd targets the active document; Me stays
with its caller.

Imported/dynamic AllowAdditions prevents append/new navigation and disables the
new-record selector entry. Pending edits use existing save validation. Targets
with insertion/update/Dirty event bindings explicitly fail 3251 before inserting;
the complete insertion pipeline is still open. Current runs for the saved row.
Complete focus/event ordering and all property transitions need an Access oracle.

Tests exercise actual clipboard, native Ctrl+C, SQL equality of all supported
types including a large exact integer, new keys, bookmarks, active-target scope,
invalid/oversized payloads, disabled additions, unsupported event/counter contexts,
filter rollback and manual-key conflicts. Clipboard integer preservation does not
claim that the bounded VBA numeric Variant supports every 64-bit value.

The unchanged original module SHA-256
`dbd3623b9404e490a814b70de6318d2c0877969b7feef081696d3e235ae3b427`
and its unchanged layout run against synthetic rows with the actual exported
14-column schema/key/indexes. The table has one manual primary key, no AutoNumber.
Direct append reports 3022. One original handler invocation and one bound button
click handle the native error message and retain exactly one row. Completed
handler status here proves its error path, not successful duplication. The button
is made visible/enabled for the probe. No original rows/source/credentials appear
in output. This is no independent Windows Access result comparison.

Full clipboard interoperability, multiple records, external TSV import, complete
insertion events, all conversions/defaults/constraints, writable query sources,
non-primary-key generators, Access clipboard formats and rendering remain open.
