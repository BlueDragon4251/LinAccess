# Native exact Decimal searches and physical form filters

The shared native FindRecord value path carries Decimal values as exact Decimal text
in its checked payload and compares exact keys without converting to Double.
Tables, query-result views and forms accept a Decimal search value for proven
DECIMAL_EXACT fields. Whole-field typed searches retain the existing bounds,
NULL skipping and navigation rules. Binary fields remain excluded. Formatted
searches decode proven Decimal storage instead of treating its key as a BLOB.

Direct-table form filters now pass through the existing typed query-filter
compiler. This preserves the physical rowid alongside exact field values and
applies proven Date/Boolean/Single/Currency/Decimal expression semantics rather
than applying an untyped SQLite predicate to encoded values. Existing bounded
DataEntry restrictions and sorting are retained.

General Access search/locale/coercion/wildcard equivalence, arbitrary writable
query forms and the remaining predicate grammar stay open. No new catalog IDs
or complete requirement closures are introduced. Fixed456/386/70 unchanged.
Existing model scope passed (0.47s); after fixing the actual form projection
boundary, the affected existing GUI scope passed (18.18s), app/helper built.
The typed filter compiler requires explicit projections; physical forms now
list their fields alongside row identity, after the existing main/TEMP proof.
Text editors/search input accept redundant zero fractional digits; Decimal
comparison uses numeric keys and does not invent a scale-text equality contract.
The test navigates through RecordsetClone/Bookmark, respecting the existing
read-only direct form Recordset interface. No new CTest or Windows oracle.
