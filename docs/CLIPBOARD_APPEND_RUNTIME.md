# Native rectangular paste with additional records

A clean manual-submit native table datasheet accepts a rectangular text clipboard
range that extends past its last record. Existing unfetched records are fetched
before deciding which rows are new. Target columns and every input value are
checked before any draft is created. Generated/AutoWert and binary targets remain
protected. Date defaults use the existing native default resolver; other omitted
columns retain their normal insert/default behavior.

The extra rows are pending records, not immediate SQL inserts. They participate
in record undo, discard, the existing transactional Save path and generated-key
capture. A failed conversion leaves neither additional rows nor changed cells;
a failed database constraint rolls back the SQL transaction and preserves the
pending edits for correction or discard. Model structural signals stay active
when rows are created or removed.

Empty tables start at their first visible writable column. The explicit native
context-menu append command and acCmdPasteAppend (38) append after all fetched
records, starting at the selected writable column, or the first writable column
when no cell exists. Normal paste keeps its current-cell behavior. All of these
paths leave pending drafts for the existing Save command. Native typed form-record
MIME is refused explicitly rather than inserting its textual header as data.
Pasting over unrelated drafts, typed binary MIME, arbitrary regional input masks
and full Access range undo remain open under
F340. This does not close F340 or establish Windows clipboard equivalence.
