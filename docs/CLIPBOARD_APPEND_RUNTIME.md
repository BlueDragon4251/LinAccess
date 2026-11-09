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

The initial target is an existing editable cell. An empty table or append-only
clipboard command is not yet covered. Pasting over unrelated drafts, typed binary
MIME, arbitrary regional input masks and full Access range undo remain open under
F340. This does not close F340 or establish Windows clipboard equivalence.
