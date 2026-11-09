# Native table navigation with pending edits

A clean current native table record can navigate through the loaded/fetched
record range without saving unrelated pending changes. Keyboard arrows and
native GoToRecord/RunCommand navigation use the same path. Existing activation,
editor validation, readonly/query checks and one-million-record bounds apply.
The shared find-navigation guard also permits a clean current record without
committing unrelated drafts.

Before leaving a dirty current record, automatic Save is allowed only when it is
the sole pending record and has a primary key. Navigation additionally allows a
single inserted record with a supplied non-generated non-NULL value; the normal
transactional Save/key-capture path establishes its real position before moving.
A wholly empty pending inserted record is not silently committed. Pending new
rows are recognized as records for Escape/record undo even before an ordinary
field is edited.

The opt-in inserted-record path belongs to record navigation. Other callers of
the existing current-record guard retain their previous bounds. Multiple pending
records requiring saves, a deleted current record, arbitrary writable query
sources, Access's automatic addition row and full event/error identity remain
open under F368/F168/F080. No new Windows Access navigation oracle is claimed.
