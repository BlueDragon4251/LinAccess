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


## Current existing record with other drafts

The navigation/find save path now commits a current existing record independently
of other update/insert/delete drafts. It requires unchanged non-NULL primary keys.
The original key must still exist after UPDATE inside the transaction; trigger
rekey/deletion rolls back. Qt selectRow refreshes only that record after commit,
without a full model reset or reconstructing unrelated drafts. SQL failure keeps
all drafts. Once committed the row is retired even if refresh reports failure,
preventing duplicate writes on retry. This SQL boundary invalidates block undo.

Sole populated inserted records keep the existing full-save/key-capture path.
Selective save of inserted/deleted current rows, complete trigger/concurrency/
Access event equivalence and automatic addition rows remain open. Existing
GUI/field targets2/2 (6.21s); no new target or Access2003 observation.


„Neuer Datensatz“/acNewRec spart jetzt das pauschale Speichern fremder Entwürfe:
sauberer aktueller Datensatz darf bei weiteren Entwürfen eine neue Zeile öffnen;
bestehender geänderter aktueller Datensatz nutzt den gezielten Save-Pfad. Ein
alleiniger gefüllter Insert wird regulär gespeichert, bevor die nächste neue Zeile
entsteht. Leere aktuelle Inserts oder selektive Inserts mit weiteren Entwürfen
bleiben verweigert. Bestehende GUI-Prüfung1/1 (6,42s), App gebaut: tatsächliche
RunCommand/VBA-Aufrufe, SQL-Abgrenzung, anderer Entwurf, neuer Zeilenfokus und Escape
ohne fremden Datenverlust. Keine neue CTest-ID/Windows-Beobachtung. F368/F173/F168
und456/384/72 unverändert. Direkt weiter.

Native PageDown now fetches enough bounded source rows for the actual viewport
before calculating its destination; [scope](TABLE_PAGE_FETCH.md). Other drafts
remain pending through the shared navigation handler.
