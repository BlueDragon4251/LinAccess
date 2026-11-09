# Native form Requery and displayed records

Read-only query forms that become nonempty select their first actual record after
Requery. A blank new-record cursor is retained only for an eligible writable
direct-table form. This prevents a populated read-only child from exposing Null
bound fields while its record selector remains at zero.

Requery compares actual displayed columns, types, values, key and index. Changed
snapshots refresh initialized ancestor calculated controls after the Current
event. Unchanged snapshots retain their cached values. The native execution lease
protects closure during events. Child Requery does not save a parent draft.

The existing cache regression covers empty/nonempty transitions, changed row
values, parent calculated references and independent SQL verification that the
parent draft remains unsaved. The original corpus observes real SendKeys source
values at their original debugger checkpoints; each keyboard case rolls back its
own SQL changes before the next independent handler.

Verification: existing layout and GUI targets 2/2 (14.90 s total), original Kassa
corpus 1/1 (118.10 s), final DEB payload/Build-ID/MDBTools 1.0.1/SQLite 3.54.1
checks, four fixture hashes and three uploaded-original hashes passed. No new
CTest ID, Windows run, full-suite run or physical-printer proof is claimed.

The original keyboard workflow now reaches explicit 3146 at source line 1023;
a separate genuine tagged-button probe progresses to missing macro error 2485.
Four macros exist in both original MDB catalogs, but the existing bridge archive
contains zero macro sources. This is an export gap, not a missing user database.
Full dependency scheduling, arbitrary query updates and full original sale
workflow remain open. Catalog remains 456 / 386 confirmed partial scopes / 70
open full requirements; no Access-parity claim follows from these counts.
