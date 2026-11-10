# Native form physical source and inserted row identity

Direct-table forms now read and mutate the physical main source explicitly.
Before loading and after user BeforeUpdate/Delete callbacks, a checked main-table
existence/TEMP-shadow boundary prevents mixing cached main values with another
same-name temporary source. Save checks again after evaluating field values.
An unavailable source or TEMP shadow refuses instead of writing another table.
Dirty buffers survive a refused save and can be saved after removing the shadow.

INSERT with supplied columns and DEFAULT VALUES both use RETURNING rowid, as
UPDATE already does. The form consumes exactly one non-NULL returned identity,
including the valid integer0, before completing counter synchronization. It no
longer depends on Qt's lastInsertId omission at0. DataEntry membership, selection,
snapshots and subsequent edits therefore retain that physical row. Delete uses
main qualification after the source check; its existing rollback/events remain.

This does not add writable form query sources, provider locking, drop/recreate
identity generations, exact native deletion errors or arbitrary trigger identity
rewrites. Record clipboard insertion is a separate path and is not changed here.

App/current-record/layout/delete/Currency-insert/update targets built. An initial
fixture compile used private C++ save/delete methods; it was corrected to actual
public VBA RunCommand/Dirty routes before any checks executed. One existing
five-check bundle passed5/5 in14.42s. Actual keyboard insertion of key0 in an entry
session, current selection/count, subsequent update, independent physical SQL,
shadowed save/delete refusal, unchanged main/TEMP data and retry after shadow
drop are covered alongside existing layout/form events and Currency checks.
No new CTest/catalog IDs;456/386/70 unchanged. Latest checked DEB/Originalc67cbd9e
predates these changes; no new Original/package/oracle claim in this scope.
