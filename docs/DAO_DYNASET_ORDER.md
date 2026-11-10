# Native updatable dynaset row ordering

Non-invalidating reloads of native table-backed dbOpenDynaset cursors now merge
by exact physical row identity: retain surviving known rows in their existing
family order, then append previously unseen rows. This implements the documented
AddNew rule that a caller's new dynaset record is appended regardless of key order.
The merge also applies after later Edit/Update/Delete refreshes, so an ordinary
update cannot silently revert the family to SQLite physical ordering.

Ordering is established before existing bookmark/current/clone restoration.
LastModified and current-record preservation remain independent. Table-type
cursors retain the existing actual Index/physical ordering path. Explicit Requery
still invalidates and rebuilds the cursor instead of freezing the prior cache.
Query-backed editable joins/sorts, complete external-user incremental visibility,
all Delete/current-position semantics, locking, default Jet initial order and
full Windows dynaset equivalence remain open. This is a bounded native
materialized family, not an actual Jet provider.

The merge is linear in materialized row count and uses the existing100000-row
limit. No SQL is added, no values are converted through Double, no peer cursor
or pending data buffer is reconstructed from guessed key text.

App/DAO built; one existing dao_execution bundle1/1 passed3.24s. Actual table and
dynaset families insert physical rowid0: table order0/1/2/3 versus dynaset order
1/2/3/0. Held clones enumerate each actual family, preserve their bookmarks, and
verify order remains after Edit and another insertion. Existing current retention,
LastModified, retry and pending-navigation matrices continue passing. No new
CTest/catalog IDs;456/386/70 unchanged, F054/F055 remain open.
No fresh GUI/Original/DEB proof for this scope.

Public DAO contract, not Windows Access execution-oracle evidence:
https://learn.microsoft.com/en-us/office/client-developer/access/desktop-database-reference/recordset-addnew-method-dao
