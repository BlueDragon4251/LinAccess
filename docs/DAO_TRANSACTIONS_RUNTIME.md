# Native project-scoped DAO transactions

DBEngine, Application.DBEngine, Workspaces(0)/Item/Count/enumeration and the
workspace's Databases(0) resolve native objects for the current SQLite project
connection. This is one default workspace/database per project connection, not
a complete process-global Jet engine or a multiple-database workspace. Typed
DAO.Workspace/Workspaces/DBEngine locals and Workspace parameters are accepted.
Default workspace/database references retain identity while held. CurrentDb
continues to create separate Database wrappers on the same actual connection.

BeginTrans opens a real SQLite savepoint, up to five nested levels (3003 beyond
that). CommitTrans releases the current level. Committed inner changes remain
subject to an outer Rollback. A single-level Rollback restores all changes since
that BeginTrans. A completion without BeginTrans raises 3034. Transaction state
survives temporary VBA object references and covers actual Database/QueryDef/
Recordset SQL on that connection. No per-statement approval dialog is added.

CurrentDb.Execute statements now dispatch through the real Database object;
SQL plus positional dbFailOnError no longer falls into a one-expression shortcut.
DoCmd.RunSQL keeps its separate host command path.

CommitTrans supports omitted/zero flags. Null flags raise 94, invalid flags 5;
dbForceOSFlush is recognized but rejected with 3251 because its distinct durable
flush contract is not established. Nested Rollback is rejected with 3251 before
changing state pending an Access 2003 reference trace. Clone inside an active
transaction and BeginTrans with an open clone are also rejected with 3251,
rather than pretending to implement Jet's implicit rollback on clone close.
Workspace.Close/custom workspaces/OpenDatabase/multiple connections, ODBC/ADO,
workspace security, forced flush and complete Jet transaction error/locking
semantics remain unsupported. Database/Recordset Transactions reports support
for this native connection scope, not all Access backends.

## Rollback and lifetime

Rollback invalidates materialized recordset values/bookmarks with 3420 until
Requery succeeds. This conservative native rule prevents returning rolled-back
writes from a cached row vector; it is not a claim of exact Jet cursor behavior.
Requery rebuilds rows and invalidates prior bookmark families. Frozen form
snapshots need a compatible form refresh/reacquisition; a failed refresh must
leave the old snapshot invalid. Close remains available for invalidated cursors.

A private SQLite generation callback invalidates wrappers when their connection
closes. Reopening the Qt connection creates a new context; old workspace/engine/
database/recordset references do not silently attach to it. Closing an actual
connection with an unresolved transaction lets SQLite roll back its writes.
External raw transaction termination and broader recovery remain limited.

## Evidence

Tests use a temporary project file and two independent Qt connections. They
check isolation before commit, committed visibility, actual rollback of
Recordset.Update, nested commit/outer rollback, unsupported-option preflight,
five-level limits, stale fields/bookmarks, successful/failed snapshot refresh,
clone restrictions, connection-close rollback and reopen generation. A complete
VBA transfer and a typed Workspace helper execute against the real database.

A real Qt form button executes the unchanged synthetic test handler through its
normal binding: two account updates commit, explicit rollback restores both,
and a handled duplicate-key error (3022) rolls both updates back. Me.Requery
refreshes the actual bound control; global/default/Application collection paths
and object identity are exercised. This is runtime/SQL/UI evidence, not an
original business workflow. Original modules have no DBEngine/transaction tokens;
no additional original-handler execution is claimed.

Contract sources (current documentation, not an Access 2003 oracle):
[Workspace.BeginTrans](https://learn.microsoft.com/en-us/office/client-developer/access/desktop-database-reference/workspace-begintrans-method-dao),
[CommitTrans](https://learn.microsoft.com/en-us/office/client-developer/access/desktop-database-reference/workspace-committrans-method-dao),
[Rollback](https://learn.microsoft.com/en-us/office/client-developer/access/desktop-database-reference/workspace-rollback-method-dao).
