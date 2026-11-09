# Native DAO QueryDefs catalog and owned creation/deletion

Database.QueryDefs now exposes an actual native DAO.QueryDefs collection with
Count, Item/default item, Refresh and For Each. The catalog reads current main
owned definitions and imported query artifacts. Historical imported names collapse
to the latest canonical entry; owned/imported or owned case-fold collisions fail
explicitly. Native index order is case-folded name order, zero-based; Windows DAO
collection ordering is not claimed. Item reuses live weak-cached objects within a
held collection and obtains real canonical names/origins from the source reader.
Unsupported SQL in an imported item can still prevent instantiation/enumeration;
this is not a complete opaque Jet catalog or parser.

CreateQueryDef(Name, optional SQLText) creates a real owned named definition or,
for an empty Name, a temporary instance excluded from the catalog. SQL can initially
be empty, then assigned through the persistent owned SQL path; executing unset SQL
is refused rather than returning invented rows. Duplicate names, table/permanent
view collisions and invalid names fail. Only a stale, proven installer-managed
temporary view can coexist with a newly recreated catalog name; the normal view
revision mechanism replaces that view before data execution.

Named creation compiles before mutation and inserts through a unique savepoint,
creating the owned metadata table if necessary. It never commits an outer user
transaction. QueryDefs.Delete removes only owned query metadata through its own
savepoint. Imported deletion requires an archive-preserving overlay and is explicitly
unsupported; archived source bytes are unchanged. Native create/delete advance a
workspace namespace generation, invalidating held old objects after deletion or
same-name recreation, even when metadata rollback restores a deleted row. Fresh
handles must be acquired after such rollback. External SQL/manual metadata edits
and cross-connection same-name recreation remain outside this generation proof.

Workspace BeginTrans snapshots the five query-view registry properties; outer
Rollback restores them along with SQLite's transactional temporary-view schema.
This prevents a restored view becoming unmanaged and retaining obsolete SQL after
a later definition edit. Commit retains the current registry; the existing nested
rollback guard remains. These caches contain names/revision metadata, no runtime
host or cursor references.

Existing actual VBA/SQL cases cover typed DAO.QueryDefs, real named and temporary
creation, named arguments, live counts/default/name/index items, For Each, retained
item identity, Refresh, duplicate3012, optional SQL assignment and actual execution.
Creation rollback removes catalog rows. Delete/recreate invalidates old handles3420.
A delete followed by dependent recordset execution inside a transaction and rollback
restores registry ownership; subsequently edited SQL returns the new value through
another query's view dependency. Independent SQL verifies catalog rows, retained
metadata and unchanged source row count. No new CTest target or fixed-catalog item.

Append/attributes, full imported overlays, arbitrary namespace/history/order/
error/concurrency identity and Windows Access2003 execution-oracle parity remain
open. DAO.QueryDefs joins the explicit closed native type registry; this adds no
COM provider and does not infer an interface from a reported TypeName.

Owned persistent QueryDef.Name now supports an atomic bounded rename. It changes
only the catalog name, preserving the entire definition byte-for-byte, SQL,
parameter bindings and the caller's held field/parameter objects. Case-only
renames retain the requested spelling. Collisions, invalid names, read-only
connections and active outer transactions refuse before mutation. The latter
restriction avoids claiming rollback identity for moving namespaces before that
identity has been implemented. Temporary/imported queries remain protected.
The existing installer rebuilds managed views within the same savepoint; failure
rolls back both SQL state and its five registry properties. Old namespace handles
are invalidated, the initiating handle follows its new name, and the held catalog
reuses a renamed live item while rejecting stale cache entries. New queries can
reuse the vacated name without returning the renamed query accidentally.
Other saved SQL, forms, reports and VBA bindings are not automatically refactored.
This is not Name AutoCorrect, cross-connection generation proof or a complete
Access2003 QueryDef rename/transaction oracle.

The database window now offers Ansicht → Abfrage umbenennen for owned queries.
Open query datasheets must be closed first; imported names are disabled/protected.
The selected live QueryDef is held across the name dialog so deletion/recreation
cannot silently redirect the action to a different catalog object. Explorer
selection follows the actual new name. Other stored bindings remain unchanged.
Table and query rename share a checked core-only datasheet-page catalog move:
existing settings follow the case-folded identity without changing their bytes,
unknown catalog shape or a target settings collision aborts the entire rename.
No QtGui dependency was added to the DAO/data library.
