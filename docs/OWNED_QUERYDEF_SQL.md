# Persisted SQL for owned native QueryDefs

CurrentDb.QueryDefs(name) now retains the actual resolved main-catalog query name
and its owned/imported origin. Name is readable; renaming remains unsupported.
An owned QueryDef SQL assignment compiles the replacement first, preserves other
JSON definition members and conditionally updates the actual main._la_objects row
against the previously read definition. A conflicting update fails3197; it never
commits the caller's transaction. Only after a successful write are local SQL/plan,
parameter bindings and parameter-object generation replaced.

Owned QueryDef members refresh from the current saved SQL. Parameters held across
an SQL change fail3420; refreshed Parameters collections expose the current plan.
QueryDefs and their parameters also use the existing workspace/connection/thread
liveness guard. Non-table DAO recordsets refresh the common query-view revision
before loading; owned action execution does likewise, preventing stale dependency
views after an SQL change. No cursor/host/document pointer is added to plan caches.

The existing imported-query behavior is preserved: imported named QueryDefs retain
an instance SQL working copy, and temporary CreateQueryDef remains temporary.
Imported SQL assignment does not overwrite archived source bytes or claim native
original-format persistence. A full imported overlay/write adapter, QueryDefs
collection/catalog editing, rename/append/delete, immutable identity across a
same-name drop/recreate and complete Access concurrency/error/oracle parity remain
open. This change does not claim those features are implemented.

An existing actual VBA/SQL DAO case reads the canonical Name, binds a real parameter,
executes the query, changes/persists SQL, observes the new SQL from another instance
and rejects the held old Parameter3420. Another instance changes the parameter
schema; the first refreshes. Updating a saved base-query SQL changes actual rows
returned from the dependent query. A real workspace transaction rolls back a SQL
change, and a subsequent getter sees the restored definition. Close invalidates
SQL access3420. Independent SQL verifies retained metadata and original row count.
No new CTest target or fixed-catalog item.
