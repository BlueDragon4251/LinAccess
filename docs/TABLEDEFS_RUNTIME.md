# Bounded native DAO TableDefs

CurrentDb.TableDefs now returns a native collection of actual main-project user
tables. Count, Item(name/index), bang/default lookup and For Each are available.
Named lookup is case-insensitive. Internal `_la_`/`sqlite_` tables and views are
excluded. Enumeration follows SQLite catalog creation order; Access ordering is
not claimed. The held collection observes current schema changes on each access.

A TableDef exposes its real Name and OpenRecordset with the supported native
dynaset (2, current default) or snapshot (4) mode. These are actual recordsets over
the table, with existing typed field/edit/update behavior. Other modes and members
fail explicitly. Collection and item identity are retained while their objects
remain held. An observed missing table invalidates its old TableDef with 3420;
a later lookup after recreation returns a fresh object. Closed connections also
invalidate held collection/item access. Null indices raise 94, missing names or
indices 3265 and invalid arity 450.

This is not a complete DAO catalog. TableDef Fields/Indexes/Relations, attributes,
linked-table Connect/SourceTableName, Append/Delete/Refresh and Access ordering
remain open. A drop and recreation between two accesses can evade name-based
invalidation if the missing state is never observed; full schema generation
identity is not implemented. All project metadata tables remain internal.

## Original execution and regressions

The hash-verified original module
`4be67ece759c7816795451ec075ca3fe4fd87cd5aab632fcaca6ecb534fea5ba`
contains the real standard-module table-existence function. The native project
resolver now executes that unchanged function for each of the 29 original MDB
tables, then an absent name. Its For Each iterates actual native TableDef objects;
all present results are Boolean True and the missing result is Boolean False.
Each table's snapshot RecordCount is checked against independent SQL COUNT(*).
Only counts and result categories are logged; private names/rows are not printed.

The initially red original test exposed a second parser defect: its single-line
`If ... Then result=True: Exit For` branch was parsed in a temporary wrapper which
lost the surrounding loop. Exit For validation now uses the enclosing complete
procedure after inline branches are inserted. A nested numeric loop checks the
exact remaining execution count, and an invalid out-of-loop inline exit fails
before preceding effects execute.

Function result variables previously started as Null regardless of declaration.
They now use the same bounded basic defaults as procedure locals: Boolean False,
empty String, numeric zero, Empty Variant or Nothing Object. Unsupported declared
result types are rejected before execution. Full declared-result assignment
coercion and nominal type enforcement remain open; these tests prove defaults,
not a complete VBA type system.

`dao_execution` additionally covers collection/item identity, named and ordinal
lookup, hidden metadata/views, dynamic create/drop counts, observed invalidation,
recreation, closed connections, real editing through TableDef.OpenRecordset and
independent persisted SQL checks. All original MDB/Bridge files stay unchanged.

The current syntax inventory accepts 182/187 procedures. This remains separate
from actual procedure execution coverage and does not mean that the Kassa
application runs end to end or matches a Windows Access oracle.
