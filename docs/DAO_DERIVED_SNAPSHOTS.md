# Native DAO Filter, Sort and derived snapshots

Query-backed dynaset/snapshot cursors expose writable String Filter and Sort
properties. They affect a subsequent explicit OpenRecordset(dbOpenSnapshot),
not the existing rows, position, bookmark or source definition. Real checked Jet
predicate and output-field sort compilers construct the new SQL query. Sort
supersedes source ordering; filtering composes with the existing query, including
previously derived predicates. Existing bound parameters remain bound to their
original values. Source filter changes or closing the source do not change an
already opened result. Requery executes the derived plan. Clones copy property
strings independently while retaining their existing shared row/bookmark family.

Only explicitly requested snapshots (4) are implemented here. Omitted Type,
editable derived cursors, table/frozen adapters and additional Options remain
unsupported. This prevents inventing writable identities or default cursor
semantics. Null property assignment94, wrong arity450 and existing closed/stale/
transaction guards remain. Invalid SQL/order fails at derived opening without
moving or changing the original cursor. Unknown sort fields refuse3251.
CopyQueryDef on transformed results remains unavailable: an original definition
is not silently misreported as the transformed SQL. Exact Windows property
values, default/mode/error/clone identity and complete DAO parity remain open.
Existing Currency filtering limitations are inherited explicitly.

Public API contracts, not Access2003 runtime-oracle evidence:
- https://learn.microsoft.com/en-us/office/client-developer/access/desktop-database-reference/recordset-filter-property-dao
- https://learn.microsoft.com/en-us/office/client-developer/access/desktop-database-reference/recordset-sort-property-dao
- https://learn.microsoft.com/en-us/office/client-developer/access/desktop-database-reference/recordset-openrecordset-method-dao

App built; one existing DAO bundle passed3.37s after correcting a field-name typo
in the newly extended check (currencyUnits, not the currency factory). No runtime
failure or repeated unrelated suite. Actual SQL rows/order, composed empty
filter, parameter retention, independent clone properties, error isolation,
Currency subtype/200000 units/dbCurrency5 and named typed native VBA call covered.
No new CTest/catalog IDs; F054/F055 remain open;456/386/70 unchanged.

The preceding published source ad42eadace27a518fbbaf1f7a9819facc1e383c1 has a
fresh checked DEB at /workspace/work/dao-copy-querydef/package/linaccess-0.7.0-Linux.deb.
Package identity/content/dependencies passed. Rebuilt existing integration3/3:
mainwindow19.85s, query inventory47.59s, original120.91s; total188.36s. Original
missing-macro boundary retained. All seven original/fixture hashes unchanged.
Inventory100 classified,78 executed,22 explicit boundaries; no Windows result
oracle. This DEB predates the derived snapshot change documented above.
