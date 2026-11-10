# Actual native DAO Field.SourceField and SourceTable

Recordset, QueryDef and TableDef fields expose read-only native String origin
properties. Query metadata comes from actual SQLite prepared-statement physical
column lineage, restricted to main-schema table/field pairs. Aliases, direct views,
joins and derived Filter/Sort snapshots preserve those proven physical origins.
Result origins are retained beside rows/types, including clone sharing and Requery
restoration. Metadata retrieval does not step an additional query or execute user
functions. QueryDef.Fields uses the same schema-only preparation with unbound
parameters; held metadata invalidates on source/schema/rollback/generation changes.
Origin changes also participate in QueryDef.Fields column identity checks.

Actual TableDef.Fields supplies its held physical table context through existing
schema/namespace guards. Index/Relation field adapters receive no such table
context and retain unsupported semantics. Recordset table-type SourceField is
available; SourceTable refuses3251 because Microsoft's DAO contract explicitly
warns that it does not return a meaningful table name for that cursor kind.
Frozen form adapters, computed/unknown origins, TEMP/attached origins and compounds
including UNION refuse3251. No origin is invented from an output label or the
first UNION arm. Access default-empty/error identity for unsupported cases and
all provider/external-concurrency/ancestor-lineage semantics remain open.

Wrong getter arity450; writes383; existing closed/stale3420 guards remain.
Metadata remains readable at EOF. Normal row values, SQL and original sources are
unchanged by property reads. Fields.Properties full catalogs remain separately
bounded. This does not resolve F107's original unresolved source-column mismatch.

App/DAO/query-model/Currency targets built. One existing bundle5/5 passed3.90s:
DAO3.41, query model0.18, three Currency cases0.09/0.06/0.16. Actual view aliases,
two-table joins, parameterized QueryDef schema without execution, TableDef fields,
source replacement invalidation, snapshots, EOF/close, read-only errors,
computed/UNION/TEMP refusal and actual typed VBA origin reads are covered.
No new CTest/catalog IDs; F054/F055/F121/F122 remain open;456/386/70 unchanged.

Public contracts, not Windows execution-oracle evidence:
https://learn.microsoft.com/en-us/office/client-developer/access/desktop-database-reference/field-sourcefield-property-dao
https://learn.microsoft.com/en-us/office/client-developer/access/desktop-database-reference/field-sourcetable-property-dao

Preceding published source4ebc3f6c3e381c653647faad0df0e9dedfaaa21b has a fresh
checked DEB at /workspace/work/calendar-names/package/linaccess-0.7.0-Linux.deb.
Package identity/dependencies/content passed; rebuilt original integration2/2 in
167.38s (inventory47.31, original120.07). Already-passed GUI was not repeated.
All seven protected upload/fixture hashes unchanged. Inventory100/78/22; original
missing-macro boundary retained. This package predates field-origin additions.
