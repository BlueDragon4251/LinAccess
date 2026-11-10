# Native DAO Recordset.Name

Opened physical table cursors retain the requested table identity. Database
OpenRecordset on saved queries retains the requested query name; actual QueryDef
OpenRecordset retains its nonempty Name. Raw SQL and unnamed QueryDefs retain the
first256 UTF16 code units of the original Access SQL, before compilation and
independently of the physical writable target. Clones preserve this opening
identity. Later QueryDef changes do not rewrite an already opened cursor's name.

Name is read-only383, takes no arguments450 and respects closed/stale connection
and transaction guards. Frozen host adapters and Filter/Sort-derived snapshots
without a proven opening-name contract explicitly refuse3251. This does not add
provider names, a Recordsets catalog or prove every Windows error-order case.

Public API reference (contract, not an Access2003 runtime oracle):
https://learn.microsoft.com/en-us/office/client-developer/access/desktop-database-reference/recordset-name-property-dao

App and existing DAO target built. One existing dao_execution bundle passed1/1
in3.41s: actual table and named-query openings, QueryDef and clone names,
long original SQL truncation, writable SQL identity, read-only/argument/closed
boundaries. No new CTest/catalog ID;456/386/70 unchanged, F054/F055 remain bounded.
Latest private published predecessor8de1ead has matching verified tree; latest
checked package/Original is still sourcec67cbd9e before query-writing/name changes.
