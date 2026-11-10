# Native DAO restartability and origin Property objects

Recordset.Restartable returns an actual Boolean. Table-type cursors returnFalse
as documented by DAO; supported query/dynaset/snapshot types2/4 returnTrue.
Frozen form adapters returnFalse because they cannot independently Requery.
Connection/table-identity and closed guards remain. A rollback-invalidated query
cursor can still inspect restartability before Requery, instead of being blocked
by its stale-row revision. True is a provider/mode capability, not a guarantee
that a changed dependency or a temporary source conflict cannot make Requery fail.
Table-type native Requery retains its pre-existing bounded refresh extension;
this change does not remove it or claim Windows table-type equivalence.
Writes383, getter arity450 and closed3420 are explicit.

Existing native Field.Properties now exposes SourceField and SourceTable as actual
read-only Property objects (Type=dbText10, canonical Name, actual String Value and
scalar default). It delegates to the already proven field-origin getter. Native
TableDef fields and supported physical dynaset fields retain their existing reader,
source/generation and property-write guards. Query-field Properties, full property
Counts/enumeration/ordinal catalog, Index/Relation origins and unproven provider
contexts remain separately unsupported; no guessed origin or blank value.

App/DAO built; one existing dao_execution bundle1/1 passed3.37s. Actual table/query/
frozen capability flags, stale rollback followed by successful requery, getter/
write/close guards, real Property values/type/read-only behavior and typed VBA
Property/Restartable usage are covered. No new CTest/catalog IDs;456/386/70.
F054/F055/F121 remain open; no new Original/GUI/package evidence here.

Public DAO contract explicitly states that table-type Restartable isFalse:
https://learn.microsoft.com/en-us/office/client-developer/access/desktop-database-reference/recordset-restartable-property-dao

Additional contract inspection (not Windows runtime-oracle evidence): published
Microsoft.Office.Interop.Access.Dao15.0.4420.1018 Recordset/Recordset2 metadata
marks GetRows.NumRows optional but has no default constant. Thus omitted NumRows
remains3251; no default was inferred from ADO or from optionality. AddNew has no
parameters. Public DLL was inspected only in the workspace, not executed on user
databases or committed. No native/Windows GetRows default proof is claimed.
