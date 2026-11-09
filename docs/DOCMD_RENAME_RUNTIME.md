# Native DoCmd.Rename for owned tables and queries

The form/report object model accepts DoCmd.Rename NewName, ObjectType, OldName,
including reordered named arguments. This bounded path requires explicit
ObjectType and OldName. acTable and acQuery delegate to the actual native DAO
TableDef/QueryDef Name setters through the owning database window. Scalar
arguments are evaluated once; Null reports94, unavailable implicit/type scope3251.
Missing optional arguments remain Missing (also with named arguments and comma
holes), so an omitted ObjectType cannot silently become acTable through Empty.
The caller is guarded across scalar evaluation and native dispatch.

The dispatcher refuses closed projects, modal/busy query contexts, an open query
datasheet or a matching open editable table model. Existing source, read-only,
namespace, transaction, schema, value and metadata guards remain in force.
Successful renames refresh the explorer and canonical selection. Persisted
printing page settings move with the underlying object name. A committed schema
change with failed connection restoration is reported as committed and requiring
reopen, rather than pretending rollback or retrying the mutation.

Other saved SQL, form/report sources and VBA references are not rewritten. This
is not Name AutoCorrect, implicit current-object resolution, form/report/module
renaming, full modal scheduling or an Access2003 event/error oracle. Original
files and imported archives remain protected. F052/F055/F080/F121 remain open;
fixed catalog456/384/72 is unchanged. Existing window integration tests execute
real VBA for named query rename, table rename, Missing/Null/unsupported errors,
real resulting SQL access and persisted page settings; no new test target.
