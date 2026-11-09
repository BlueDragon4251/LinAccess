# Native table/query object deletion

The database window now offers Ansicht → Ausgewähltes Objekt löschen for owned
native tables and queries. Its actual confirmation defaults to No. A held DAO
object is checked after that dialog, so deletion/recreation or a rename cannot
redirect an earlier confirmation to a different object. Imported source objects
are disabled in this bounded window command. Open matching datasheets/editable
table models must be closed first; original archives/files are unchanged.

DoCmd.DeleteObject supports explicit acTable/acQuery and ObjectName, including
reordered named arguments. It calls the same actual window/DAO deletion path.
VBA calls do not receive an additional SQL consent prompt. Missing optional type
or name stays Missing and is explicitly unsupported3251; Null reports94.
Implicit selected-object resolution, form/report/module deletion and full Access
confirmation/error/event behavior remain unproven. Form callers are guarded
across one-time scalar argument evaluation and dispatch. Busy/modal/closed
project contexts refuse rather than mutating detached or unavailable objects.

QueryDefs.Delete now removes its managed executable view through the existing
installer in the same savepoint as the owned definition and datasheet page row.
A failed step rolls back SQL/TEMP views and the five installer registry properties;
namespace handles are invalidated only after successful release. Existing outer
Workspace rollback restores its saved registry and transactional views. Read-only
connections refuse before mutation. New objects with a vacated name receive fresh
views/page state instead of inheriting the deleted object's state.

Native table DROP/TableDefs.Delete now removes the matching table page row within
its existing schema savepoint, alongside the existing source/index/counter cleanup.
Malformed page catalog metadata rejects the complete mutation. Raw SQL rename
with stored page settings refuses until it can rebind them; explicit native
TableDef.Name uses the existing checked rename path. Dependent relations retain
the existing refusal/foreign-key protection. Other saved bindings are not silently
rewritten or deleted. No QtGui dependency is added to the core DAO library.

Existing DAO/window targets exercise real VBA, actual table/query/catalog/view
removal, open-object rejection, confirmation No/Yes, unchanged unrelated data,
page cleanup, name reuse and rollback after a deliberately invalid page schema.
No generated business data, extra CTest target or fixed-catalog ID is introduced.
Full F052/F055/F080/F121 and Windows Access2003 parity remain open;456/384/72 fixed.
