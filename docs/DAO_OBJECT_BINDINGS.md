# Native DAO object identities and typed bindings

Native SQLite-backed DAO objects report their own stable class identity:
Database/Databases, Workspace/Workspaces, DBEngine, Recordset, Field/Fields,
QueryDef, Parameter/Parameters and TableDef/TableDefs. Singleton collections
carry an explicit native collection identity. The bounded form Recordset
adapter reports Recordset while retaining its existing guarded member subset.

Supported DAO-qualified scalar module fields initialize to Nothing and retain
actual references in host module state. Matching local/module/argument/result
assignments validate the native identity; mismatches refuse 13 before replacing
the binding. Existing ByRef declared-type checks and connection/recordset/clone/
transaction guards remain active. Required and optional supported DAO parameters
accept matching native references/Nothing; omitted and explicit Nothing defaults
use independently initialized slots. Native function results accept the same
qualified type names.

The existing core and DAO targets verify a real SQL cursor retained and moved
across procedure calls, actual TypeName metadata, a typed Database return with
preserved object identity, optional Nothing/matching Database arguments, and a
wrong Set leaving the original Recordset/cursor intact. No extra CTest target.

This is bounded native type binding, not a COM/reference-library loader or a
claim of complete DAO/Access compatibility. Unqualified DAO type resolution,
automatic As New, unsupported native/external classes, typed DAO arrays, complete
object/default-member Let semantics, interface/subclass compatibility and the
Access 2003 compiler/error precedence oracle remain open. Original MDBs remain
unchanged; native DAO writes target project/test copies.


## Native DAO identity and qualified object bindings (2026-11-04)

DAO_OBJECT_BINDINGS.md defines actual native DAO identities, qualified scalar
module/local/parameter/result binding and independently omitted Nothing
defaults. Actual SQL cursor state persists across procedure calls; wrong Set
leaves the original typed Recordset intact. Existing owner/connection/clone/
transaction boundaries and the bounded form Recordset adapter remain active.
Full desktop 48/48 (77.79 s), freshly rebuilt GUI-free core 1/1 (1.23 s),
targeted core/DAO 2/2 (1.59 s), Debian payload/permissions/Build-ID/pinned SQLite
and all four original fixture hashes pass. No extra CTest target. Scoped
checklist 371 checked / 72 open, not complete Access 2003 parity.
