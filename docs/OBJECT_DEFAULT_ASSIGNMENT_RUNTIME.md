# Native scalar default-member Let assignments

Value assignment to an explicitly declared native object variable uses its
actual scalar default setter instead of replacing the object reference. Both
plain value assignment and the existing explicit Let syntax share this path.
DAO Field and Parameter defaults route to the real Value setter; bound native
controls route to the existing preview Value setter/event/write guards. Set
continues to replace references through checked binding.

The runtime retains the actual target object before evaluating the scalar RHS.
Nothing refuses 91; classes without an advertised scalar default retain an
explicit 3251 boundary. DAO edit/readonly/type/transaction/closed-context checks
and native form/control update/readonly behavior remain active. This adds no
COM activation and does not write original MDB files.

Existing DAO cases exercise typed Field and generic Object references to the
same real Field, preserve their identity, observe the old SQL value until Update
and the actual new value afterward. Existing real QueryDef parameter-collection
execution assigns through Let parameter=3 and checks the SQL result. Existing
form-object regression assigns a bound Control via Let ctl=..., saves the real
form record and verifies both native widget text and SQL. No extra CTest target.

Variant-held object Let/coercion rules, unsupported object/default properties,
full error/evaluation precedence, classes/interfaces and an Access 2003 execution
oracle remain open. This does not claim general object/COM assignment parity.


## Native scalar object default-member value assignments (2026-11-04)

Declared object value/Let assignments invoke actual Field/Parameter/Control
Value setters and preserve their reference binding; Set continues reference
replacement. Real DAO edit/SQL and bound form/save paths execute in existing
cases, including explicit Let syntax (already parsed previously).
OBJECT_DEFAULT_ASSIGNMENT_RUNTIME.md records Nothing/no-default and remaining
Variant/coercion/oracle boundaries. Full desktop 48/48 (79.37 s), freshly rebuilt
GUI-free core 1/1 (1.49 s), targeted core/DAO/forms 3/3 (10.76 s), Debian payload/
permissions/Build-ID/pinned SQLite and four original fixture hashes pass. No new
CTest target. Scoped checklist 376 checked / 72 open, not complete Access parity.
