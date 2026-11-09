# Native Set RHS before dynamic target resolution

MS-VBAL5.4.3.9 evaluates the right expression as an object-reference value before
resolving an unbound Set destination. Native member and indexed Set now evaluate
and retain that object once, then refresh actual aliases and resolve receiver,
setter metadata and index expressions. A RHS that changes a receiver or resizes
an array therefore affects the actual subsequent target. Owning property setters,
class default setters, native array elements and public form setters use the same
cached RHS. Existing Let order and ordinary unindexed variable Set are unchanged.
The explicit unbound-ParamArray mutation guard still precedes an unsupported call.

A failing RHS performs no receiver/index/setter call. A receiver changed to Nothing
by a successful RHS fails91 before an index/setter call. New objects are not
constructed again merely to perform assignment. This does not complete static
VBA l-expression classification, declared-target coercion before dynamic lookup,
all With/automatic-variable/reentrancy/event behavior or a Windows Access oracle;
those surrounding F040/F043/F045/F047/F104 boundaries remain open.

Shared technical core sources check actual member receiver changes (R,T,I),
default receiver changes (R,I), a RHS resizing/reinitializing a real object array
before selecting its changed index, owning named setters, RHS failure7012 with
no LHS calls and a RHS clearing its target with native91. Actual project class
and real form cases prove the native providers. Fixture references and its owned
class row are explicitly cleaned up before later catalog-shadow tests.

App built. First existing core/GUI/DAO/Original bundle4/4 passed145.89s:
5.53/19.19/2.75/118.41s. No targeted retry or new CTest ID; seven original/upload
hashes remain unchanged. Original retains its known missing-macro limit and is
not a positive initialized-sale proof. No Windows/full-suite/package refresh.
Catalog456/386/70 unchanged. Prior named-setter d618eaf was privately published
and its remote tree verified. Latest checked DEB9affc5a at
/workspace/work/error-text-functions/package/linaccess-0.7.0-Linux.deb precedes
ParamArray aliases, named setters and this Set ordering change.
