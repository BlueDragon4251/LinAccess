# Native project Err environment

MS-VBAL6.1.3.2 specifies one Err instance for a VBA environment, accessible in
any procedure. The runtime now obtains the actual ErrObject through its host,
instead of constructing a fresh object for every procedure. Generic hosts keep
one object for their lifetime; delegated class service hosts forward their real
provider. Native projects store the object in the existing connection-owned VBA
project context. Standard module hosts, actual form/report hosts, native classes,
SQL-VBA hosts and the legacy imported-procedure runner use this same provider.
No thread-global object shares errors between independent databases.

The five mutable error properties, retained references and default Number value
therefore agree across nested procedure calls and independent native module hosts.
OnError/Resume/explicit Exit clearing operates on that actual shared object.
Handler enable/active state remains local to each executing procedure. Raised
errors update the shared object before propagation/debugger handling, retaining
Number, Source, Description, HelpFile and HelpContext. A successful procedure
entry does not silently clear previous error properties.

Native error property access, assignment, default reads and Clear validate the
real connection/driver/database/thread and weak project-context owner. Explicit
project reset creates a new environment and invalidates retained old references;
closed connections reject access3420. External callers reusing a raw connection
must use the existing project-reset lifecycle; unmanaged reopen/complete external
connection-generation semantics remain outside this proof. Full error history,
Error/Error$, complete compile/runtime diagnostics and Windows Access behavior
remain open under F051; this is not complete VBA error compatibility.

Existing app/core/GUI/DAO/Original targets built. The first bundle exposed one
old core fixture assuming a new procedure silently reset Err to zero after an
earlier handled error. Its intended zero starting state now explicitly calls
Err.Clear. Only core was rebuilt/repeated, passing5.03s. GUI18.42s, DAO2.75s and
Original118.10s passed in the initial bundle. Shared technical sources verify
nested identity/properties, locally handled and propagated errors, actual class
service delegation, retained identity across two real module hosts and a real
form, isolated connections, reset invalidation and closed-connection writes.
Existing SQL-VBA cases ran in the GUI target; no dedicated SQL singleton identity
oracle is claimed. Seven original/upload hashes remain unchanged. Original retains
its known missing-macro limit. No new CTest ID/full suite/Windows oracle.

Checked DEB55096a7, before this environment change, is retained at
/workspace/work/error-clear-boundaries/package/linaccess-0.7.0-Linux.deb;
its payload/build-ID/rights/no-fixtures checks, MDBTools1.0.1 and SQLite3.54.1
nestedJOIN/NULL/counter checks passed. Catalog456/386/70 unchanged.
