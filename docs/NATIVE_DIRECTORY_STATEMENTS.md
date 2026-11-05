# Native MkDir and RmDir statements

The native interpreter compiles unqualified `MkDir pathExpression` and
`RmDir pathExpression` statements and executes them only in active branches.
Expressions use the normal scalar/default-member evaluation and VBA error path.
Actual UTF-8 Linux paths are supported, including relative paths. MkDir creates
one directory, with normal operating-system permissions/umask; no missing-parent
or recursive creation is invented. RmDir uses the native Linux rmdir operation,
which removes only empty directories and cannot unlink a normal file.

Explicit bounded errors: Null94, empty/NUL/wildcard path52, missing path76,
existing MkDir target75 and denied/nonempty/wrong-target operations75. Windows
backslash/drive paths and RmDir symbolic-link targets are explicitly unsupported
3251. Other platform/Windows path, permission, locale and exact error precedence
behavior has no Access2003 execution oracle. FileCopy/Kill/Open/Print/input and
complete VBA filesystem support remain open. No recursive deletion is added.

The existing core target operates on a unique temporary directory: Unicode name,
creation, repeated-creation error, preservation of a nonempty directory and a
normal file, removal after emptying, missing directory, Null and empty pathname.
No original fixture is used as a mutation target; no new CTest target is added.

Desktop48/48 (82.86 s), fresh GUI-free core1/1 (1.76 s), Debian package/build ID,
pinned SQLite and all four original fixture hashes passed. This is bounded
language-dispatch progress under the existing F040 requirement; that complete
parser/AST requirement remains open. Catalog stays456/384/72.

## Working directory

Unqualified `ChDir pathExpression` now changes the actual process working
 directory using the native filesystem operation. A failed change retains the
previous directory. `CurDir`, `CurDir()`, `CurDir$()` and qualified `VBA.CurDir`
return the current UTF-8 Linux directory; optional named `Drive` accepts the
empty string. Nonempty Windows drive selectors remain explicitly unsupported
3251. Null selectors94, excess arguments450; unavailable current directories
raise76/70/75. Bare module variables/procedures and host identifiers retain
priority over the unqualified intrinsic. Syntax validation does not read cwd.
ChDir empty path76; other invalid paths follow the bounded rules above. RmDir
refuses removal of the current directory, including equivalent path aliases75.
The working directory is process-wide, not a per-form property or Windows
per-drive directory emulation. SQL CurDir was not registered in the first segment.

Existing core cases cover a Unicode directory, actual changes, failed changes,
current-directory removal protection, named/qualified/dollar/bare reads, local
variable shadowing and restoration of the original working directory. Desktop
48/48 (83.85 s), fresh GUI-free1/1 (1.79 s), current Debian build-ID/SQLite3.54.1
and all four original fixture hashes passed. F040 remains open;456/384/72.


The SQL compiler now rewrites `CurDir(...)` to a native non-deterministic scalar
adapter which invokes the same VBA calculation. It is not marked deterministic
or innocuous and is not evaluated as a compile-time constant. A prepared actual
SQLite SELECT is reused after a native directory change; both results reflect
the current directory. DAO OpenRecordset returns the actual path as a String.
SQL dollar/qualified spelling and complete Windows SQL metadata/error semantics
remain open. Existing targeted core/DAO/query3/3 (2.33 s), updated Debian build-ID/
SQLite backend and four original hashes pass. Last full48/48 (83.85 s), fresh
GUI-free1/1 (1.79 s); runtime code unchanged in this SQL adapter extension.
