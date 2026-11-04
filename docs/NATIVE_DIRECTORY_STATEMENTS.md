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
