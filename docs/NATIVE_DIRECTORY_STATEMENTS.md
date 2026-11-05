# Native directory and filesystem statements

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
behavior has no Access2003 execution oracle. Open/Print/input and
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

## Native FileCopy

The unqualified `FileCopy sourceExpression, destinationExpression` statement now
executes actual C++ filesystem copying in active branches, with normal scalar
evaluation and String coercion. Existing destination files are replaced. Binary
contents including embedded NUL bytes and UTF-8 Linux pathnames are preserved.
Validation compiles expressions without doing filesystem I/O.

The bounded implementation requires regular source/destination files. Final-path
symlinks, identical files/hardlink identities and Windows drive/UNC mapping are
explicitly refused3251. It does not claim race-free pathname resolution, Windows
open-file sharing/readonly flags, complete error precedence, timestamps/ACL
equivalence or atomic replacement on an I/O failure. Native file-channel Open/Close
integration remains open. Null94, invalid NUL/wildcard names52, missing source53,
missing parent76, denied operations70 and other native failures75 use the native
VBA error path; no pathname or contents are copied into diagnostics.

The existing Core target checks actual binary copying, replacement of a longer
destination, missing-source failure preserving the destination and same-file
refusal preserving the source. Desktop core1/1 (1.77 s), fresh GUI-free core1/1
(1.69 s), updated Debian/build-ID/SQLite3.54.1 and all four fixture hashes pass.
Last broader baseline is typed make-table48/48 (87.98 s); no new target/ID.
F040/F068 remain open, fixed456/384/72 unchanged.

## Native Name statement

`Name sourceExpression As destinationExpression` now executes actual native
filesystem renaming for regular files and directories within one filesystem.
The compiler finds As outside strings, brackets and parentheses; embedded As
text inside a pathname is not treated as the delimiter. Normal scalar/String
coercion and native errors apply. Execution occurs only in active branches.
Assignments to local Name/FileCopy/MkDir/RmDir/ChDir identifiers retain ordinary
assignment binding instead of being captured by file-statement dispatch.

Linux renameat2 with RENAME_NOREPLACE supplies the atomic no-overwrite operation.
An existing destination, including an existing directory or link, raises58.
No check-then-overwrite fallback is used when the kernel/filesystem does not
support the operation. Terminal source symlinks and special files refuse3251.
The working directory and its ancestors also refuse3251; directory moves must
not silently change the interpreter's active file context. Cross-filesystem
operations and Windows drive/UNC mapping remain unsupported3251. Parent symlink,
concurrent source replacement, Windows open/share/readonly rules and complete
error precedence/reference equivalence are not claimed.

Null94, invalid NUL/wildcard pathname52, empty/missing source53, missing parent76,
denied operation70 and other native failures75 use the existing VBA error path.
Diagnostics do not contain supplied paths or file contents.

Actual core cases rename a binary file, parse a parenthesized literal pathname
containing As, reverse the move, preserve a pre-existing destination, rename a
directory with its contents and execute local keyword-named variable assignments.
Final desktop core1/1 (1.78 s), fresh GUI-free core1/1 (1.69 s), updated Debian/
build-ID/pinned SQLite3.54.1 and all four original fixture hashes pass. Last
broader full remains preceding multi-record milestone48/48 (84.19 s). No added
target/ID; F040 remains open, fixed456/384/72.

## Native Kill

Unqualified Kill pathnameExpression now compiles to a native operation evaluated
only in active branches. A local identifier named Kill still supports ordinary
assignment. The normal scalar/default-member path coercion and VBA error handler
are used. Actual UTF-8/relative Linux regular files are unlinked; no shell command
or recursive deletion is used. The unlink syscall cannot delete directories.

Null94, NUL52, empty/missing file53, missing parent76, denied/read-only file70
and wrong-kind/other operation75 remain explicit. Read-only here means no Unix
write permission bits; this does not claim Windows attribute/ACL equivalence.
Wildcard deletion, terminal symlinks, drive/backslash mapping, Windows open-file
sharing behavior and exact error precedence remain unsupported/unvalidated.
Path replacement races and external open handles have no Windows oracle.

The existing core target deletes a synthetic binary test copy, preserves its
source and directory, checks an inactive branch and distinguishes missing file
from missing parent. No original fixture is a deletion target. No new test
program or catalog ID; F040 remains open, fixed456/384/72.

Reference: [Kill](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/kill-statement).

Final filesystem segment: full current desktop48/48 (84.89 s), fresh
headless1/1 (1.83 s), current Debian/build-ID/SQLite/four hashes pass.
Original queries78/100; no Windows result oracle. No new test target/ID.
