# Native Scripting.FileSystemObject subset

CreateObject("Scripting.FileSystemObject") and New Scripting.FileSystemObject
now construct a real independent C++ filesystem adapter. No COM server, Windows
DLL, Qt object, host reference or persistent database handle is retained.
TypeName returns FileSystemObject and VarType returns9 from the actual object.

Supported members:

- FileExists(FileSpec): checks an actual regular file.
- FolderExists(FolderSpec): checks an actual directory.
- BuildPath(Path,Name): joins native Linux path components without requiring existence.
- GetAbsolutePathName(Path): resolves the process working directory and normalizes a native absolute path.
- GetFileName(Path), GetBaseName(Path), GetExtensionName(Path), GetParentFolderName(Path): extract lexical path components without accessing the filesystem.
- GetFile(FilePath), GetFolder(FolderPath): return actual native file/folder objects, with the bounded members below.
- CreateFolder(Path): creates one actual directory and returns its Folder object.
- CopyFile, MoveFile, DeleteFile: real bounded single-file operations described below.

Existence checks distinguish files, directories and absent paths; empty paths
return False, and missing/not-a-directory results return False. Other errors use
native codes: Null94, arity450, NUL/invalid pathname52, denied access70 and other
filesystem failures75. Existence wildcards, drive/UNC mapping and rooted BuildPath
names are explicit3251 boundaries. Path normalization does not claim Windows
case/drive/share/attribute behavior. Linux status follows filesystem symlinks.
Diagnostics do not include supplied paths or contents. Methods perform actual
filesystem reads; the adapter stores no predetermined existence results.

The qualified native type is accepted by the existing local/parameter/result
binder, with Nothing initialization and explicit native interface checks. Wrong
Scripting classes raise13 before replacing the target. The same closed native
family is used by the runtime's type/default validation; this is not arbitrary
reference-library resolution, short-name search, As New support, COM activation
or a claim that every object/array/storage context has an Access oracle. Remote
activation remains refused. Other FSO members remain unimplemented3251; no files/
folders/collections are fabricated. The SQL host's global object-state/context
restrictions remain unchanged.

Actual core VBA constructs and releases late/early objects, returns one through a
typed function, passes one to a typed parameter, checks TypeOf/TypeName/VarType,
refuses a wrong class without losing the prior reference, and checks real UTF-8
files/directories, missing/wrong-kind paths, named arguments and working-directory
resolution. Initial core/DAO/original3/3 (9.78 s); final full desktop48/48
(85.42 s), fresh GUI-free core1/1 (1.69 s), updated Debian/build-ID/pinned
SQLite3.54.1/four original fixture hashes pass. Original queries remain78/100
without Access result oracle. No new target/ID; F040/F042/F091 remain open,
fixed456/384/72.

Reference: [FileSystemObject](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/filesystemobject-object),
[GetAbsolutePathName](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/getabsolutepathname-method).

Lexical component methods accept both Windows and Linux separators, preserve
UTF-8/case, and need no existing file. Drive-only input returns an empty string;
ordinary drive-qualified and UNC file paths can be parsed without drive mapping.
Trailing separators, UNC roots, dot-directory components, leading-dot extension
semantics and drive-relative parent extraction remain explicit3251 boundaries.
This does not enable Windows paths for physical filesystem operations.

Path-component increment: desktop core1/1 (1.73 s), fresh GUI-free core1/1
(1.72 s), current Debian/build-ID/SQLite/four fixture hashes pass. The broader
48/48 above precedes this increment. No new test target or catalog ID.

## Native GetFile objects

GetFile(FilePath) resolves an actual existing regular Linux file and returns
a real independent native File object (TypeName File). Path is the
normalized absolute path, Name is its final component, and Size reads the
filesystem again on each access. No host, stream or database handle is retained.
Sizes through signed Long maximum retain Long; larger exact sizes use Double.
Beyond the exact Double integer range raises6 instead of returning rounded data.
Missing files raise53; final symlinks and nonregular objects remain bounded.

The additional DateLastModified property uses the shared native FileDateTime
local-time conversion and rereads the filesystem. Qualified Scripting.File
locals/parameters/results use explicit native identity. The scalar default is
Path, so VarType(file) now reports String8 while TypeName(file) remains File;
CStr(file) reads the path. File property writes remain explicit3251.

Additional bounded native file operations and ParentFolder are described below.
Creation/access timestamps, attributes,
sharing, Windows drive mapping and race/identity semantics remain unimplemented
or unvalidated. A retained path does not pin the original inode. This is not a
full Windows File object contract.

References: [GetFile](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/getfile-method),
[File](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/file-object),
[Size](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/size-property-filesystemobject-object).
Numeric size policy cross-checked against the public Wine scrrun implementation;
no Windows Access result oracle is claimed.

GetFile validation: desktop core1/1 (1.82 s), fresh headless1/1 (1.72 s),
current Debian/build-ID/SQLite/four original fixture hashes pass.

The default-Path contract is cross-checked against the public scrrun IFile IDL
(DISPID_VALUE); [VarType documentation](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/vartype-function)
specifies inspection of the default property when present.

File binding/default/timestamp increment: targeted core/DAO2/2 (2.62 s),
fresh GUI-free1/1 (1.66 s), current Debian/build-ID/SQLite/four hashes pass.

Final filesystem segment: full current desktop48/48 (84.89 s), fresh
headless1/1 (1.83 s), current Debian/build-ID/SQLite/four hashes pass.
Original queries78/100; no Windows result oracle. No new test target/ID.

## Native GetFolder and Files enumeration

GetFolder(FolderPath) resolves an actual existing native directory and returns
a Folder object. Qualified Scripting.Folder and Scripting.Files bindings use
explicit native identity. Folder Path/Name read path components; Path is the
scalar default (VarType8, TypeName Folder). Folder.Files returns a real native
Files object with Count and For Each. Each Count/enumeration reads the directory
again; the For Each operation holds a bounded snapshot of native File objects,
whose metadata remains live. No open directory handle survives the operation.

Only regular files are included; directories are excluded from Files. Hidden
Linux entries are included. Terminal folder links, any link/special-file entry,
permission errors and more than10000 inspected entries fail explicitly; they
are not silently skipped. Enumeration order is native and has no Windows order
or concurrent-mutation oracle. Recursive size, general folder
mutations, timestamps and complete Windows collection/attribute/identity behavior
remain unimplemented or unvalidated.

References: [GetFolder](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/getfolder-method),
[Files](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/files-collection).

SubFolders uses the same bounded fresh directory snapshot and returns native
Folder entries through a Folders collection with Count/For Each and qualified
Scripting.Folders identity. File.ParentFolder and non-root Folder.ParentFolder
return actual existing parent Folder objects; a root parent remains3251. Folder
paths with trailing separators normalize consistently before Path/Name/collection
use. Files and folders are selected by actual filesystem kind, not name guesses.

Folder increment: expanded desktop core1/1 (1.86 s), fresh GUI-free1/1
(1.79 s), current Debian/build-ID/SQLite/four original fixture hashes pass.
Last broader48/48 (84.89 s) precedes this increment. F091 stays open.

## Native CreateFolder

CreateFolder(Path) creates one actual native directory and returns a typed
Folder object usable immediately. It does not invent missing parents or overwrite
existing objects. Existing target58, empty/NUL52, Null94, missing parent76 and
permission70 flow through native VBA error handling. Linux permissions/umask
apply; Windows drive/ACL/sharing and external race behavior remain bounded.
Held SubFolders collections observe the new directory on the next Count or
enumeration because they read the filesystem again. No recursive creation is
added and the original fixture files are not used as write targets.

Reference: [CreateFolder](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/createfolder-method).

Named-argument binding follows the public IFileSystem IDL vocabulary: GetFile
FilePath, GetFolder FolderPath, CreateFolder Path, GetFileName Path and
GetAbsolutePathName Path. These correct earlier documentation-placeholder
names; positional calls keep their behavior. Public scrrun IDL is a reference
for this bounded interface, not a Windows Access execution oracle.

CreateFolder/IDL parameter validation: desktop core1/1 (1.87 s), fresh
GUI-free1/1 (1.77 s), current Debian/build-ID/SQLite/four hashes pass.

## Native file operations

FSO CopyFile(Source,Destination,[OverWriteFiles=True]), MoveFile(Source,Destination)
and DeleteFile(FileSpec,[Force=False]) now perform actual single-file operations.
File objects support Copy(Destination,[OverWriteFiles=True]), Move(Destination),
Delete([Force=False]) and writable Name. Successful File.Move/Name updates the
held object's path; a failed move/name assignment leaves it intact. Source
metadata is read from the filesystem; no file contents or outcomes are mocked.

The VBA FileCopy/Name/Kill statements share the native operations implementation.
Copy uses actual binary copy with explicit overwrite policy; Move/Name use Linux
renameat2/RENAME_NOREPLACE and never replace a target; deletion uses unlink and
cannot remove directories. A destination ending in / denotes an existing real
folder and retains the source filename. Terminal folder links remain refused
even with a trailing separator. Plain File.Name cannot escape its parent.

Overwrite=False rejects an existing target58. Native read-only here is the lack
of Unix write permission bits; Copy refuses such a destination, and Force=True
allows deletion without changing those bits when directory permissions permit.
This is not Windows attribute/ACL/open-sharing equivalence. Wildcards, terminal
links/special files, cross-filesystem moves, inode pinning/races, metadata/ACL
copy fidelity and exact Windows error precedence remain unsupported/unvalidated.
An operation can have normal OS side effects before an external-race failure; no
multi-file or filesystem transaction/rollback is claimed.

The existing core target executes a full synthetic binary file workflow through
VBA FSO and File members: copy, no-overwrite error, rename, move, delete, live
metadata and empty destination-folder check. Existing statement binary/rename/
delete cases use the same helpers. No new test target/ID; F091 remains open.

References: [CopyFile](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/copyfile-method),
[MoveFile](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/movefile-method),
[DeleteFile](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/deletefile-method).

Final file-operation segment: desktop48/48 (86.15 s), fresh GUI-free1/1
(1.92 s), current Debian/build-ID/SQLite3.54.1/four original hashes pass.
A private native Force probe confirms read-only preservation with False70 and
actual deletion with True. No new repository test target/ID; fixed456/384/72.

Native OpenTextFile/OpenAsTextStream/CreateTextFile and owned TextStream handles
are bounded separately in [NATIVE_TEXT_STREAM.md](NATIVE_TEXT_STREAM.md).

## Native collection lookup and child-folder creation

Files.Item(Key) and SubFolders.Item(Key) resolve actual direct children by their
plain native filename. The default collection call, including indexed Folder.Files
and Folder.SubFolders properties, uses the same lookup. Explicit Set returns the
real File/Folder object; scalar default evaluation reads that object's Path.
Keys are case-sensitive native text names, not fabricated numeric positions.
Null is 94, missing entries 53 (Files) / 76 (Folders); missing/extra arguments 450.
NUL/wildcards, separators, parent traversal, Windows paths and non-text keys
retain explicit errors/boundaries. Lookup uses the existing fresh bounded directory
snapshot, including its link/special-file and 10,000-entry restrictions.

SubFolders.Add(Name) creates one real immediate child and returns its native
Folder object. It validates the held parent, does not create intermediate parents,
and reports existing names as 58 without replacing the previous caller reference.
Subsequent Count/Item/enumeration sees actual filesystem state. Files has no Add.
The interface parameter names Key and Name follow the public scrrun IDL; this is
not proof of Windows case, path-key coercion, race or filesystem parity.

The existing core fixture exercises actual early-bound VBA, named Item/Add,
object/scalar defaults, indexed parent properties, Unicode child creation,
duplicate creation, missing/Null/traversal errors and independent filesystem
checks. The first run exposed missing indexed-property forwarding; that precise
problem was corrected before rerunning the same target. Final result is recorded
in the handoff/progress entry. No additional CTest or catalog IDs are introduced.

Reference: [Files collection](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/files-collection),
[Folders collection](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/folders-collection),
[public scrrun interface definitions](https://github.com/wine-mirror/wine/blob/master/dlls/scrrun/scrrun.idl).

## Native Folder.Name assignment

Folder.Name now renames the actual held directory within its existing parent.
The adapter validates a plain native filename and the real directory, then uses
the existing atomic no-replace Name helper. Existing destinations fail with58;
invalid traversal/path-component requests retain3251 and Null retains94. The
held object's Path/Name/scalar default changes only after successful rename;
aliases of that same object observe the new path. Other independently obtained
File/Folder objects retain their original pathname references.

The existing helper's terminal-symlink, active working-directory ancestor,
filesystem/error and platform boundaries remain. No folder tree copy/deletion,
Windows attribute/drive/collation or inode-pinning equivalence is added. Other
Folder properties remain guarded; F091 stays open.

The existing core filesystem scope creates an isolated real directory and file,
renames a Unicode folder twice through early-bound VBA, checks retained object
identity and live child lookup, refuses collision/traversal/Null without changing
the held path, and independently reads the original payload after relocation.
All temporary objects are confined to the fixture directory and cleaned there.
No additional CTest target or catalog ID is introduced.

Folder.Move(Destination) and FileSystemObject.MoveFolder(Source,Destination) now
relocate one actual directory to an exact absent native destination. Both reuse
the existing same-filesystem atomic no-replace helper. Folder.Move updates the
held object's path only after success; independent path objects are not silently
rewritten. Named arguments use the real Source/Destination signatures.
Existing destinations fail58; trailing-separator destination-folder interpretation
remains3251; Null fails94. Wildcards, Windows paths, cross-filesystem copy fallback,
merged trees, recursive directory copying and full Windows sharing/attribute
semantics are not introduced. The existing core scope checks both actual APIs,
Unicode, aliases, collision/Null/directory refusal, live collection refresh and
independently preserved file contents. No new test target or catalog point.


## Actual native folder metadata

Folder.Size traverses actual files and subfolders, including hidden names, adding
logical byte lengths (not allocated disk blocks). Fresh reads observe actual
changes. Empty directories contribute zero. The native numeric representation is
Double with exact integer bounds through2^53-1; Windows Variant subtype identity
is not claimed. Overflow6, traversal over1million entries or depth256 refuses3251.
Symlinks and special files refuse3251, rather than following links, silently
skipping unknown objects or blocking on streams. Permission errors70 and other
metadata errors75 propagate. Missing/root-invalid folders retain76. Traversal is
not an atomic filesystem snapshot; concurrent replacement, mount/reparse and
Windows junction/provider behavior remain unproved.

Folder.DateLastModified reads the actual current directory timestamp and uses the
existing local-time checked VBA Date conversion shared with File.DateLastModified
and FileDateTime. No creation-time approximation. Both getters reject arguments450;
no writable metadata property is added. Native folder moves/renames continue to
update the held path through the existing implementation.

One existing core_tests bundle passed5.56s; app/core built. Real nested/hidden
files total10bytes, a sparse file lifts the total past2GiB without allocating its
logical size, symlinks refuse, deletion invalidates reads, and actual Date values
match a deliberately set filesystem time. Typed native VBA reads both properties.
No new CTest/catalog IDs; F091 remains open,456/386/70 unchanged.

Public contracts (not an Access2003/Windows FSO execution oracle):
https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/size-property-filesystemobject-object
https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/datelastmodified-property


## Actual access and creation timestamps

Native File and Folder now expose DateLastAccessed and DateCreated. Linux statx
requests the actual access/birth timestamp and checks its returned availability
mask and the actual object kind. No inode-change time, modification time or current
clock substitutes for creation time. Filesystems/kernels without the requested
timestamp refuse3251; permission errors70, absent files53/folders76, other failures75.
Final symlinks and changed object kinds refuse3251. This does not resolve all
ancestor-link races, network providers, Windows reparse points or filesystem
access-time update policies. Linux relatime/noatime remains the real filesystem
policy. Properties are read-only, accept no arguments, and return native Date
through the existing local-time/range conversion. That conversion now also accepts
checked Unix seconds directly, avoiding a needless filesystem-clock roundtrip for
statx values. Existing modified-time/FileDateTime behavior retains the same path.

App/core built, one existing core_tests bundle passed5.45s. Actual file and directory
access/modified times deliberately differ by a day; both direct objects and typed
VBA preserve that distinction. Creation time, when available, remains distinct
from the deliberately old access time; unavailable creation time is checked3251.
Missing objects and getter arity are covered. No new CTest/catalog IDs;456/386/70.
No Windows FSO/Access execution-oracle or complete interface claim.

Public contracts:
https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/datecreated-property
https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/datelastaccessed-property
