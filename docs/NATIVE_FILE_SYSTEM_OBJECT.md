# Native Scripting.FileSystemObject subset

CreateObject("Scripting.FileSystemObject") and New Scripting.FileSystemObject
now construct a real independent C++ filesystem adapter. No COM server, Windows
DLL, Qt object, host reference or persistent database handle is retained.
TypeName returns FileSystemObject and VarType returns9 from the actual object.

Supported members:

- FileExists(FileSpec): checks an actual regular file.
- FolderExists(FolderSpec): checks an actual directory.
- BuildPath(Path,Name): joins native Linux path components without requiring existence.
- GetAbsolutePathName(PathSpec): resolves the process working directory and normalizes a native absolute path.
- GetFileName(PathSpec), GetBaseName(Path), GetExtensionName(Path), GetParentFolderName(Path): extract lexical path components without accessing the filesystem.

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

GetFile(FileSpec) resolves an actual existing regular Linux file and returns
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

Only these four property reads are implemented. File mutations, parent folders,
creation/access timestamps, attributes,
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

GetFolder(FolderSpec) resolves an actual existing native directory and returns
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
or concurrent-mutation oracle. Item lookup, recursive size, folder
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
