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
