# Native VBA file metadata (2026-11-04)

FileLen(PathName) and FileDateTime(PathName), including VBA-qualified names and
named arguments, read actual Linux filesystem metadata. Absolute and relative
UTF-8 native paths are supported; relative paths use the process working
directory. No file contents are read or modified by these functions. Static
expression validation does not execute filesystem calls.

FileLen returns a real VBA Long containing the file's byte size. Sizes above
2,147,483,647 refuse overflow 6. FileDateTime converts the actual last-write
instant to local calendar time and returns a native Date Variant at second
precision, retaining the existing OLE/Gregorian conversion. It does not fabricate
a creation timestamp or replace the Date value with a formatted display string.
Native directory timestamps and symlink following use Linux metadata behavior.

Missing files refuse 53; absent/non-directory parent paths 76; permission denial
70; unsupported targets/other metadata failures 75. NUL/wildcard/invalid path
forms refuse 52, Null paths 94 and bad argument counts 450. Windows drive and
backslash/UNC path mappings refuse 3251. This is a measured native subset, not
complete Windows filesystem, share-lock, locale-display or error-priority parity.
FileLen's VBA-open-handle snapshot/LOF semantics remain open with the wider file
I/O runtime. External concurrent changes are not an atomic metadata snapshot.
The later SQL integration below adds measured SQL/DAO evaluation. The remaining
Dir/GetAttr/Open/Close/file-statement APIs remain open.

The existing core target uses real disposable files with Unicode names and an
embedded NUL byte in their contents, relative paths, known last-write instants
under two local timezones and a sparse file above the VBA Long boundary. It
checks actual result types, missing/invalid paths and Null/count errors. No new
CTest target and no original fixture mutation are involved.

Primary VBA references (current docs, not an Access 2003 execution oracle):
https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/filelen-function
https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/filedatetime-function

VBA checkpoint validation: full desktop 48/48 (93.70 s), freshly rebuilt GUI-free
core 1/1 (1.65 s), Debian payload/permissions/Build-ID/pinned SQLite and all four
original fixture SHA-256 checks pass.


## SQL and DAO integration (2026-11-04)

Access SELECT expressions now compile FileLen/FileDateTime to native SQLite
callbacks backed by the same C++ VBA metadata functions. Parameters bind actual
path values; compilation/metadata inspection does not perform file I/O. The
functions are not marked deterministic and are not added to pure-expression
rewrite allowlists, so materializing a query again reads current metadata.
Unsupported complex rewrite/dependency cases retain existing guards.

Proven projections expose canonical INTEGER4 and DATETIME types. Actual DAO
Fields therefore return Long/Date values, dbLong/dbDate Type metadata and native
recordset values. Field.Type additionally recognizes proven DATE/DATETIME
sources as dbDate; unknown/other unimplemented source types retain 3251.
SQL function errors retain the checked native path/permission codes through the
existing sanitized error adapter without including private file paths.

Existing query-model/DAO targets use actual temporary files, typed parameters,
known timestamps, real size changes and missing-file/Null failures. A real DAO
OpenRecordset checks both Field.Type and the actual Long/Date Value types. No
new CTest targets. Full Windows filesystem/handle/locking and Jet error-priority
parity remain open.

SQL/DAO checkpoint: full desktop 48/48 (94.77 s), unchanged GUI-free runtime
rechecked 1/1 (1.89 s), rebuilt Debian package payload/permissions/Build-ID/
pinned SQLite and all four original fixture SHA-256 checks pass.
