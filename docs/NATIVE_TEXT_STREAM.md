# Native bounded TextStream input/output

FSO OpenTextFile(FileName,[IOMode=1],[Create=False],[Format=0]) and
File.OpenAsTextStream([IOMode=1],[Format=0]) now return a real native
Scripting.TextStream object. The object owns a POSIX regular-file handle;
O_NOFOLLOW/nonblocking open and descriptor checks refuse terminal links and
special files without opening a blocking FIFO. No host/database handle is held.

Supported: Read(Characters), ReadLine, ReadAll, AtEndOfStream, Line, Column and
Close. Reads use actual file bytes, preserving embedded NUL and CR/LF sequences.
ReadLine strips its terminator; CRLF advances one line. Positions start at1.
Explicit Close and destruction release the file handle. Qualified native type
binding uses explicit interface identity rather than TypeName guessing.

ASCII and explicit UTF-16LE formats are implemented as described below.
Default-code-page formats and certain EOF/position cases remain
explicit3251. No locale encoding or Unicode conversion is fabricated. Read
supports positive counts through1000000 within the remaining byte extent;
zero/negative/over-end requests are3251 pending reference validation. ReadLine
is bounded to1000000 bytes, ReadAll to16MiB. Failed content/limit reads restore
the prior position where seeking succeeds; I/O failures are75. Reading a line
already at EOF raises62. Closed-stream error identity remains explicit3251.

A process-wide weak registry records device/inode identities of native open
streams, not paths or hosts. Shared copy-destination/rename-source/unlink guards
refuse mutation of these identities3251 while open. Close releases the guard,
including across hard-link aliases. This is a conservative native sharing bound,
not full Windows process/attribute/concurrency/open-sharing equivalence. External
process handles, races, replacement identity and asynchronous I/O remain open.

The existing core target reads real synthetic CRLF/LF text through native VBA,
checks line/column, fixed-length and remaining reads, EOF, held-stream deletion
refusal, then Close/actual binary overwrite/reopen/embedded-NUL reads. No new
test target or catalog ID; F040/F091 remain open, fixed456/384/72.

Validation: targeted core/DAO2/2 (2.67 s), fresh headless1/1 (1.68 s), current
Debian/build-ID/SQLite/four fixture hashes pass. A private manual probe verifies
hard-link guard/release, EOF62 and non-ASCII read position recovery. The broader
48/48 (86.15 s) predates this increment. No new test target or catalog ID.

References: [OpenTextFile](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/opentextfile-method),
[TextStream](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/textstream-object).

## Native ASCII writing and appending

OpenTextFile and File.OpenAsTextStream now implement modes1 (read),2 (write/
truncate),8 (append), with Format0. OpenTextFile Create=True can create a missing
file. FSO CreateTextFile(FileName,[Overwrite=True],[Unicode=False]) returns a
write stream; Overwrite=False uses atomic O_CREAT/O_EXCL and raises58 without
changing an existing file. The default follows IFileSystem interface metadata
and the method parameter table; no Windows execution oracle is claimed. Folder
CreateTextFile is supported for plain child names as described below.

Write(Text), WriteLine([Text=""]) and WriteBlankLines(Lines) write actual ASCII
bytes. WriteLine/WriteBlankLines emit CRLF; NUL is preserved. Complete text is
validated before writing so unsupported non-ASCII input writes no prefix. Each
call flushes its stdio buffer; normal partial OS/I/O failures can still occur and
are not filesystem transactions. Text is bounded to16MiB per call, blank-line
counts to positive1000000; unsupported ranges are3251. Read/position properties
on write streams and output on read streams are currently explicit3251.

The registry now distinguishes readers/writers under one mutex and refuses
writer overlaps or readers overlapping a writer. Registration and descriptor
validation happen before ftruncate, so a refused write-open preserves existing
bytes. Open uses no O_TRUNC before these checks. Native read-only permission
bits are checked before truncation. This remains a conservative native bound,
not Windows cross-process sharing/ACL/locking equivalence.

The existing core VBA workflow creates a file, writes text/CRLF/blank lines,
rejects a non-ASCII write before partial output, refuses read/write overlaps,
closes, appends, reopens and reads exact resulting bytes. A no-overwrite create
refusal preserves those bytes; a read-held write-open cannot truncate them.
No new test program/catalog ID.

Reference: [CreateTextFile](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/createtextfile-method).

Final input/output segment: full current desktop48/48 (87.77 s), fresh
GUI-free1/1 (2.23 s), zero compiler warnings in both builds; current Debian/
build-ID/SQLite3.54.1/four original hashes pass. Current private native
hard-link/close/EOF/encoding-recovery probe passes. Original78/100; no Windows
execution oracle. No new CTest target/ID; fixed456/384/72.

## Explicit UTF-16LE streams

Format=-1 and CreateTextFile Unicode=True now select native UTF-16LE. New/truncated
Unicode files receive one FF FE BOM. Reading/append require a complete even-length
UTF-16LE file with that BOM; big-endian, BOM-less, odd-byte and default-code-page
inputs remain3251. Unicode read/Create=True remains3251 before file creation.
Append validates the existing BOM and does not duplicate it; an empty append
file receives its BOM. Unicode append uses a native read/write descriptor to
validate the BOM, so native OS read permission is also required.

The existing strict UTF-8/UTF-16 conversion is shared in native_utf.hpp with the
VBA string runtime. UTF-8 invalid scalars remain errors. Read counts and columns
use UTF-16 code units; a request splitting a surrogate pair fails3251 and restores
position rather than returning an invalid UTF-8 value. ReadLine/ReadAll decode
actual complete units; output validates/converts the whole string before writes.
NUL and CRLF are preserved. Unicode ReadAll/write payloads are bounded to16MiB
of encoded input/output, respectively; default-code-page/ANSI extensions and
full Windows encoding/error/position behavior remain unvalidated.

The existing native VBA workflow writes ä/emoji/CRLF, appends ß, checks exact
little-endian BOM/UTF-16 bytes independently, refuses a split surrogate read,
then successfully reads the pair/line/remainder. No new CTest target or catalog
ID; F040/F091 remain open, fixed456/384/72.

Final Unicode validation: full current desktop48/48 (89.99 s), fresh GUI-free
1/1 (1.85 s), zero compiler warnings in both builds; current Debian/build-ID/
SQLite/four hashes pass. Private malformed BOM/surrogate/refused-create probes
pass. Original78/100; no Windows result oracle. No new CTest target/ID.

## Stream navigation and folder-local text creation

Skip(Characters) and SkipLine now use the actual existing read/line paths,
discarding the resulting text while retaining checked positions and encoding/
count/surrogate bounds. AtEndOfLine checks the next real ASCII/UTF-16 unit for
CR/LF without consuming it. Its EOF behavior remains explicit3251 pending an
Access reference result; it is not guessed from AtEndOfStream.

Folder.CreateTextFile(FileName,[Overwrite=True],[Unicode=False]) now creates
a real child through the same atomic create/guarded-truncate/encoding pipeline
as FSO CreateTextFile. This bounded method requires a plain child filename;
path separators, drive syntax and . or .. are refused3251 before filesystem
mutation. UTF-8 filenames and explicit UTF-16LE text output are supported.
Full Windows path/reference/default/error contracts remain open.

Existing core VBA checks a real line marker, SkipLine consumption, Skip followed
by a read containing NUL, and actual Folder-created Unicode file/readback.
No new CTest target or catalog ID; F040/F091 remain open, fixed456/384/72.

Navigation/folder increment: core/DAO2/2 (2.69 s), fresh GUI-free1/1
(1.92 s), zero compiler warnings; current Debian/build-ID/SQLite/four original
hashes pass. Last broader48/48 (89.99 s) precedes this increment.
