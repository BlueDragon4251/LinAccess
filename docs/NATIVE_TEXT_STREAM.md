# Native bounded TextStream reading

FSO OpenTextFile(FileName,[IOMode=1],[Create=False],[Format=0]) and
File.OpenAsTextStream([IOMode=1],[Format=0]) now return a real native
Scripting.TextStream read object. The object owns a POSIX regular-file handle;
O_NOFOLLOW/nonblocking open and descriptor checks refuse terminal links and
special files without opening a blocking FIFO. No host/database handle is held.

Supported: Read(Characters), ReadLine, ReadAll, AtEndOfStream, Line, Column and
Close. Reads use actual file bytes, preserving embedded NUL and CR/LF sequences.
ReadLine strips its terminator; CRLF advances one line. Positions start at1.
Explicit Close and destruction release the read handle. Qualified native type
binding uses explicit interface identity rather than TypeName guessing.

This first adapter accepts ASCII content only. Non-ASCII, Unicode/default-code-
page formats, write/append modes, Create=True, Skip and AtEndOfLine remain
explicit3251. No locale encoding or Unicode conversion is fabricated. Read
supports positive counts through1000000 within the remaining byte extent;
zero/negative/over-end requests are3251 pending reference validation. ReadLine
is bounded to1000000 bytes, ReadAll to16MiB. Failed content/limit reads restore
the prior position where seeking succeeds; I/O failures are75. Reading a line
already at EOF raises62. Closed-stream error identity remains explicit3251.

A process-wide weak registry records device/inode identities of native open read
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
