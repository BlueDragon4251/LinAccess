# Native ASCII and UTF-16 character functions

Asc/AscW/Chr/ChrW now execute in the native C++ expression runtime and actual
form/VBA/DAO workflows. Asc/AscW return Integer, with first-character behavior;
AscW observes the first Windows UTF-16 code unit and its signed 16-bit value.
Chr/ChrW use checked rounded Long input and return a String. Empty Asc arguments
raise 5, Null 94, type/Error 13, Nothing 91, arity 450 and Long overflow 6.
Scalar defaults are evaluated once.

Asc and Chr deliberately support the codepage-independent ASCII range 0..127.
Other ANSI/DBCS values diagnose 3251 until an explicit codepage contract/reference
exists; Linux UTF-8 is not silently treated as the Windows ANSI codepage. ChrW
supports representable UTF-16 units in -32768..65535 (negative values wrap to
unsigned units); out-of-range units raise 5. The existing strict UTF-8 string
representation rejects isolated surrogate output with 3251. AscW of a valid
supplementary character can still return its signed leading surrogate unit.

Core regression covers ASCII/control/NUL, first-character and result subtype,
rounding, Greek/signed Unicode/emoji, scalar-once and failure boundaries. Native
character_function_documents uses real bound/calculated Qt controls, DAO Long/
Text parameters, actual SQL ASCII/Unicode round trips, pending Text vs accepted
Value and explicit record save. Unsupported conversion prevents subsequent SQL.
Targeted core/native tests pass 2/2 (1.15 s); GUI suite now has 30 CTests.

## Unchanged original KeyPress procedures

The original corpus has three KeyPress procedures containing nine Asc and three
Chr calls; all nine Asc literals are ASCII. A new access_2003_layouts regression
reads and SHA-256-verifies the original module bytes without changing them, then
binds each handler through its event prefix to prepared native test controls.
It executes 384 direct ASCII input cases and 285 printable Qt keyboard cases,
comparing actual field effects/parameter outcomes between direct execution and
native event wiring. All three unchanged procedures execute in these contexts;
access_2003_layouts passes (12.61 s). No original source/data values are printed.

The wiring comparison uses the same interpreter, not an independent Windows
oracle.
It does not prove original full-form setup, original business workflows, login,
full keyboard/locale/event identity or complete Access character-function parity.
ANSI/DBCS, byte-string functions, dollar aliases, isolated-surrogate BSTR storage,
SQL function rewriting and all Windows boundaries remain open. Full validation
is recorded in PROGRESS.md.

References:
- https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/asc-function
- https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/chr-function
