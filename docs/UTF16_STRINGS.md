# Bounded UTF-16 string functions

Runtime text crosses native APIs as UTF-8. VBA Len/Left/Right/Mid now count and
index UTF-16 code units, rather than bytes. German characters and complete
surrogate pairs round-trip correctly; embedded NUL is included. Left$/Right$/
Mid$ are recognized. Non-dollar substring functions return Null for Null input;
the String-returning variants raise 94. Scalar default properties are resolved
once, including Null. Length/start arguments use bounded Long conversion with
banker's rounding: invalid positions/lengths raise 5, overflow 6, invalid arity
450 and numeric Null conversion 94.

A shared validated code-unit counter backs the native SQL Len adapter. It
preserves Null and embedded NUL, counts supplementary characters as two units,
rejects malformed UTF-8 with 13 and participates in pure projection-alias
expansion. SQLite's separate length function keeps its SQLite semantics.

The numeric lexer also recognizes leading-dot decimals and E/e exponents.
Malformed exponents fail explicitly and literal overflow reports 6. This does
not implement every VBA numeric suffix, literal type or constant-folding phase.

Native tests cover direct expressions, scalar getter call counts, complete VBA
procedure execution and On Error/Err.Number, the SQL result model/alias paths,
malformed text, embedded NUL, boundary lengths and error codes. The original
restored VBA/query token inventories contain no direct calls to these functions;
no additional original-procedure execution is claimed. Original corpus
regressions remain mandatory.

## Remaining scope

Isolated UTF-16 surrogates cannot currently cross the UTF-8 runtime boundary. A
substring that splits a surrogate pair raises 3251 rather than emitting corrupt
text. Full BSTR/binary-string storage, LenB/MidB, primitive declared-type storage
sizes for Len, code-page conversions, locale-aware case/collation/Like behavior
and broader string intrinsics remain open. SQL scalar coercion also retains
SQLite's limited numeric/Boolean type representation; this is not complete Jet
Variant or VBA string compatibility.

## Trim family

Trim/LTrim/RTrim and their dollar variants now remove only U+0020 spaces at
the requested ends, preserving tabs, line breaks and other characters. Variant
forms preserve Null; String forms raise 94. Scalar default properties are
resolved once and invalid arity raises 450. Core tests check actual VBA
procedures and On Error/Err.Number as well as expressions and getter effects.
No direct trim calls occur in the restored original VBA token inventory, so
this adds no original-handler execution claim.

## Search and comparison

InStr uses one-based UTF-16 positions and bounded Long start arguments. StrComp
normalizes comparison results to -1/0/1. Both honor Option Compare and explicit
vbUseCompareOption/vbBinaryCompare/vbTextCompare/vbDatabaseCompare, preserve
Null string arguments and resolve scalar defaults once. Invalid numeric text
raises type mismatch 13; Null numeric arguments raise 94. Binary comparisons
use UTF-16 ordering, including supplementary characters. Tests cover actual VBA
Option Compare execution, empty strings, omitted arguments and error handling.

Text/database comparison delegates to the runtime host. The GUI host uses Qt
case-insensitive comparison; the standalone host supports ASCII folding and
rejects unsupported non-ASCII comparisons. InStr scans equal-code-unit-length
substrings without splitting surrogate pairs. Windows/Jet linguistic expansion,
contraction and locale ordering are not implemented or claimed. Explicit InStr
LCIDs are rejected with 3251. No Windows Access oracle has verified this scope.
