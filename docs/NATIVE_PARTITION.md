# Native VBA and SQL Partition

`Partition(Number, Start, Stop, Interval)` now executes in the C++ VBA
expression engine. SQL calls are translated to the SQLite scalar adapter, which
uses that same implementation. Function-call rewriting leaves SQL `PARTITION BY`
tokens alone. Named VBA arguments use the existing binder.

The bounded implementation rounds inputs through VBA Long conversion (nearest
even), validates Start >= 0, Stop > Start and Interval >= 1, and returns Null if
any resolved argument is Null. Missing required arguments raise 450; invalid
ranges raise 5. It right-aligns both colon-separated boundaries to the width
required by Stop + 1 and Start - 1. Below/above-range values have a blank lower/
upper endpoint. The last bucket is clipped to Stop. Intermediate arithmetic uses
64-bit integers to avoid C++ overflow. Stop = 2147483647 is explicitly rejected
with 3251: extreme VBA Long boundary behavior has no Access 2003 oracle yet.

Examples, with spaces preserved:

```text
Partition(0, 0, 99, 5)       -> "  0:  4"
Partition(-1, 0, 99, 5)      -> "   : -1"
Partition(100, 0, 99, 5)     -> "100:   "
Partition(1001,100,1010,20)  -> "1000:1010"
```

The existing core target checks named calls, boundaries, Null, nearest-even
rounding and invalid intervals. The existing DAO target executes a real GROUP BY
query over seven nullable numeric rows and checks every bucket/count, including
Null and out-of-range groups. No additional CTest target was added.

Reference: [Microsoft VBA Partition](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/partition-function).
Padding and bounded bucket construction also agree with the public .NET
VisualBasic reference algorithm's stated VB3 compatibility rule. This is source
corroboration, not execution against Windows Access 2003, and establishes no
complete VBA/Jet parity.

Validation: desktop 48/48 (84.06 s); fresh headless core 1/1; Debian package,
matching binary build ID, pinned SQLite and all four original fixture hashes
passed. Catalog remains 456 / 384 bounded entries / 72 open requirements.


Typisierte native SQL-Funktionsargumente (2026-11-10)
--------------------------------------------------
RGB/QBColor und Partition erhalten bei nachgewiesenen einfachen Quellen die
Date/Boolean/Single/Currency/integralen Argumenttypen über den gemeinsamen
Scalar-Adapter. Date bedeutet den OLE-Date-Wert; True bleibt -1 trotz physischer
BOOLEAN1. Reine ISO-TEXT-Werte erhalten keinen geratenen Date-Typ. Projektionen,
unterstützte Verschachtelungen/Filter/Gruppierungen/Sortierungen bleiben gebunden.
RGB/QBColor-Projektionen tragen INTEGER4 mit tatsächlichem DAO dbLong; Partition
bleibt TEXT ohne geratenen Text/Memo-Typ. Vorhandene Rundung, Null-Regeln,
Palette, Bereichs-/Extrem-/Codepage-/Quellgrenzen bleiben erhalten. Dies erweitert
keine Windows- oder allgemeine Jet-Kompatibilitätsbehauptung.
