# Native bankers Round in VBA and Access SQL

Reproduced the Access-SQL adapter's delegation to SQLite round: the half-case
regression fails before the change (0/1). Round is now rewritten to a private
Jet callback and shares the native C++ helper with VBA expressions. Supported
Double half cases round to even, including positive/negative 2.5/3.5 and the
Microsoft 0.12335/0.12345/0.12355/0.12365 and 0.00005 examples. Raw SQLite's own
round function and quoted SQL text are unchanged.

The bounded helper supports nonnegative precision 0..15, Null, numeric text,
Boolean/Empty and explicit Byte/Integer/Long values. Numeric scalar defaults are
evaluated once. Explicit native Currency values quantize fixed-point units
without a Double conversion, preserving large fractional values; checked overflow
raises 6. Required expression/arity raises 450, invalid negative precision 5,
Null precision 94, type mismatch 13 and Nothing 91. Optional omitted precision
uses zero. Single, precision above 15 and unvalidated scaling overflow diagnose
3251. Exact SQL integers beyond the existing 2^53 envelope are not silently
converted to a rounded Double.

The core/round_documents Qt/VBA/DAO/SQL regressions pass 2/2 (1.00 s), including
actual calculated controls, both VBA and SQL results stored through real QueryDef
parameters, pending Text/accepted Value, explicit record save and a mixed valid/
invalid UPDATE leaving no partial changes. Currency 922337203685477.1234 rounds
to the exact text 922337203685477.123 and is stored as TEXT; overflowing Currency
prevents subsequent SQL. GUI suite now has 32 CTests. Full validation is in
PROGRESS.md.

This is not complete Round/IEEE/locale/type-promotion or Jet Currency SQL parity.
Imported Currency/Single metadata remains ambiguous; native explicit Currency
rounding does not make SQLite REAL an exact money store. High precision,
Single, Date/Decimal, all argument validation order/result-type/Windows boundaries
and complete original business execution remain open. No Windows/Wine/Excel
runtime or new app dependency is introduced.

Reference:
https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/round-function
