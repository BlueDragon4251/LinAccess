# Checked Currency variable storage

Declared Currency module fields, procedure locals, parameters, function results
and defaults now use the existing signed 64-bit four-decimal Currency value.
Checked assignment uses the existing exact decimal conversion and ties-to-even
rounding. Defaults retain Currency kind; VarType/TypeName retain 6/Currency.
Known reference mismatches reject with 13; matching and parenthesized temporary
references retain their bounded behavior. Currency For counters reject with 3251
before procedure effects pending full loop semantics.

The DAO value adapter now transfers Currency as canonical decimal text instead
of rejecting it as an unsupported object or passing through a Double. This
enables exact binding to declared TEXT parameters. Actual native Qt/VBA/DAO tests
persist the upper Currency limit 922337203685477.5807 into a TEXT column, verify
the SQL storage class and preserve the complete decimal value. This is an exact
text-transfer regression, not a claim that SQLite REAL/NUMERIC or Jet Currency
columns and SQL arithmetic are fully emulated. SQL PARAMETERS Currency remains
unsupported, and Parameter type metadata/assignment timing remain incomplete.

Core and native tests cover declared defaults/locals/parameters/results, ByVal
input preservation, exact same-Currency addition at the upper boundary, overflow
retaining the prior module value, public field assignment rounding, invalid
arguments with no SQL effects, matching/rejected/temporary references and the
pre-effect unsupported loop guard. Targeted tests: 4/4 (42.72 s). Full validation
is recorded in PROGRESS.md. Original business execution coverage is unchanged.

Full mixed numeric promotion/literal types, arrays/classes, Currency loops,
object-default/Variant/reference/locale/compiler/error timing and complete SQL/
DAO/MDB Currency storage semantics remain open. Len storage identity is not added.


Declared Currency For/Next counters now use exact signed64 ticks for the start,
frozen end/step and checked increment. Positive/negative four-decimal steps,
full-range boundaries, zero iterations, explicit zero step with Exit For,
body-written counters and the final post-loop value retain Currency kind.
Increment overflow raises6 before replacing the prior counter. Existing structured
loop stacks, Exit For and error recovery remain shared; no Double conversion is
used for this counter family. The existing execution budget still bounds zero-
step loops. Core/DAO checked together2/2 (3.99s), application built. Actual Access
2003 loop-oracle coverage, mixed Variant counter promotion and complete numeric
For/Next semantics remain open; no full parity assertion. Existing F041/F043/F045/
F105 only, fixed456/384/72 unchanged. Continue implementation without a final.
