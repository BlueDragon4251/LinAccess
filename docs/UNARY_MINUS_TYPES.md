# Native unary minus result types

MS-VBAL 5.6.9.3 and 5.6.9.3.1 specify the effective scalar types.
The native VBA evaluator now preserves Integer for negated Byte, Boolean and
Empty values, Integer/Long for in-range negations, Single for Single, and Date
for a negation inside the supported Date range. Existing checked Currency,
Decimal and Null paths remain in use. Scalar defaults are resolved once through
the existing evaluator. Double and numeric String operands retain Double.

This bounded change does not alter Access Eval's separate path. Signed Integer
and Long minima and out-of-range Date negations retain their previous behavior:
the distinction between a statically declared operand and a Variant operand
still requires complete declared-type binding. This does not claim that their
exceptional promotion/error behavior is complete. Unary plus is unchanged.
F041 remains open, and the fixed catalog remains 456/386/70.

The existing core_tests matrix checks actual result VarType and values for Byte,
Integer, Long, Single, Double, Boolean, Empty, Date, String, Null, Currency and
Decimal. The application and core target built successfully; the single bundled
core run passed in 5.29 seconds. No new CTest IDs were added.

Before this scope, source 8fd1d50 also passed DEB build-ID, pinned SQLite 3.54.1,
MDBTools 1.0.1, package content/permissions/desktop and no-fixtures checks at
/workspace/work/sql-typed-arithmetic/package/linaccess-0.7.0-Linux.deb.
Its original database inventory and original-path checks passed 2/2 in 170.19
seconds (48.34/121.86). Inventory remains 100 classified, 78 executed and 22
bounded cases, without a complete Windows query oracle. The missing original
macro-source boundary remains. All seven upload/fixture hashes were unchanged.
That checked package predates this unary-minus scope.
