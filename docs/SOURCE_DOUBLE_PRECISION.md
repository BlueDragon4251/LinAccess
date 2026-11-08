# Native exact Double import scope

New actual MDB copies preserve finite MDB_DOUBLE values from their original
binary64 bits. The existing read-only helper's --double-values mode uses public
mdb_fetch_row offsets/mdb_get_double and serializes sixteen hexadecimal digits;
no decimal text is used as storage truth. Native Single/Double column declarations
are checked against the helper's actual types and the table schema. The shared
access_floating_metadata decoder verifies version, columns, table, row widths,
NULLs, bit widths, finite values and source SHA-256. Negative zero and nonfinite
values refuse explicitly. Payload16MiB and one-million-cell bounds remain.

The numeric CSV verifier uses classic-locale C++ defaultfloat at the same
significant precision as MDBTools1.0.x (Single8, Double16), with a bounded decimal
normalization for spelling/exponents. It never reparses a Double export as a
floating storage value. This matters at Double.MAX_VALUE: its sixteen-digit
text rounds beyond the finite Double range, although the actual stored bits are
valid. Smallest subnormal, maximum finite value, ordinary fractions, positive
zero and NULL are copied without that overflow/precision loss. A halfway Single
format case also checks deterministic rounding; parent locale is not changed.

The low-level API explicitly distinguishes caller-owned Csv from NativeMdbBits.
CSV is the backward-compatible default and does not open a diagnostic sourceFile
label. It supports existing technical data scenarios without pretending the data
came from an original MDB. Full-database and Inspector copies explicitly select
NativeMdbBits; actual-source tests do likewise. Missing native proof/source never
falls back to CSV. Both modes preserve the existing structural index metadata.
Full import obtains raw values from its already verified immutable MDB snapshot,
using the same file as its schema and CSV. Original sourceFile stays the log
identity. Empty structure creation validates declarations; the actual data phase,
including zero-row copies, validates raw row counts, NULLs, CSV and before/after
snapshot hashes within the existing atomic transaction.

The optional JetDoublePrecisionReference.java writes and independently reopens
actual Jet4 files with known bit patterns, including the16-digit counterexample,
finite extremes and a negative-zero refusal file. The existing import target
checks exact SQLite and DAO Double bits, Type7/Size8, full/direct copies, bad CSV,
mode provenance and atomic refusal. No business records or app Java dependency.
Enable with LINACCESS_JET_DOUBLE_PRECISION_REFERENCE.

JetDoubleOracle.java independently reads supplied test copies through Jackcess,
without loading/executing VBA. It writes private sorted per-row Double-bit tuples
and column-name headers under a source-SHA directory. Output stays outside Git;
existing directories are never overwritten. LINACCESS_DOUBLE_ORACLE_ROOT enables
the existing original Kassa regression to check all tuples against actual SQLite
values. Both supplied Kassa versions contain13 affected tables,57 Double columns,
69405 non-NULL Double values and514 NULLs; no negative zero/nonfinite values were
found. Before the fix, the actual decimal export failed to reproduce2404 original
Double row tuples in two tables. Original/source hashes remain unchanged.

These are source data and native field-value proofs, not a Windows Access2003
SQL/event/rendering oracle. Legacy arbitrary CSV, exact Currency/Decimal, generic
SQL/default writes, query/type lineage, unsupported source patterns, original-
format writes and broad Jet/locale semantics remain open. Fixed catalog456/384/72;
F014/F015/F041/F055/F105/F122 remain open. No new CTest target.

Validation: corrected current full Debian13 desktop48/48 (143.10 s), both actual
Kassa versions with the full independent Double oracle, and the uploaded version
with78/100 original queries passed at their documented scope. Current app/package/
app-helper Build-IDs, actual packaged Double helper output and independent source
bits/hash, SQLite3.54.1, four original and three upload hashes passed. No new
Double distribution matrix or Windows oracle is claimed.
