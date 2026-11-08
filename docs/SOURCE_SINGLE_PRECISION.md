# Native exact Single import scope

Real Jet Single imports now obtain binary32 values directly through the existing
read-only MDBTools helper, bypassing its lossy %.8g decimal export for storage.
The --single-values mode reads public MDB_FLOAT column row offsets after
mdb_fetch_row and uses mdb_get_single plus IEEE754 bit serialization. It does not
bind numeric string buffers, execute VBA or write the original file. Actual
column names, row order, NULL values and raw bits remain separate from SQL schema
text and are not persisted as extra user-data metadata.

The source index metadata now also declares actual Single columns. The importer
checks those names against the proven SINGLE source declarations, then requests
only those fields. A strict versioned decoder checks source table, ordered names,
row widths, exact eight-digit hexadecimal values and source SHA-256. There is a
one-million-cell/16MiB payload bound. The helper streams the source hash before
and after scanning; the importer checks it before and after the transactional
copy. Original filenames/values are absent from generic error messages.

For data-bearing phases the number of raw rows must equal the CSV row count.
NULL markers must agree. Each finite raw value must reproduce the actual current
MDBTools %.8g CSV numeric value before its exact binary32-to-binary64 expansion
is bound into SQLite REAL. This guards mismatched export snapshots without using
the decimal representation as stored source truth. Full import's initial empty
schema-creation phase validates declarations without scanning row data; the later
copy phase enforces the actual count, including zero-row copies. Whole import
reads its protected MDB snapshot. Inspector imports use the same guarded path.
The low-level API selects NativeMdbBits explicitly for actual MDB copies; its
caller-owned Csv mode does not consult a diagnostic source filename.
Caller-owned transactions and existing savepoints retain atomic rollback.

Finite binary32 values, including the smallest positive subnormal and maximum
finite value, are represented exactly in binary64 SQLite REAL storage. Positive
zero and NULL are preserved. NaN, infinities and negative zero explicitly refuse
instead of silently normalizing them; SQLite REAL's negative-zero bit identity
is unproven. An arbitrary legacy caller-supplied CSV without native source proof
keeps its previous narrower contract. Double/Decimal/Currency precision, generic
SQL writes/defaults/arithmetic, every Jet variant, arbitrary source concurrency
and full Windows coercion/rendering remain open. No blanket Float parity claim.

JetSinglePrecisionReference.java independently creates/reopens a real Jet4
precision file containing the %.8g counterexample (bits509022793), ordinary
fractions, finite extremes, positive zero and NULL, plus a separate negative-zero
refusal file. LINACCESS_JET_SINGLE_PRECISION_REFERENCE enables the existing import
target's real full/direct copy, exact known bits in SQLite/DAO, deliberately
mismatched CSV values/NULL/count, source hashes and negative-zero atomic refusal.
The original lossy CSV is independently checked to reconstruct the wrong first
bit pattern, making the before/after measurable. Pure decoder checks also refuse
malformed/nonfinite payloads. No new CTest or app Java/.NET dependency.
F014/F015/F041/F055/F105/F122 remain open; fixed catalog456/384/72.

Validation: relevant existing GUI/DAO/import/original Kassa5/5 (27.76 s); final
affinity guard/import target1/1 (1.05 s). App/package/helper Build-IDs, actual
packaged --single-values output against known independent Jet bits/source hash,
SQLite3.54.1 and all four original/three upload hashes passed. Last full suite
74d75c7:48/48 (85.91 s), preceding this increment.

The shared codec now also handles Double; see SOURCE_DOUBLE_PRECISION.md.
Classic-locale C++ significant formatting verifies CSV without mutating locale.

Unveränderliche a11b95e-Folgeprüfung: Debian12/GCC12/Qt6.4 und Ubuntu24.04/GCC13/Qt6.4 jeweils5/5 (28,31 s /27,15 s), Paketinstallation, Originalhashes/Backend und tatsächlich installierter --single-values-Helfer mit unabhängig bekannten Originalbits/SHA bestanden. Diese Snapshots schließen DAO.Indexes und Single-Rohpräzision ein, noch keine spätere Double-Korrektur. Keine GitHub-CI-/Vollsuitebehauptung für diese Matrix.

Native Single creation and supported DAO/SQL/data-sheet writes now use a shared
binary32 converter and explicit Single declaration. Native literal defaults,
older literal default insertion and direct-column query type identity are
covered. Final GUI/DAO2/2 (6.12s), preceding import1/1 (0.06s). This preserves
the original raw-bit import scope above and does not prove unsupported complex
writes, full arithmetic/locale/rendering or Windows parity. Details and current
limits: DAO_TABLEDEF_FIELDS.md. Fixed456/384/72 unchanged.
