# Exact native Jet Numeric import

Jet4 Numeric fields now use checked original precision/scale and17-byte raw
coefficients from the read-only MDB metadata helper. The helper's bounded
--decimal-values mode records the original source hash, column order, NULLs and
raw bytes. Native parsing validates sign, four little-endian32-bit words,
coefficient precision and scale before creating a96-bit Decimal value. CSV values
must agree numerically with those original bytes and their NULL/row alignment.
No conversion through Double is used.

MDBTools backend type names are insufficient: the independent Numeric fixture's
SQLite schema declares these columns INTEGER, while MySQL reports numeric(28,4)
and numeric(28,28). Native metadata supplies the authoritative type. MDBTools'
numeric(p,s) exporter uses col_scale as precision and col_prec as fractional
scale, despite those member names. This is checked against the asymmetric
independent fixture and the upstream backend implementation:
https://github.com/mdbtools/mdbtools/blob/v1.0.1/src/libmdb/backend.c
The observed runtime metadata library is1.0.1.

Copied values use DECIMAL_EXACT and the existing24-byte sortable numeric key.
Original precision/scale persist in _la_table_schema. A native SQLite CHECK
accepts only valid exact keys within that declared precision/scale, or NULL when
otherwise permitted. Decimal source values outside their declaration, absent
precision/scale, raw/CSV disagreement and mismatched source snapshots refuse;
the existing import savepoint/project transaction rolls back partial work.
The invariant CSV mode can also import explicitly declared Numeric(p,s) values;
it does not claim independent MDB evidence from a source filename.

The data grid displays/copies exact text. Proven physical fields now support
bounded exact-fit grid/DAO edits (NATIVE_DECIMAL_FIELD_WRITES.md); computed
outputs retain their read-only contract. DAO/SQL values retain Decimal/dbDecimal20.
Native creation of Decimal fields, full Field.Size/Precision/Scale/Properties identity, non-NULL Decimal
source defaults, all validation/format/relationship semantics, signed-zero/scale
bit identity and original-format writing/round trips remain open. Non-NULL
Decimal defaults explicitly refuse. Existing approximate Numeric imports are not
silently promoted; metadata that claims Decimal over REAL backing refuses and
requires a fresh source import. Java is not an application dependency.

Optional independent fixture: tests/jet_binary_reference/JetDecimalReference.java
uses the existing Jackcess5.0.4 fixture tool (SHA256
19a7494a56f439a57fda85c059b019b665802085cec5ae8186d1e028cc1bdfbc).
It creates and independently reopens a real Jet4 MDB with28-digit precision,
scales4/28, positive/negative values, values above2^53,1e-28, zero and NULL.
No business data, MDB or Java binary is committed or shipped. With an existing
JDK21 and the jar/Commons Logging paths supplied locally:

```
java -m jdk.compiler/com.sun.tools.javac.Main -classpath "$jackcess_jar" -d "$decimal_scratch/classes" tests/jet_binary_reference/JetDecimalReference.java
java -classpath "$decimal_scratch/classes:$jackcess_jar:$commons_logging_jar" JetDecimalReference "$decimal_scratch"
LINACCESS_JET_DECIMAL_REFERENCE="$decimal_scratch/decimal-probe.mdb" ctest --test-dir build -R '^table_import_contracts$' --output-on-failure
```

The generator refuses to overwrite its fixture. The existing import target checks
raw source values, source metadata, exact storage/display, numeric ordering, NULL,
rollback, failed approximate writes, changed CSV refusal and unchanged MDB hash.
This is independent Jet-format evidence, not a new Windows Access2003 execution
oracle. F014/F017/F055/F067/F068 remain open; fixed456/386/70 unchanged.

The old persistent-counter import check now follows the already documented
[AddNew current-record contract](DAO_ADDNEW_CURRENT.md): preserve the prior row,
then navigate by LastModified to inspect the insertion. The stored ID was already
correct; no generator/cursor implementation was changed for that check.

Validation: app/metadata helper built. Existing GUI/import targets passed2/2
in17.43s (17.08s/0.35s), including the real Jet fixture and actual copy/paste
into a native text field. Seven protected source hashes rechecked unchanged.
No new test target or Access execution oracle.
