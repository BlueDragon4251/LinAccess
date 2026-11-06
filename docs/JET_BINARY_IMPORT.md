# Bounded raw Jet binary import

Full MDB imports and the table inspector now explicitly request MDBTools
`-b hex`. The actual CLI format is unprefixed ASCII hex. The typed import API
accepts it only with `AccessBinaryEncoding::MdbToolsHex`; its default still refuses
non-NULL binary CSV. No codec is guessed from a payload which happens to look
hexadecimal. Odd lengths, prefixes, whitespace and non-hex characters are refused.
Validated bytes bind as SQLite BLOBs, never as UTF-8 text. Unquoted ASCII0x1e still
means SQL NULL; an explicitly encoded empty value remains a non-NULL empty BLOB.

The inspector displays byte counts and NULL rather than pretending to render an
OLE object or displaying a huge hex string. Both individual-table and full-import
paths use the same explicit codec. Full import retains its one verified source
copy and whole-project transaction; table import retains its savepoint. Source
MDBs are never opened for writing by these paths.

MDBTools stdout is now drained and bounded while the process runs, rather than
only checked after all output was buffered. Bounds are20MiB stdout,64KiB stderr,
and the existing30-second execution deadline. Oversized output kills the child
and fails the import. The cell decoder accepts at most16MiB decoded data; the
whole-table20MiB CSV cap is stricter in practice because hex doubles payload size.
These are output/cell bounds, not a promise about the child library's own heap.
Exit errors and damaged-page diagnostics still reject partial results.

This preserves opaque value bytes, not OLE package interpretation/rendering,
ACE attachment/complex-field handling, encryption, all Jet storage/type semantics,
large-object streaming or original-format writing/round trips. F014/F016/F017 and
the complete interoperability/oracle requirements remain open. Native DAO Byte()
Value, FieldSize and bounded GetChunk read the copied bytes; that is distinct from
full Access Variant/container identity. No VBA is extracted/executed by this import.

## Independent technical fixture

The optional fixture generator in
`tests/jet_binary_reference/JetBinaryReference.java` creates real Jet4/V2000 MDBs
using Jackcess5.0.4 and independently reopens them read-only. It contains no
business records. Its normal table has NULL and7/4097/65539 binary bytes, covering
NUL/high bytes, quote/comma/newline and inline/multi-page values. Raw sidecar files
are the independent expected values. A separate11MiB field exercises output-limit
rollback. No generated MDB, binary payload or Java dependency is committed or
shipped in the application. Java/Jackcess are fixture-generation tools only.

With an existing JDK21 and an empty private scratch directory:

```sh
# Download jackcess-5.0.4.jar from Maven Central into the scratch directory.
# Required SHA-256:
# 19a7494a56f439a57fda85c059b019b665802085cec5ae8186d1e028cc1bdfbc
java -m jdk.compiler/com.sun.tools.javac.Main \
  -classpath "$binary_fixture_dir/jackcess-5.0.4.jar" \
  -d "$binary_fixture_dir/classes" tests/jet_binary_reference/JetBinaryReference.java
java -classpath "$binary_fixture_dir/classes:$binary_fixture_dir/jackcess-5.0.4.jar" \
  JetBinaryReference "$binary_fixture_dir"
export LINACCESS_JET_BINARY_REFERENCE="$binary_fixture_dir/binary-probe.mdb"
ctest --test-dir build -R '^(table_import_contracts|mdb_inspector_gui_tests)$' --output-on-failure
```

Set `binary_fixture_dir` to that empty scratch directory first; the generator
refuses to replace an existing fixture. Maven artifact:
https://repo.maven.apache.org/maven2/com/healthmarketscience/jackcess/jackcess/5.0.4/jackcess-5.0.4.jar
The source/CLI signatures, artifact checksums and actual exported format were
checked before use. This is a real technical Jet fixture, not a Microsoft Access
2003 execution oracle or evidence about an original business OLE object.

The existing table-import target compares actual MDB→SQLite bytes with sidecars,
checks SQL NULL/type/length, native DAO Type11/Size0/FieldSize/Value/GetChunk, source
hashes and full oversized-import rollback with an untouched sentinel row. Existing
inspector GUI tests use real controls and verify the partial import bytes and
byte-count preview. Malformed/unspecified codecs and native empty BLOBs are checked
separately. There is no new CTest target or fixed-catalog point.

Validation: targeted import/inspector2/2 (12.26 s), then current full Debian13
48/48 (91.47 s) with both technical MDBs and original recipe PDFs enabled. The
newly supplied Kassa copy again passes the existing import/cell/counter checks and
100-artifact query inventory:73 ordinary plus5 parameterized queries execute;
22 boundaries remain explicit. Current .deb/payload/Build-ID/SQLite3.54.1 and all
four original fixture hashes pass. Earlier local distribution evidence remains
separate; no fresh platform matrix or GitHub CI success is claimed for this stage.
Catalog456 total/384 scoped confirmations/72 open full requirements, unchanged.
