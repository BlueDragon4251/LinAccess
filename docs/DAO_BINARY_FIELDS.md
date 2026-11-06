# Native DAO binary field values

DAO Recordset Field.Value now reads actual SQLite BLOB storage as a copied Byte()
Variant, including zero-length BLOBs. Physical BLOB fields expose native DAO type
11 (dbLongBinary). SQL-query field typing remains subject to existing proven type
metadata; unavailable Field.Type still refuses explicitly. This is a native SQLite
mapping, not proof of original Jet/OLE field metadata or container extraction.

During Edit/AddNew, Byte() assignment is accepted only for declared BLOB/BINARY/
VARBINARY/LONGBINARY fields and stages a real QByteArray binding. Empty Byte()
creates a non-NULL zero-length BLOB. NULL remains SQL NULL. Other scalar assignment
coercions to binary fields are explicitly unsupported; Byte() into a non-binary
field raises type mismatch before mutation. A BLOB stored in a declared non-binary
field, or a non-BLOB scalar stored in a binary field, is refused rather than silently
converted. Existing scalar field conversion is unchanged.

Values are copied; no recordset/widget/host pointer is stored in a Byte() result.
Normal Update, CancelUpdate, transaction, cursor and closed-field guards remain.
The existing actual VBA/SQL DAO case reads 00FF42, changes and saves the middle byte,
checks pending Value, saves an empty binary, cancels a pending NULL-to-binary edit,
then checks a held Field fails with 3420 after Close. Independent SQL verifies exact
hex bytes, typeof(blob), zero length and NULL after execution.

GetChunk/AppendChunk, QueryDef binary parameters, OLE/attachment import and original
MDB/ACCDB binary round trips, OLE rendering and full Access type/coercion/error
parity remain open. The general Qt form/report scalar converter still refuses
binary display; this change does not turn binary data into text or publish bytes.
No new CTest target or fixed-catalog item.

## Stored binary FieldSize

Field.FieldSize now reports the current non-NULL stored binary observation's byte
length as a Long while the cursor is in browse mode. It obtains the QByteArray size
without first allocating a copied VBA Byte() value. Zero-length BLOB size is zero;
read-only assignment fails with383. Current-record, closed and transaction guards
apply. Non-binary/Memo encodings, NULL cardinality and pending Edit/AddNew size
semantics are explicitly unsupported rather than substituted with memory/string
length. Original Jet/OLE database storage measurement is not claimed.

The existing actual VBA case checks three bytes, Long identity, read-only383 and
an empty BLOB size0; independent SQL already verifies stored bytes/type/length.
Reference: [DAO FieldSize](https://learn.microsoft.com/en-us/office/client-developer/access/desktop-database-reference/field-fieldsize-property-dao)
distinguishes database byte count from Unicode string/memory length. No new target
or catalog item.

Field.Size now separately exposes the defined LongBinary metadata size0 on
proven native BLOB fields, including empty recordsets. It does not replace
FieldSize's current stored data length. See [schema field scope](DAO_TABLEDEF_FIELDS.md).


## Bounded stored binary GetChunk (2026-11-06)

Recordset Field.GetChunk(Offset, Bytes), including named arguments, copies the
requested range of actual non-NULL native BLOB storage into an independent Byte()
Variant. Both arguments use checked Long conversion. Offset0 starts at the first
byte; a positive requested count exceeding the remaining bytes returns the actual
remaining bytes. Embedded0/255 bytes are retained; changing the returned array
does not change the stored field. The returned range is limited to16MiB. Existing
recordset/thread/transaction/current-record guards apply, and closure remains3420.

This deliberately supports stored browse-mode observations with an offset inside
nonempty data and a positive byte count. NULL/empty fields, offsets outside the
value, nonpositive counts, Edit/AddNew observations, Memo/encoding and unsupported
source types refuse3251 pending reference validation; their full Access result/
error precedence is not guessed. Wrong BLOB storage is13, arity450, Null numeric
arguments94 through existing checked conversion. This is copying from the native
loaded rowset, not streaming original MDB/OLE containers. AppendChunk, original
OLE/attachment extraction/roundtrip and full DAO binary semantics remain open.

Existing DAO binary cases verify named middle-byte extraction, first/tail ranges,
over-request truncation, isolated copied bytes, boundary/Null/arity/closed errors
and unchanged independent persisted SQL bytes. No new CTest target or catalog ID;
F016/F055/F121/F122 remain open and456/384/72 unchanged.
Reference: https://learn.microsoft.com/en-us/office/client-developer/access/desktop-database-reference/field-getchunk-method-dao
This current public contract is not an Access2003 execution oracle.
