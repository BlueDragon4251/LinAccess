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
