# Direct native form recordset field reads

Me.Recordset.Fields exposes actual loaded table/query result fields by name or
ordinal, Count, Item, default indexing and For Each. Direct recordset default and
bang reads use those same live fields. Native DAO.Field/Fields identity and named
Index arguments use the existing binder. This is a read-only native facade, not
an editable DAO dynaset or a claim that a query is updateable.

A held field reads the form's current stored row after normal bookmark/selector
navigation. Values come from the actual loaded typed snapshot; unsaved widget
or record-field drafts remain separate until saved. An unavailable current stored
row raises 3021. Snapshot replacement on Requery invalidates held fields and
collections with 3420; reacquiring Fields uses the new snapshot. Closed/destroyed
forms reject further reads with 2450. Independently held RecordsetClone cursors
retain their documented separate lifetime.

Name, Type, Size, OrdinalPosition, Value, FieldSize and GetChunk delegate to the
actual native DAO snapshot. Unsupported/unprovable metadata remains an explicit
error, including frozen-source Attributes. DataUpdatable is False because this
facade has no field-write path. Values preserve native typed Null/numeric/text/
binary representations; binary GetChunk uses actual bounded bytes. Field writes,
AppendChunk, Properties and unsupported direct-recordset navigation remain errors.
Field reads do not save the form or navigate any clone. GetChunk argument scalar
defaults are resolved before rechecking the live source and positioning the
snapshot; indexed field lookup rechecks form lifetime/source after index coercion.

The existing navigation fixture exercises actual VBA Fields(Index:=...), bang
reads and typed For Each, held-field navigation, blocked unsaved edits, stale
source rejection, closure, read-only SELECT rows and actual BLOB/partial bytes.
Validation results and any correction are recorded in handoff/progress. F054/F160
remain open for full editable recordsets, concurrency/locks and Access oracle.
No new CTest or catalog ID is introduced.
