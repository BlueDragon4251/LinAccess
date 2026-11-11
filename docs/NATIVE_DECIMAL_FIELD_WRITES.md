# Bounded exact native Decimal field edits

Physical Decimal fields with proven precision/scale and the matching generated
SQLite CHECK now support exact-fit DAO and data-grid writes. Computed query
results, absent/mismatched write declarations and read-only cursors stay read-only.
The constraint recognizer finds actual CHECK clauses, ignoring quoted literals,
identifiers and comments; text merely containing the function name proves nothing.

Conversion occurs before changing the pending value. NULL remains NULL, Empty
uses the existing CDec conversion, and Decimal/invariant text retain all digits.
Values requiring rounding or exceeding precision refuse6; no truncation or silent
scale reduction is introduced. Floating inputs retain the existing bounded CDec
contract rather than acquiring a new Windows conversion guarantee.

DAO retains exact values through Edit/AddNew/Update and GetRows. DataUpdatable
reflects the physical write context; Decimal Size/precision-property identity is
not fabricated. Default-property evaluation that closes/changes a cursor refuses
before replacing its buffer. Original MDB files remain read-only sources.

The grid uses a real text editor rather than a Double editor, supports validated
text/typed paste, ordinary record undo, range undo/redo, NULL clearing and native
transactional Save. Internal history keys are decoded only while replaying owned
history; external binary clipboard values do not become Decimal inputs. Formatting
and normal clipboard text use the existing exact Decimal representation.

Qt reports this declared type as QString. Its original-record cache reads through
virtual EditRole and therefore captures formatted Decimal text instead of the
stored key. Decimal tables now obtain original storage values directly from the
query model and bind them in UPDATE/DELETE filters. An atomic count guard requires
one matching row; duplicate rows without a unique identity refuse before changing
them. Existing transaction rollback
and pending-edit retention apply. Other table types retain their prior
path. This does not claim full Access append ordering, external concurrency,
trigger/rekey behavior or arbitrary row identity.

Bound physical form writes are documented in NATIVE_DECIMAL_FORM_WRITES.md.
Native Decimal field creation, arbitrary SQL action conversion, non-NULL source defaults, complete DAO properties, regional coercions,
rounding-at-scale behavior and original-format write/round trips remain open.
F014/F017/F055/F067/F068/F340 remain open; fixed456/386/70 unchanged.

Existing import/GUI scopes cover the real Jet reference and unchanged source,
DAO exact edits/NULL/AddNew/GetRows, rejected scale changes, real line editing,
paste/history/Save, NULL clearing, numeric storage, unique/ambiguous row writes
and rejection of fake constraint text. No new CTest/catalog ID or Windows oracle.

App/helper build and existing GUI/import bundle passed: 2/2 in17.78s
(17.52/0.25), including the independent Jet fixture.
