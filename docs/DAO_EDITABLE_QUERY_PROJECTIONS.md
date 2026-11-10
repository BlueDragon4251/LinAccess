# Native DAO physical-field query projections

The same strict single-table validator now maps explicit physical-field lists to
native DAO mode2 dynasets. Subsets, reordered fields and qualified names retain
exact visible field count/order, physical types and SourceField/SourceTable.
The physical row identity is selected separately and need not be visible. UPDATE
and DELETE still target that identity and checked WHERE membership; clones retain
the projection. Existing full-star behavior is preserved.

Only selected fields supply native draft defaults/type metadata. Omitted physical
columns are left to their real SQLite defaults, including a Date default; no
attempt is made to read a missing default field from a shorter buffer. Unfiltered
AddNew is supported when generator fields needed for checked reservations remain
projected. If a generated field is omitted, AddNew refuses3251 before any draft
or allocation. Filtered AddNew remains3251. Actual aliases, duplicate physical
fields, computed columns, parameters, joins and keyless targets are not promoted;
mode4 remains read-only. Complete Windows refresh/locking/metadata identity and
saved-query-target resolution remain open.

App/DAO/query-model/Currency built; one existing bundle5/5 passed3.98s (DAO3.48).
Review found that omitted generator fields also need a preallocation boundary;
that boundary was added before mutation and only the affected existing DAO check
was rerun1/1 in3.43s. Actual keyed SQL proves subset/reordered writes, hidden-key
selection, clone values, source metadata, omitted Text/Date defaults, hidden-
generator refusal without rows/drafts, alias/duplicate/snapshot/keyless bounds.
An actual VBA procedure edits a selected field without exposing its key.
No new CTest/catalog IDs;456/386/70 unchanged; F054/F055 remain broader open scopes.
Latest checked package/Original is stillc67cbd9e before query-writing changes;
latest inventory4ebc3f6c100/78/22 is not a Windows result oracle.
