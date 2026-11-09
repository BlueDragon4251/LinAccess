# Native table Recordset, Index and Seek

Stand: 2026-11-09. This is bounded native functionality within F054/F055/F160,
not full Jet locking, indexed cursors or Access2003 differential proof.

Database.OpenRecordset defaults to table type1 for actual main user tables;
TableDef.OpenRecordset also defaults to1. Explicit2/4 remain dynaset/snapshot.
Queries default to2, and QueryDefs/SQL sources refuse type1. dbOpenTable is exposed
as the native Long constant1; Recordset.Type remains a readonly Integer.
Existing query write-capability limits remain independent of opening kind.

Table-type Recordsets expose Index. Selection checks the complete actual index
catalog and uses its ordered fields/directions for the actual SQL snapshot read.
Only structurally non-null own native numeric keys are accepted. Text, Date,
Currency, GUID, nullable/unproven keys and linked providers are not inferred.
Single keys require the owned quantized SINGLE_REAL backing declaration.
Index changes retain existing row identity/bookmarks; they do not rewrite table
rows. Pending edits refuse3219. A change of order while clones remain open
refuses3251; clones created after selection inherit the selected index and have
independent positions. Clearing Index restores the unindexed native snapshot.

Seek accepts every key in a supported index (up to the existing native10-field
bound), including named Comparison/Key arguments. Numeric keys use shared field
conversion. Equality works for ASC/DESC definitions; <, <=, >, >= work for fully
ascending indices. The search scans the checked ordered native snapshot forward
for =/>/>= and backward for </<=. This is not a B-tree performance claim or a live
Jet cursor. Duplicate ties use native row identity ordering, not an Access leaf
page identity guarantee. Missing/Null/partial keys, descending range seeks and
unsupported families refuse explicitly. NoMatch is set on failure and the
current record becomes unavailable; a retained bookmark can restore it. Full
Access BOF/EOF/error identity after failure remains unproven.

Selection, reading Index, Seek and Requery validate the current definition.
Changes require Requery/reopening rather than silently using another index.
Existing schema/rollback/closed/family guards still apply. Concurrent external
row visibility, independently changing clone indices, partial-key semantics,
full comparison/collation behavior and the complete Windows error/event oracle
remain open. Existing Find functionality is separate from this indexed scope.

Public behavioral references (current documentation, not a2003 execution oracle):
- https://learn.microsoft.com/en-us/office/client-developer/access/desktop-database-reference/recordset-index-property-dao
- https://learn.microsoft.com/en-us/office/client-developer/access/desktop-database-reference/recordset-seek-method-dao

Existing Core/UI/DAO/original-event4/4 passed (122.64s; actual original import
113.62s), app built. Original uploads3/3 and fixtures4/4 SHA256 unchanged.
No new CTest target, full-suite, Access2003 differential or CI claim.

Canonical own non-null Date keys are now also accepted, including mixed numeric/
Date indices. Date keys require actual Date Variants and use shared whole-second
storage conversion; string keys refuse13. Comparison follows canonical timestamp
order. Old/imported/noncanonical Date backing remains unsupported. Existing
Date/OLE/precision limits apply; no full Jet date-identity oracle is implied.
Focused DAO1/1 passed (2.20s), app built; the preceding original-event run is not
claimed as freshly repeated for this additional key family.

Omitted trailing keys in supported non-null composite indices now receive the
documented Null suffix. Equality cannot match that suffix; ascending ranges use
its lowest-key position, so >=prefix locates the corresponding range. Explicit
Missing at the suffix is accepted; explicit Null inputs, nonterminal holes and
missing first key remain unsupported. Numeric and typed Date prefixes are
checked in existing DAO1/1 (2.25s), app built. No new target/full-suite claim.
