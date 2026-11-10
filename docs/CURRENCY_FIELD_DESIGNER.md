# Native Currency field authoring and checked type changes

The field editor now offers Currency for new fields and numeric target types,
shows the actual four-place type/Size8 and retains working default/Required/
display/rename/reorder operations. Its tooltip gives the actual supported amount
range. Currency ValidationRule/ValidationText editing remains disabled while that
semantics is unimplemented. New fields use the actual DAO creation path.

The native designer type-change operation now accepts Currency alongside the
existing numeric types. Every non-NULL stored numeric value and actual default
must round-trip through the target storage with the same binary64 numeric bits.
Original row identities and other values retain the existing migration digest/
constraint/savepoint proof. Numeric defaults becoming Currency normalize only
after this exact round-trip check; binary serialization digits are not treated as
a new source literal. Original/unknown schemas, counters, altered values/rounding,
unsupported ranges and Currency targets carrying numeric rules/messages remain
refused. Attached DAO Field.Type stays readonly383; this is the explicit native
designer operation. No complete Jet type-conversion/Currency range claim.

App/DAO/GUI targets built; one existing two-check bundle passed2/2 in22.84s:
mainwindow_documents19.25s with its unchanged20s watchdog, dao_execution3.59s.
Actual GUI creates a Currency field, verifies it, changes it to Double and back,
edits its default, independently checks stored values and deletes the technical
field. Backend cases cover Currency/Double round trips, integral/NULL retention,
exact Currency units, stored-Double-default normalization, and rejection of actual
rounding, range and rule cases without replacing source values/types. Existing
Currency schema/default/Required/index/reorder/rename cases pass in the same DAO
bundle. No new CTest/catalog ID;456/386/70 unchanged; F014/F015/F055/F105/F121 remain
broader open requirements. No Windows conversion or full GUI/distro oracle.

Private published predecessor6786526d was tree-verified. Latest checked package/
Original is stilla23c7009 before subsequent changes; a fresh grouped package/
Original/query-inventory check follows this saved stage. Earlier inventory is
4ebc3f6c100/78/22 without an Access result oracle.
