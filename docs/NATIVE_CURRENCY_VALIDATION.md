# Native Currency ValidationRule and ValidationText

Native Currency fields support the existing bounded numeric comparison grammar:
comparisons joined by And/Or, or Between. Rule bounds must be explicit decimal
literals with at most four places and within the existing exact native Currency
storage range. They pass through the same Currency conversion/storage codec as
the actual field values; predicates compare the corresponding exact REAL
encoding. Exponents, excess precision, calls, identifiers and out-of-range bounds
are refused before the field declaration changes. This intentionally bounded
contract does not cover complete Jet expression or Currency range semantics.

Creation, attached field changes, persisted metadata reading and the actual field
editor share that contract. ValidationText supplies the existing checked SQL
constraint error message. NULL behavior uses the existing validation predicate
implementation; Required remains an independent field constraint. Table rebuilds
retain field rules/messages and use the existing row identity/value digest and
savepoint checks. A tighter rule violated by existing rows fails atomically.

Lossless numeric type changes can now retain compatible Currency rules/messages.
A Double rule with a five-place boundary is rejected when targeting Currency;
its original field type and rule remain intact. Attached DAO Field.Type remains
readonly; conversion is the explicit native field designer operation.

App/GUI/DAO targets built. The existing GUI check passed in12.45s with the
unchanged20s watchdog. The first DAO case reused a field handle invalidated by a
previous column addition; its fixture now reacquires the actual field. Only the
affected DAO check was rebuilt/repeated and passed1/1 in2.44s. Actual cases cover checked
VBA creation, SQL enforcement/message, invalid-bound retention, failed rule
tightening with preserved rows, compatible type changes and rejected five-place
Double-rule conversion. GUI creates/edits a rule and retains it across actual
Currency/Double conversions before verifying raw SQL constraint enforcement. No new CTest or catalog identities. F014/F015/F055/F121 and full
Windows validation semantics remain open.
