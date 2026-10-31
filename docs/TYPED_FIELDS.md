# Typed VBA field bridges

The runtime now asks hosts for typed field values and passes typed assignments
back to hosts. The original optional-string host API remains a compatibility
fallback; Qt forms and DAO table recordsets use the typed bridge.

Verified behavior:

- Numeric bound controls produce numeric addition, rather than text concatenation.
  An actual Qt button computes 12 + 3.5 as 15.5. Numeric values assigned to unbound
  controls retain their numeric kind for subsequent VBA expressions.
- Boolean controls/fields return Boolean Variants; CStr shows True/False. Native
  and imported table metadata determines the field type.
- The importer reads structural source types from MDBTools' MySQL backend because
  SQLite schema output erases Boolean identity. All 16 original Kassa Boolean
  columns are recognized. CSV True=1 becomes VBA/Jet True=-1 in the SQLite copy;
  False remains 0. The complete original-cell comparison uses this declared
  Boolean mapping, and DAO Field.Value is checked on original imported columns.
- NULL and empty text remain distinct through bound control loads, assignments
  and saves. Invalid numeric input fails without changing the persisted record.
- Integer assignments use tie-to-even rounding; text columns receive actual text
  bindings. ISO date/time strings normalize to the canonical import format.
- Integer values outside the exact range of the current Double numeric Variant
  are explicitly rejected rather than silently rounded.
- Native designer forms use the same converters as imported form layouts and DAO.
  Read-only query controls are displayed read-only, rather than accepting edits
  that cannot be persisted.

This is not complete Variant/type support. Currency, Decimal, native Date Variants,
full 64-bit integer arithmetic, declared VBA variable enforcement, source width/
precision/defaults/lookup constraints, locale-specific formatting and complete
DML coercion remain open. Saved query expression types are not completely inferred.
Existing older SQLite projects are not automatically rewritten based on guesses;
new imports provide the source-type metadata and normalized Boolean values.
