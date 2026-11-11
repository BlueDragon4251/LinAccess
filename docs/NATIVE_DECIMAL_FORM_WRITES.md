# Exact-fit physical Decimal form edits

Bound controls and unbound record-field assignments use the physical table's
proven Decimal precision/scale and actual storage constraint. Conversion happens
before changing a control or pending record value. Save, dirty comparison and
formatted pending-value searches use the same context-aware conversion. Query
results without a physical write context retain the previous refusal.

This shares the grid/DAO exact-fit contract: NULL and fitting Decimal values are
preserved without Double conversion; values requiring rounding refuse error6.
No regional text coercion, scale-rounding, arbitrary query updating, Decimal
field creation or MDB write-back parity is claimed. The fixed catalog remains
456 points, 386 confirmed partial scopes and70 open full requirements.

App/helper build succeeded. Existing import scope passed (0.24s); after correcting
the GUI fixture to use its real Controls/Value interface and a proper Section,
the affected existing GUI scope passed (18.62s). Exact stored value, rejected
scale change with retained clean value, and NULL Save are verified. No new CTest ID.
