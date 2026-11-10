# Native Currency with numeric operands

Native Currency + and - now use the checked scaled-unit Currency kernel with
Byte/Integer/Long/Single/Double/Boolean/Empty operands in either order. Currency
* integral/Boolean/Empty uses the same exact kernel. Operand selection uses real
value kinds and numeric subtypes, never a numeric-value guess. Decimal's prior
path and Date/Null handling keep their existing precedence and guards.

The actual integer partner is converted through the existing Currency converter;
all supported32-bit integer inputs are exactly representable at four-place
scale. Addition/subtraction/multiplication use the existing int128 intermediates
and signed64 overflow checks. No generic Double intermediate replaces a full
Currency value. Scalar defaults are resolved once before this path.

The addition/subtraction precision order is documented by Microsoft's
[+ operator](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/plus-operator)
and [- operator](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/minus-operator),
as well as MS-VBAL5.6.9.3. Their Currency precedence over Single/Double justifies
routing those + and - operands through the existing checked Currency converter.
Multiplication has a different documented precision order. Mixed Double products
retain Double; full Single/product promotion remains unvalidated, including the
difference between the general multiplication documentation and MS-VBAL.
String/locale and complete other mixed promotions remain open.

This changes native language arithmetic, including real native VBA functions
called from SQL. It does not add general raw SQL mixed-Currency translation,
change the physical CURRENCY_REAL lossless range or provide full Jet MONEY/
Decimal/MDB round trips. F014/F041/F068/F105 remain open; fixed totals456/386/70.
Access Eval's separate observed path is unchanged.

Application/core/DAO targets built. The final existing bundle passed2/2 in8.64s
(core5.37/DAO3.27), covering both orders, four-place ticks,32-bit integer limits,
full signed Currency boundaries, Boolean/Empty, checked overflow, Double/Single
addition/subtraction, Double-product distinction and exactly one scalar-default
read. A real native VBA function verifies Currency type inside its body and
returns2.0003 through SQL/DAO Type5/Size8 with actual Currency units20003.
No new CTest targets/catalog IDs; no fresh Windows/full-suite claim.
The latest checked DEB4cd0f848 predates logic, division and this Currency scope.
