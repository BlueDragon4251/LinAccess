# Native Byte Variant arithmetic widening

Microsoft's [+ operator](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/plus-operator),
[- operator](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/minus-operator)
and [* operator](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/operator)
explicitly describe overflow of Byte-containing Variants widening to Integer,
and Integer-containing Variants widening to Long. The proven native declaration
binding now applies this missing Byte step. Addition/subtraction and Byte products
that fit Integer produce Integer; larger supported Byte products produce Long.
The existing fixed-Byte error6 and unbound/Access Eval boundaries remain.

This uses actual Variant declaration provenance and Byte subtypes, not guessed
runtime values. Subsequent arithmetic retains the Variant category. Native
CVar/Nz and real local Variant variables exercise the same path. No physical
column widening or general SQL dynamic-result type inference is introduced.
Full F041/F122 and the fixed456/386/70 catalog totals remain open/unchanged.

Application/core/DAO targets built. First existing bundle passed2/2 in8.68s
(core5.50/DAO3.18). It checks signed underflow, Byte addition, Integer/Long product
boundaries, preserved fixed errors and a real native Integer-returning VBA
function through SQL/DAO Type3/Size2 with actual Integer value256. No new CTest
or catalog IDs. The checked4cd0f848 DEB predates these later arithmetic scopes.
