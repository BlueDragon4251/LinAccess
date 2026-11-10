# Native primitive arithmetic result types

MS-VBAL 5.6.9.3 through 5.6.9.3.5 define effective operand/result types.
Native +, - and * now retain Byte, Integer or Long when the result fits the
combined effective type. Boolean contributes Integer; Empty combines as zero,
with Empty/Empty producing Integer. Single combined with Byte, Integer, Empty or
Single produces Single; combining Single with Long or Double produces Double.
Single floating division retains Single with these same compatible operands;
primitive integral division continues to produce Double.

The implementation applies after existing scalar-default resolution and checked
numeric calculation. It retains the prior fallback for results outside Byte,
Integer, Long or Single range. It does not claim correct declared-type-dependent
Variant overflow promotion: that remains an explicit F041 limitation. Existing
Currency/Decimal/Date paths, strings, Null and checked numeric exceptions remain.
Access Eval retains its separate existing arithmetic behavior. This is not a
claim of complete SQL type inference or Windows arithmetic equivalence.

The application, core, DAO and query model targets built successfully. One
existing bundled core/DAO/query-model run passed 3/3 in 8.91 seconds
(5.33/3.37/0.21). The core matrix covers real VarType and values, compatible/mixed
numeric operands, Boolean/Empty, fractional division, nested arithmetic, Null
and string concatenation. No new CTest IDs or catalog entries were added.
Fixed catalog totals remain 456/386/70; F041 stays open.
