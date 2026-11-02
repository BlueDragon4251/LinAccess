# Declared SQL parameter types and DAO metadata

Access SQL INTEGER is a four-byte Long, distinct from VBA Integer (two bytes).
The parameter compiler now canonicalizes INTEGER/INT/INTEGER4/LONG to Long,
SMALLINT/SHORT/INTEGER2 to Short, and TINYINT/BYTE/INTEGER1 to Byte. The former
16-bit INTEGER binding limit is corrected without changing VBA Integer storage.
The earlier explicit 16-bit DAO regression now declares SHORT.

REAL/SINGLE/FLOAT4/IEEESINGLE canonicalize to Single; FLOAT/DOUBLE/FLOAT8/
IEEEDOUBLE/NUMBER to Double. Single parameter binding rounds to 32-bit float,
checks overflow before execution and rejects unvalidated below-minimum underflow
with 3251. Floating bindings reject nonfinite values; NULL remains NULL.

DAO Parameter.Type reads the declared Boolean/Byte/Short/Long/Single/Double/
Date/Text type as an Integer containing the documented DataTypeEnum code. Common
unqualified DAO data-type constants are available in VBA. Type is read-only for
an Access workspace; native writes reject with 383 and preserve SQL/value/type.
Closed or replaced QueryDefs invalidate Type access with 3420; wrong getter arity
raises 450. Internal generic/inferred types reject with 3251 rather than guessing
a type. Exact Windows automation error-description/timing identity remains open.

Microsoft's Parameter.Type reference states: “For a Parameter object in a
Microsoft Access workspace the property is read-only.” ODBCDirect mutability is
not added. Parameter.Size/width validation, assignment-time coercion, complete
inference, SQL Currency/Decimal and all DAO/VBA type/locale contracts remain open.

Synthetic DAO tests execute real parameterized UPDATEs across integer and floating
aliases, verify bounds/rounding/NULL and unchanged values after rejection, inspect
Boolean/date/text metadata, and exercise read-only/replacement/close behavior.
Native VBA branches on Parameter.Type=dbLong and writes 40000 through an INTEGER
parameter. Original application execution coverage is unchanged. Validation is
recorded in PROGRESS.md.

References:
- https://learn.microsoft.com/en-us/office/client-developer/access/desktop-database-reference/sql-data-types
- https://learn.microsoft.com/en-us/office/client-developer/access/desktop-database-reference/equivalent-ansi-sql-data-types
- https://learn.microsoft.com/en-us/office/client-developer/access/desktop-database-reference/parameter-type-property-dao
- https://learn.microsoft.com/en-us/office/client-developer/access/desktop-database-reference/datatypeenum-enumeration-dao
