# Direct declared numeric parameter result types

A directly selected declared Byte, Short, Long/Integer, Single, Double or
Boolean/YesNo parameter now carries its real normalized declaration into result
metadata. Ordinary aliases and complete enclosing parentheses are supported.
The existing Jet parameter alias normalization, checked binding and Currency/Date
paths remain authoritative; types are not guessed from the supplied value.

Actual query columns, QueryDef.Fields and DAO Fields.Type/Size/Variant values
retain Byte/Integer/Long/Single/Double/Boolean identity. A bound Null remains Null
while the declared column type remains available. The change annotates only a
direct parameter projection; it does not silently preserve a Byte type across
negation or infer arithmetic/function promotions. VALUE/text-width/general
expression/reference-library typing and full DAO/Access parity remain open.

The existing access_query_model and dao_execution cases jointly check normalized
TINYINT/SMALLINT/INTEGER/REAL/FLOAT8/BIT/YESNO declarations, parenthesized aliases,
actual DAO types/sizes/subtypes, Null and the negative Byte projection boundary.
No new CTest or catalog point. F068/F121/F122 remain open.
