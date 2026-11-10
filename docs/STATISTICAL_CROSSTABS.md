# Fixed-heading numeric statistical crosstabs

TRANSFORM now accepts Avg, Var, StDev, VarP and StDevP over a proven numeric
physical field with the existing fixed literal PIVOT headings and simple row
headings. Each fixed output column uses a real conditional aggregate over rows
matching that pivot value. Nonmatching/NULL input does not contribute to sample
or population counts. Missing groups remain Null; sample statistics for one
number remain Null, while population statistics return zero.

The input field is resolved from the actual unshadowed physical table, including
its actual alias. Real/Double, Single and Byte/Integer/Long storage types use the
existing checked native field conversion. Result metadata is Double in the query
plan and real DAO Fields.Type/Value. Currency, Decimal, text, Boolean/date and
unproven input fields retain explicit3251 boundaries for this statistical path;
existing typed Min/Max handling is preserved. No numeric type is inferred from
an arbitrary expression or saved view. Existing row-heading, fixed-heading,
source, parameter and SQL shape restrictions remain.

Existing currency_schema_queries checks all five functions on actual grouped
SQLite data, missing headings, NULL-only rows and singleton sample/population
differences; checks actual DAO output types and parameterized QueryDef.Type16;
and refuses text/Currency input and temporary source shadows. Existing
access_query_model covers the shared aggregate implementation and SQL metadata.
No new CTest or catalog ID is introduced. Dynamic headings, Sum precision and
full TRANSFORM grammar/collation/event/Access oracle remain open under F068.
