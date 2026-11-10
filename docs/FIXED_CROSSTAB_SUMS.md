# Fixed-heading Double and Currency sums

The existing fixed-heading TRANSFORM path accepts Sum over proven Double/REAL
or native Currency physical fields. Double inputs use checked field conversion
and normal SQL sums with Double output. Native Currency inputs use the existing
exact checked Currency accumulator for each pivot column, preserving scaled
integer arithmetic and Currency DAO metadata. It does not sum money through a
plain SQLite floating accumulator.

Conditional aggregation retains Null for absent/NULL-only cells. The existing
physical-table/alias, fixed-heading, parameter and source-shadow checks remain.
The source is resolved before execution and selects the existing native Currency
storage conversion and sum callback only for a proven Currency field.

This does not expand the supported physical MONEY range, implement Decimal sums,
or guess the result/promotion semantics of Integer/Single/Boolean/date sums.
Those types remain3251 in this path. Dynamic headings, First/Last pivot selection,
joined/saved-view sources and complete TRANSFORM semantics remain open in F068;
full exact original MONEY storage remains F014.

The existing currency_schema_queries case checks real Double sums, exact Currency
cancellation, NULL/absent cells and actual DAO Fields.Type/Value (7/Double and
5/Currency), together with all previous Count/Min/Max/statistical crosstabs and
an explicit unresolved Integer-sum guard. No new CTest or catalog point.
