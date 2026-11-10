# Native Count result types

A root SQL Count expression now has a proven Long (INTEGER4) result declaration.
This is result metadata for the actual aggregate, not an inference from a
particular returned integer. Count(*), Count(field) and Count(Null), grouped,
parameterized and empty-input queries keep their existing actual SQL values.
Empty input still returns one count row containing zero.

The metadata propagates through query preparation, QueryDef.Fields and real DAO
recordset Fields.Type4/Size4/Long Variant values. QueryDef metadata preparation
does not execute a project function or step source records. Arbitrary numeric
expressions and ambiguous raw integer columns retain their existing guards.
Complete aggregate promotion, SQL grammar, large-count/provider limits and
Windows/Access parity remain open in F068/F121/F122.

The existing access_query_model and dao_execution targets check independent
counts, Null versus total counts, grouped/parameter/empty results, actual DAO
metadata/subtypes and metadata-only project-state preservation. No new CTest or
catalog point is added.
