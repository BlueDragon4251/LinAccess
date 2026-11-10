# Native SQL numeric conversion functions

CBool, CByte, CInt, CLng, CSng and CDbl now execute as real registered SQLite
scalar callbacks using the existing native VBA conversion evaluator. Conversion
rules are shared with VBA, including integer banker rounding, declared ranges,
Single rounding, Boolean -1/0 and native conversion errors. This adds no host
object, COM dependency, project mutation or fabricated data.

Root conversion expressions carry the actual Boolean/Byte/Integer/Long/Single/
Double result metadata through query preparation and DAO Fields.Type/Size/Value.
Numeric/text-number input, nesting, bound parameters and WHERE predicates use
actual SQL execution. Invalid arity is450, Null94, nonnumeric/binary input13 and
range overflow6 under the existing native evaluator's bounded coercion rules.

This preserves the current SQLite scalar transport boundary: it does not restore
an arbitrary date/object/Decimal Variant type from raw SQL text/numbers, implement
CStr's Boolean/date type propagation, add CDec storage, or establish full regional
or Windows expression/type precedence equivalence. Existing safe domain/ordering
function lists recognize these pure callbacks; no side-effectful function is
added to those lists. Full F068/F071/F105/F122 remain open.

The existing access_query_model and dao_execution targets jointly check all six
functions, exact independent rounding/range expectations, raw Boolean -1/0,
nesting, typed parameters, actual filtering and precise errors, plus real DAO
type/size/Variant subtypes. No new test target or catalog ID is introduced.
