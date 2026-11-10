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


Proven native argument binding (2026-11-10)
-----------------------------------------
The six functions additionally preserve proven Date, Boolean, Single, Currency
and integral/Double source declarations through the shared native scalar adapter.
A Date input is its actual OLE Date Variant; Boolean True retains -1 despite a
physical1, Single values are rounded before conversion, Currency keeps the
existing bounded exact conversion. ISO text in a TEXT source is not inferred to
be a Date. Nested conversions carry their declared result type into enclosing
functions and real DAO Fields.Type/Value metadata. Supported simple filters,
grouping and ordering use the same typed calls. Source/locale/complex-binding
limits remain; this does not establish full coercion or Windows parity.

The original raw callbacks remain available. Their unproven input transport and
large-integer limits are not changed by this extension. NULL94, overflow6 and
invalid-type13 remain checked by the native coercion implementation. The earlier
paragraph describing an exclusively raw scalar transport is superseded only for
these proven typed expressions. CDec, full Decimal storage, reference libraries
and all other open F068/F071/F105/F122 scopes remain open.

The existing explicit Boolean-to-Byte coercion is preserved: CByte(True) is255,
while CInt(True), CLng(True), CSng(True), CDbl(True) use -1. A physical BOOLEAN1
therefore must reach the native Boolean conversion rather than raw integer1.
