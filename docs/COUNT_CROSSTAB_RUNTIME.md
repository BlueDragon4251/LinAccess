# Fixed-heading native Count crosstabs

The Jet query compiler accepts bounded TRANSFORM Count(field) / Count(*) queries
with simple grouped row fields and a fixed PIVOT field IN (...) list. Headings
can be text or signed-Long integer literals; integer labels are canonical decimal
names, so 01 and 1 collide. Year(field), Month(field) and Day(field) are supported
as pure calendar pivot expressions through the existing native date functions.
It rewrites tokens into conditional counts, preserving heading order, literal
escaping, explicit aliases, WHERE/HAVING/ORDER clauses and declared parameters.
Count(field) omits NULL values; Count(*) counts matching rows. Absent categories
produce zero counts. Names are checked for case-insensitive duplicates and row
heading collisions; the output is limited to 255 columns and 64-character pivot
headings. Count outputs have proven native Integer4/DAO Long type, including an
empty result. QueryDefs classify as dbQCrosstab (16).

Direct DAO recordset source text and saved native QueryDefs use the same query
compiler. Aggregate results remain readonly. Simple source-column lineage uses
the existing SQLite origin metadata path. Native execution budgets, parameter
binding and result bounds remain in effect.

Dynamic headings, nonintegral/out-of-Long headings, Sum/Avg and other aggregates, expression row
headings, compound/join sources and arbitrary TRANSFORM grammar refuse3251.
This implements native fixed-heading counting; no new Windows Access result,
collation, regional heading or full crosstab compatibility oracle is claimed.
F068/F067 remain open; no fixed-catalog point is added.

## Typed fixed-heading Min/Max

The same structural grammar supports Min(field) and Max(field) on an unshadowed
physical main table, including a table alias. Prepare resolves the declared input
type: Currency, Single, Byte/Integer/Long, Boolean and Date/DateTime are supported.
Existing native typed storage functions validate/normalize participating values
before the aggregate compares them. No sum or average is computed. Currency
preserves the existing four-decimal REAL storage boundary; physical Decimal and
64-bit/unspecified SQLite INTEGER are not narrowed or advertised as Jet Long.

Each pivot result preserves the proven input type, including empty result sets;
DAO QueryDef.Type stays16. NULL and absent categories produce NULL, not zero.
Direct queries, parameterized QueryDefs and saved native QueryDefs share this path.
Query view installation validates extrema sources, and the main schema version is
part of the view revision so changed source declarations invalidate old expressions.
Text/locale extrema, unsupported input types, saved-view source lineage and temp
shadows refuse3251 during prepare. QueryDef creation may retain source text for a
query whose source cannot currently execute. Arbitrary outer SELECT projections
through a saved view do not gain new guaranteed aggregate-type lineage here.

The existing currency_schema_queries target checks Currency/typed DAO values,
integer/date extrema, NULL/missing/empty results, parameters, saved definitions,
unsupported types and source shadowing. No new CTest target or Windows oracle.
F067/F068 remain open; catalog stays456/386 confirmed partial scopes/70 open full
requirements.

Temporary table definitions also participate in view invalidation. Adding a temp
shadow removes the formerly validated extrema view; dropping the shadow restores
it from its actual saved definition. The existing schema test checks both paths.
