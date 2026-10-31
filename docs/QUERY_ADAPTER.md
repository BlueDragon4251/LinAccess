# Stored Access queries and DAO QueryDefs

This is an incremental Jet ANSI-89 adapter, not complete Jet/ACE SQL support.
Original query source bytes remain unchanged in the Bridge artifact catalog.

Implemented with execution tests:

- Load imported named queries via CurrentDb.QueryDefs(name). Read SQL, set typed
  named/indexed Parameters(...).Value, OpenRecordset and Execute action queries.
- Temporary CreateQueryDef("", sql), SQL replacement on that object and
  RecordsAffected. Unset parameters fail with error 3061; named persistent
  CreateQueryDef and persistence of SQL edits remain unsupported.
- PARAMETERS declarations, case-insensitive names, bound values, integer ranges,
  floating/text/Boolean and canonical date-time parameters. Basic ANSI-89 LIKE
  '*'/'?' patterns preserve literal SQL '%'/'_' characters. Character classes
  remain explicitly unsupported.
- ANSI-89 quoted strings, bracket identifiers and True=-1/False=0.
- Single-table DISTINCTROW is ignored, as in Jet. It must not become DISTINCT,
  which can incorrectly remove projected duplicates. DISTINCTROW is also ignored
  for a bare all-source wildcard. Other multitable projections still require
  source-record identity mapping; see [derived tables and origins](QUERY_ORIGINS.md).
- TOP integers without ORDER BY use LIMIT. Ordered TOP uses a hidden RANK column
  so ties at the boundary are retained; that column is excluded from recordset
  fields. Explicit ORDER BY aliases are bound to their expressions. Ordered
  TOP DISTINCT, ordinal ordering, PERCENT and UNION combinations remain rejected.
  Ranked TOP dependencies now use owned temporary helper/public views; hidden
  rank fields stay internal, ties remain included and dependency chains retry.
- Compatible parameterless SELECT artifacts become temporary SQLite views for
  dependency resolution and form record sources. Refresh updates owned views,
  preserves user tables/views and rolls back DDL if active readers block it.
- Date literals with explicit four-digit years, Date()/Now()/Time(), and canonical
  timestamps. MDBTools exports both date formats as yyyy-MM-dd HH:mm:ss so dates
  are not compared against locale-dependent CSV strings. Subsecond Jet precision,
  regional calendars and exact Date Variant semantics remain open.
- SQL execution is bounded using the native distro Qt SQLite handle and the
  loaded SQLite progress-handler API (no SQLite implementation is embedded).
  Interrupted fetches fail instead of returning partial recordsets. A recursive
  query and a query interrupted after its first row exercise both paths.

## Original corpus measurement

All 100 original Kassa query artifacts are hash-checked and classified against a
separate imported SQLite database. Current local run: 72 executed without extra
parameters and five with parameters selected from actual data (77 total), nine
missing query dependencies, eleven other Jet/binding failures, one source-column
mismatch and two actions unexecuted by this runner. Parameterized results are
compared against positive independent filter counts. Action mutation is verified
separately with typed QueryDefs tests. Reports contain hashes/status/error codes,
not values or raw SQL. See QUERY_ORIGINS.md for the remaining six-column mismatch.

Execution without a Microsoft Access reference result is not proof of original
query semantics. The report says executed_no_access_reference_oracle explicitly.
Remaining dialect work includes full '&' coercion details, functions, nested parameter queries,
arbitrary form/report property chains, multitable DISTINCTROW,
crosstabs, implicit form parameter resolution, collations, precise types and editable query joins.

## Application execution path

Imported query objects now run through this adapter from the navigator. Explicit
PARAMETERS declarations prompt for values and pass through the same type checks
as DAO. SELECT results use AccessQueryModel; actions use QueryDef.Execute against
the project copy without additional per-statement consent. The native SQLite SQL
editor retains its separate SQL behavior.

DAO SELECTs, the result model and saved-query form sources share bounded result
materialization, hidden TOP-rank removal, tie handling and interrupted-fetch
rejection. Result views and query-bound forms remain read-only; editable query
joins and undeclared parameter inference are not implemented. Live explicit
Forms/Formulare and Reports/Berichte references and typed Value parameters are
implemented with native control lookup; see [SQL object values](SQL_OBJECT_VALUES.md).
Existing native SQLite views remain usable as read-only form sources.

The access_query_model test checks TOP ties, headers, typed parameters, NULL,
read-only flags, rejection of DML without mutation, clearing stale results after
errors and recovery. The GUI event suite exercises a named saved TOP query as a
real form record source. The original 100-query corpus continues to run through
DAO and the shared executor; no Access reference oracle is available.

## IIf argument evaluation

Jet IIf calls are rewritten to an internal native SQLite scalar function. SQLite
therefore evaluates both argument expressions before the callback chooses the
result, instead of using SQLite's built-in short-circuit IIf. The callback retains
the selected SQLite value type and NULL, treats NULL conditions as false, and
rejects invalid condition coercions. Registration probes the live connection so
close/reopen does not retain stale native-handle state.

The model regression contrasts the SQLite built-in with the Jet adapter using
an integer-overflow error in an unselected argument; the Jet path must fail.
Numeric/text/NULL results, condition coercion and connection reopen are covered.
Current original-query coverage is measured above. Remaining expression differences include SQLite division by
zero producing NULL and exact Jet numeric/locale behavior; eager IIf alone does
not resolve those differences.

## First and Last aggregates

First(expr) and Last(expr) use native SQLite aggregate callbacks with one copied
value per group. They retain NULL at the selected end, unlike Min/Max, and preserve
the selected SQLite storage type. Empty inputs return NULL. Finalization releases
copied SQLite values. The regression covers separate groups, NULL endpoints,
empty input, exact SQLite 64-bit integers, real values, embedded-NUL BLOB values
and connection reopening. This BLOB test concerns query results; it does not
implement MDB OLE import or a VBA BLOB Variant.

Selection follows the SQLite input scan order. Access also does not promise a
particular First/Last result without an established record order; SQLite and Jet
can choose different scans. No cross-engine equality claim is made. Current corpus status is measured above; it remains without an Access oracle.

Date Format/Format$ is described in DATE_FORMAT.md. Native concatenation and
current original query/report measurements are described in SQL_CONCATENATION.md.
Explicit localized form references use native control lookup (SQL_OBJECT_VALUES.md);
implicit parameter inference and full property chains remain incomplete.

## Schema labels and bounded query plans

Direct qualified projections receive explicit field labels. This prevents SQLite
nested RIGHT JOIN column suffixes such as id:1 from breaking original query
dependencies. Source aliases remain intact. Duplicate/wildcard projection naming
is not complete Jet provenance support.

TOP 1 can use a NULL-safe boundary query for a single column sort and a direct
projection over a proven stable table/view dependency graph. Parameters, OR
conditions, NULL ties, descending order and aliases have tests. Volatile/unknown
functions and complex cases keep the rank plan. A 20,000-row fixture verifies the
indexed path inside a 10,000-operation outer budget. Simple two-table RIGHT JOIN
queries can swap to LEFT JOIN when both sources are normal tables, projections
are explicit and the ON clause is a qualified equality. Independent native
RIGHT JOIN results verify duplicates and NULL extension. Wildcards, grouped and
complex joins are excluded from this optimization.

Ranked view schema discovery does not execute value expressions (tested with a
counting resolver). Helper names are owned separately, collisions are preserved,
missing dependencies retry and user columns named like the rank field remain
visible. DAO recordsets retain the complete shared query plan. Nested execution
budgets charge all enclosing budgets and restore the parent progress handler.

Bracketed derived tables, proven column lineage across saved queries and bounded
pure alias expansion are described in [QUERY_ORIGINS.md](QUERY_ORIGINS.md).

Pure IIf expressions can feed later SELECT aliases, provided all nested functions
are known pure and no actual source column shadows the alias. Three additional
original queries must return positive, non-NULL results; synthetic regressions
check independent tax totals, NULL cases and prevention of volatile callback
duplication. Leading-dot decimal/exponent literals are also tokenized intact.
