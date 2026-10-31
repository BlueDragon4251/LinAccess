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
  which can incorrectly remove projected duplicates. Multitable DISTINCTROW
  remains unsupported without complete source-record identity mapping.
- TOP integers without ORDER BY use LIMIT. Ordered TOP uses a hidden RANK column
  so ties at the boundary are retained; that column is excluded from recordset
  fields. Explicit ORDER BY aliases are bound to their expressions. Ordered
  TOP DISTINCT, ordinal ordering, PERCENT and UNION combinations remain rejected.
  Ranked TOP queries are not yet installed as dependency views.
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
separate imported SQLite database. Current local run: 37 executed, 28 missing a
query dependency, 23 unsupported function/compilation cases, 8 other unsupported
Jet/binding cases, 2 execution-budget failures and 2 action queries unexecuted by
this corpus runner. Action query mutation is verified separately with typed
QueryDefs tests. No row values or raw SQL are emitted in corpus reports.

Execution without a Microsoft Access reference result is not proof of original
query semantics. The report says executed_no_access_reference_oracle explicitly.
Remaining dialect work includes '&' coercion, functions, nested parameter queries,
form/report references, saved ranked TOP dependencies, multitable DISTINCTROW,
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
joins and automatic form/control parameter resolution are not implemented.
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
The original query corpus still has 37 successful executions without an Access
reference oracle. Remaining expression differences include SQLite division by
zero producing NULL and exact Jet numeric/locale behavior; eager IIf alone does
not resolve those differences.
