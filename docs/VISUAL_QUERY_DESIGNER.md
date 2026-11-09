# Native visual selection-query designer

The existing SQL query editor now has an Entwurf action. Its actual Qt dialog
lists real project tables and fields from schema metadata. Output checkboxes,
aliases, field order, ascending/descending sort, per-field criteria and DISTINCT
compose a real native SELECT. Field up/down controls set projection/sort priority.
Hidden fields can filter/sort; DISTINCT requires sorted fields to be visible.
Criteria are predicate tails such as >=4, Like "A*", Is Null, Between 1 And 3 or an
IN literal list. Explicit numeric/string/date/Boolean literals can omit =. Native
predicate validation rejects extra SQL clauses/subqueries. The same existing SQL
compiler and execution path provide expression semantics; no alternate data engine
or invented records are introduced.

The preview prepares columns without stepping records or executing SQL/VBA value
callbacks. It never runs an action query. Applying transfers SQL to the existing
editor as one native text undo operation; only its explicit Run action reads rows.
Cancel leaves editor text unchanged. Invalid/empty selections, duplicate output
names, malformed criteria or source changes disable Apply. Reload explicitly
rebuilds from the current source. No table is fabricated or chosen when no source
has been selected.

Saving an owned query records optional visual_query metadata with a SHA256 of
its exact SQL text and the design. It is reused only for matching text; manual SQL
changes remove stale visual metadata on save. Existing unknown/advanced SQL is
kept in the SQL editor; the visual dialog does not guess its structure. Applying a
new design is an explicit replacement. Project/generation entry guards reject design,
run and save for a closed/replaced project; full callback/window scheduling remains
outside this scope. Imported sources remain outside this
native-owned authoring path.

Scope is one physical table, fields/literal criteria/native supported functions,
AND/OR field groups, declared parameters, calculated scalar projections, grouped
aggregates, distinct and ordered output. Joins and arbitrary SQL round-trip parsing and the complete Access QBE/window/locale oracle remain open.
The native 255-field/64-character name and criteria/predicate bounds apply. The
SQL editor retains its broader existing capabilities. No full designer parity is
claimed; F067/F068/F080/F148 remain open, fixed catalog456/386/70 unchanged.

The existing document target drives the real nested dialogs: actual schema fields,
ordering/alias/filter/sort, disabled invalid input, Cancel, actual query execution
compared with independent SQL, save/restore and stale-metadata removal. An actual
SQLite callback counter proves preview/apply/restore do not execute row expressions;
Run then evaluates the real records. Technical fixture objects are cleaned up.
There is no new CTest ID, production demo data or Windows reference output.


Validation: application built. The existing SQL model target passed in the initial
bundle; the new GUI fixture caught a QJsonValueRef copy error during row movement.
Using a real value snapshot fixed that defect. Only the affected GUI target was
repeated afterwards:1/1 passed (14.73s). No new core/original/Windows/full-suite
or package proof is attributed to this frontend change.


## OR groups and declared parameters

Up to8 criteria groups are editable as Kriterium/ODER columns. Fields within a
nonempty group combine with AND; nonempty groups combine with OR. Empty groups
are ignored. Moving a field retains all its criteria. Add/remove group actions
change only the draft; existing SQL changes only on Apply. Old saved designs
without group metadata retain their original single criterion group.

The parameter dialog declares up to64 unique, quoted names with native Text,
Long, Integer, Byte, Single, Double, Currency, DateTime or YesNo types. It produces
an actual PARAMETERS header; criteria use bracketed parameter names. No parameter
response is persisted in visual metadata. The existing query Run/DAO binding path
prompts and converts values. Type/range/locale/Currency boundaries are those of
the existing native binder. Preview/Apply/restore remain schema-only, including
parameterized queries. Saving preserves group/parameter design metadata under
the existing exact-SQL hash rule.

The existing GUI fixture drives both parameter declarations and actual successive
input dialogs, compares the OR query against independent SQL and reopens the design
with its groups/parameters intact. Cancel and later manual SQL invalidation remain
covered. The existing GUI/SQL bundle passed2/2 (14.51s) without follow-up reruns.
No new CTest ID, core/original/Windows/full-suite or new package proof is claimed.
F067/F068/F080/F148 remain open; the fixed456 catalog is unchanged.

## Calculated fields

The draft now adds and removes actual calculated output rows. Each expression
requires an alias and supports the same output, sorting and AND/OR criteria as
physical fields. The emitted SQL contains the expression itself (parenthesized),
not a fabricated stored column. Expressions and row order survive saving and
reopening under the existing exact-SQL hash rule. Removing a row changes only the
draft, never a table column or its data. The255-row and4096-character limits apply.

Expressions use the native bounded scalar expression compiler: arithmetic, field
references and supported scalar functions. Aggregates, subqueries, unknown VBA/SQL
callbacks and live object/file/clock values are refused in this designer scope;
manual SQL remains available with its existing explicit runtime boundaries.
Repeating expressions for criteria/sorting must not introduce external side effects.
Preview/Apply/restoration remain metadata-only. This scalar-field stage adds no joins; grouping is described below.

Validation: existing GUI/SQL bundle2/2 passed (15.09s), application built.
The same GUI fixture compares calculated values and sorting against independent
SQL, reopens the saved expression and refuses a callback without invoking it.
No new CTest ID, Windows/original/full-suite or package proof is claimed.

## Grouping and aggregates

The Gruppierung dialog enables a totals query and assigns each physical or
calculated row one of Group By, Where, Count, Sum, Avg, Min, Max, Var or StDev.
Group By expressions form GROUP BY; aggregate rows project the actual aggregate
and their criteria form HAVING. Where rows must be hidden and unsorted and filter
input records through WHERE. Each clause combines its own nonempty criteria groups
with OR, and fields within each group with AND. WHERE and HAVING apply together;
this is not a single OR spanning input and aggregate predicates. Both clauses and
GROUP BY are visible in the generated SQL before Apply.

Sorting applies to the group expression or aggregate. Alias, parameters, row order
and grouping choices persist in the same visual metadata; disabling totals keeps
the choices in the draft. Native SQL/DAO aggregate type, Currency and unsupported
expression boundaries still apply. There is no alternate aggregation engine,
record calculation in the preview, new CTest ID or claimed Windows oracle result.

Validation: GUI/SQL2/2 passed (14.62s), application built. Schema extension
review found that newly hidden columns must not change an existing GROUP BY;
they now default to Where with no criterion. Only affected GUI1/1 was repeated
(14.36s), including reopening after an actual added/updated technical column.
Results of grouping, Count, Sum, WHERE and HAVING match independent SQL.
