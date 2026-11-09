# Native visual query and source designer

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
editor as one native text undo operation; only its explicit Run action executes the selected query type.
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

Scope includes physical tables or installed saved selection-query views, up to
eight ordered INNER/LEFT/RIGHT joins, fields/literal criteria/native supported functions,
AND/OR field groups, declared parameters, calculated scalar projections, grouped
aggregates, distinct and ordered output. General join syntax and arbitrary SQL round-trip parsing and the complete Access QBE/window/locale oracle remain open.
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

## Owned report sources

The native report editor exposes the same real visual query designer through
Datenquelle entwerfen. Applying replaces its SQL source as one text undo edit;
Cancel leaves it intact. The report stores optional visual_query metadata with
its exact SQL hash. Matching saved text restores the design; manual SQL changes
remove stale metadata on save. The existing native report layout/preview/PDF path
uses this generated SELECT, with its existing32-column and formatting bounds.

Report sources currently have no parameter prompting/binding, so this entry hides
the parameter declaration dialog and refuses persisted parameter declarations.
It does not save responses or guess values. Query editor parameter support is
unchanged. Previewing the design and saving the report definition prepare only;
explicit report/query execution reads rows. A project identity/generation guard
also covers report source editing, definition generation, saving and output.
Full report rendering/event/pagination parity and the Windows oracle remain open.

Validation: application built; existing GUI/SQL2/2 passed (15.31s). The GUI
fixture authors and saves a report source, verifies undo/redo and Cancel, reopens
its saved design and invalidates metadata after manual SQL. A callback counter
proves no values execute during design or save; explicit execution returns the
actual expected records. No new Windows/original/full-suite/package proof.

## Saved selection-query sources

The source list includes actual physical tables and installed read-only saved-query
views, excluding internal project/SQLite objects. Existing native dependency-view
installation runs before opening the designer; field discovery still prepares
without stepping data. A design can select, calculate, filter, group and sort the
output fields of an existing supported saved selection query. The metadata keeps
that source's real name. Saved-query dependency, type-lineage and parameter-view
boundaries are unchanged: sources that cannot be installed as native views are
not guessed or populated with fallback records. Arbitrary SQL reverse parsing,
a full QBE/Access oracle remains outside this scope.

Validation: existing GUI/SQL2/2 passed (15.84s), application built. An enum
conversion build error was fixed before running the bundle. Actual nested query
results/filter/order and saved design restoration are covered; callback counts
prove metadata preparation does not execute row values. No new CTest ID or
Windows/original/full-suite/package proof.

## Joined sources

Quellen verknüpfen builds up to eight ordered INNER/LEFT/RIGHT joins across real
tables or installed saved selection-query views. Each new source joins an already
present source through explicit field comparisons. LEFT preserves previous-side rows
with actual Null values on unmatched new-side fields. The dialog previews the
actual parenthesized FROM clause; no records execute there. Removing the final
join also removes its draft field rows; calculated expressions that still refer
to a removed source must be corrected before Apply. Cancel changes nothing.

Fields carry their source separately from their displayed name, preventing
ambiguous duplicate physical names from being resolved to the wrong source.
Joined field labels show source.field; duplicate output names require explicit
aliases. Criteria, ordering, calculations, parameters, grouping and report source
reuse go through the existing native SQL compiler and execution semantics.

The draft preserves the ordered joins and source/field identities on reopening.
Source aliases permit source reuse as described below, with at most255 combined
source fields. Composite and non-equality comparisons are now supported as described below;
OR-connected join conditions remain outside this visual entry.
No Access SQL reverse parser or complete QBE/window/locale parity is claimed.

Validation: existing GUI/SQL2/2 passed (14.92s); application built. After
strict metadata/partial-row cleanup guards only GUI1/1 repeated (14.90s). The
fixture compares two-level LEFT results including Null against independent SQL,
reopens joins, removes dependent sources, changes to INNER, and rejects malformed
metadata. No new CTest ID or Windows/original/full-suite/package proof.

## Owned form sources

The native form designer also exposes Datenquelle entwerfen. It applies the
actual generated SELECT to the editable source selector; available bound fields
come from that query's prepared columns. Query forms retain their existing
read-only behavior and explicit non-parameter source boundary. This does not
claim editable joined/aggregate queries. Parameter declarations are hidden and
refused as for report sources. Physical/unbound form creation remains available.

The form persists optional visual_query metadata only for a matching exact source
hash. Manual source changes remove stale design metadata on save. Its existing
form-level undo/redo now snapshots and restores the source design together with
selected fields, layout, events and other authoring state. Canceling the nested
designer leaves the source unchanged. Project identity/generation is checked for
source authoring, saving and history capture/restoration. The source selector has
a bounded preferred width so long SQL does not force an enormous designer width.

Validation: application built, existing GUI/SQL2/2 passed (15.47s). The fixture
verifies source/metadata undo/redo, saving/reopening/manual invalidation and opens
the actual generated query form with a real bound value and AllowEdits=False.
The callback counter distinguishes schema preparation from execution. No new
CTest ID, Windows/original/full-suite proof. The checked DEB at fe01e1d predates
this form integration; its Build-ID/payload/dependency checks and all seven
fixture/upload original hashes passed.

## Source aliases and self joins

Each newly joined source can now have an explicit alias. Reusing the same physical
table/view requires a distinct alias; all source identities are unique under the
existing native case folding. The main source retains its real name. Joins refer
to source identities, whereas metadata discovery always reads the underlying
actual table/view. Aliases are quoted, limited to64 characters and persisted with
the ordered join design. The left-side selector includes existing aliases and
shows their physical origin as a tooltip. The totals dialog qualifies source
identities too, so duplicate field names remain distinguishable.

This permits actual parent/child self joins and further joins from an alias,
including LEFT Null rows. Draft fields retain alias identity across reordering,
criteria, calculations, grouping, saving/restoring and removing a joined source.
Duplicate output labels still require output aliases independently of source
aliases. Eight joins and255 combined source fields remain the bounded scope;
multiple AND-connected comparison predicates are described below; OR-connected
join conditions remain outside this visual entry. No full Access designer or SQL reverse parsing parity is claimed.

Validation: application built; existing GUI/SQL2/2 passed (16.45s). An actual
parent/child LEFT self join matches independent SQL including unmatched Nulls.
An alias containing spaces persists/restores; alias metadata fields resolve from
the underlying source. Ambiguous aliasless reuse is refused without adding the
join. Preview does not execute rows. No new CTest ID or new original/Windows/
full-suite/package proof.


## Composite and directional joins

INNER/LEFT/RIGHT joins now accept1..16 AND-connected field comparisons with =,
<>, <, <=, > or >=. Comparisons can refer to any previously joined source/alias;
the other field always belongs to the newly joined actual table/view. RIGHT
preserves the new-side rows, with actual Null fields on unmatched earlier sources.
Native SQL type/coercion/updateability boundaries remain unchanged.

Vergleich vormerken retains a comparison and clears the current field pair for
another. The staged list and any complete current pair form the ON condition.
Partial pairs, invalid operators, missing fields, duplicate staging and excess
comparisons are refused. Changing the new physical source clears staged fields;
Neue Verknüpfung verwerfen discards only the unadded draft. OK refuses an
unadded source/condition draft. The existing FROM preview contains exactly the
added joins. Buttons share a compact row and labels wrap within the dialog.

The composite-join stage introduced version2; the current action stage writes
version3 and still reads versions1/2, including the
legacy single-equality join representation. Older designers that support only
version1 will not silently reinterpret version2 composite join metadata. Manual
SQL and its exact hash remain the authoritative stored query/source.

Validation: application built; existing GUI/SQL2/2 passed (15.54s), without a
follow-up rerun. Actual native executeAccessQueryPlan RIGHT results for compound
keys match an independent reversed LEFT query including Nulls. Restored INNER
metadata with <= also matches independent SQL. Earlier version1 designs, aliases,
query/form/report persistence and preview non-execution remain in the same target.
No new CTest ID or new original/Windows/full-suite/package proof.

## Native action-query authoring

Abfrageart offers selection, make-table, append, update and delete drafts. The
query/source SQL remains authoritative; applying the draft changes only editor
text. Metadata preparation never executes an action or reads its records. The
existing SQL editor Run uses the actual native DAO QueryDef/parameter/execution
path. No alternative DML or record-copy engine is introduced.

Make-table emits SELECT ... INTO with an explicit unused project table name;
existing targets are refused rather than overwritten. Append emits INSERT INTO
with an explicit one-to-one mapping from actual prepared output columns to real
existing target fields. Duplicate/missing mappings are refused. Omitted target
fields retain the native default/required/type/AutoNumber rules at execution.
Both derive from the current real selection, including its supported parameters,
criteria, calculations, grouping and joins. Native SELECT INTO metadata/precision
propagation boundaries remain those of the existing DAO engine.

Update/delete require a single physical project source without joins, totals,
DISTINCT or sorting. Update explicitly selects fields and bounded scalar SET
expressions (up to4096 characters), including declared parameter references.
Delete uses the actual criteria as WHERE. Both change the selected physical
source only on explicit Run; no affected-row count is invented during preview.
The source/output selection must remain valid for schema preparation. Joined or
saved-query updates and more general expressions retain their existing manual
SQL/DAO support boundaries; this is not full Access updateability.

The action mode, target/mapping or update expressions persist under the exact-SQL
hash. The action stage introduced version3; the current TOP stage writes version4
and reads versions1/2/3/4. Older designers
will not silently reinterpret the new action format. Form/report source entries
remain selection-only and hide the action chooser as well as parameter declarations.
Saving and reopening action designs does not execute them. SQL text undo keeps
its existing behavior; complete visual/SQL reverse parsing remains outside scope.


Validation: the initial make-table/append GUI/SQL bundle passed2/2 (16.23s).
Their real source/target schema and records were unchanged during design/Apply;
actual editor Run created/copied the expected records and target defaults. Saved
modes, target and mappings restore without execution; duplicate mapping is refused.
After extending update/delete, only the affected GUI target was run. A Qt JSON
iterator build error was corrected. The failure initially attributed to update
was localized to delete preparation: DELETE * FROM was changed to native
DELETE FROM. Final GUI1/1 passed (15.79s), application built. The same fixture
executes a declared Long update parameter, checks only matching rows change,
reopens its expression metadata and executes a criterion-limited delete. These
are existing targets, not new CTest IDs. No new core/original/Windows/full-suite
or package proof is attributed to the actions.


## TOP selection

The TOP control selects a positive signed32-bit integer, or Alle Datensätze (0)
for no TOP clause. Ordered queries retain the native RANK tie semantics: TOP2 can
return three actual rows when the second value is tied. Unordered TOP retains the
existing native LIMIT behavior. Hidden rank columns are not user output fields.
Raw __la_top_ helper views are excluded from source/join/target choices; the
public saved query remains a normal selectable source.

TOP is currently selection-only in this visual entry; action mode cannot silently
ignore the limit. Action selection is disabled while TOP is set, and non-selection
metadata with TOP is refused. PERCENT, ordered DISTINCT and other existing native
TOP grammar/type/budget boundaries remain explicitly unsupported here. No larger
result materialization or Access oracle scope is claimed.

Saved designs write version4 with the exact integer top value; versions1/2/3/4
are read. Fractional, negative or out-of-range persisted values are refused rather
than silently clamped. Manual SQL changes retain the exact-hash invalidation rule.

Validation: application built; only existing affected GUI1/1 passed (16.14s).
Real editor Run includes boundary ties and removes the hidden column, compared
with independent SQL using the nth sort-value boundary. Saving/reopening, clearing
the limit to return all rows, helper-view exclusion, invalid metadata and callback
non-execution are covered. No new CTest ID or additional core/original/Windows/
full-suite/package proof.
