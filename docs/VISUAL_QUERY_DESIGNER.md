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
AND/OR field groups, declared parameters, distinct and ordered output. Joins,
grouped totals, calculated projection authoring, arbitrary SQL
round-trip parsing and the complete Access QBE/window/locale oracle remain open.
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
