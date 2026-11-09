# Native document filters

Table and saved-query datasheets have a real Filter dialog with Apply, Show all
records and Cancel. The applied predicate is scoped to the open document. The
same controller handles DoCmd.ApplyFilter and ShowAllRecords from native VBA and
macros. It does not save a runtime filter into the original MDB or imported
artifact.

ApplyFilter accepts optional FilterName and WhereCondition slots, including named
WhereCondition and omitted first slots. This scope implements explicit predicates;
nonempty saved filter names refuse3251. Null arguments refuse94 and a missing
condition refuses5. VBA predicates are limited to32768 UTF-16 units; native macro
WhereCondition arguments use the documented256-character bound. Additional
clauses, subqueries, NUL and malformed Unicode are refused. The existing Jet
adapter compiles literals, LIKE and supported functions.

Editable/readonly table views filter their actual SQL table models. Editable
saved queries retain their proven original WHERE condition when a user filter is
applied or cleared. Active editors finish into the existing cache; unsaved drafts
refuse2501 before selecting a new result. There is no implicit SQL save or cache
discard. Failed SQL filtering restores the previous filter and current cell.

Readonly saved-query documents retain their actual source and supplied parameter
values. Filtering constructs a replacement native query model through the shared
query-plan filter compiler, then swaps it only after success and owner/project/
page checks. Failure retains the original result and selection. Clearing the
filter restores the source query result. Requery retains the applied predicate;
its existing parameter prompting behavior is unchanged. Filtering can reevaluate
source expressions; it is not an in-memory snapshot predicate.

Forms use their existing atomic Filter/FilterOn controller and record validation.
Reports accept ApplyFilter only in their actual bound Open event, before loading
records. ShowAllRecords remains a table/query/form operation. Event context is
scoped and restored after nested bindings. Project/query/output reentry, modal
contexts and closed/changed owners are checked, with no fabricated active form.

Saved named filters, navigation-form ControlName, sorting through saved filters,
persistence through object saving, all server filter rules, complete Access UI/
event/locale semantics and a Windows Access result oracle remain open. Existing
F080/F056/F052/F088/F087 requirements and the fixed456-point total are unchanged.

Existing GUI1/1 (8.62s) and layout1/1 (11.77s) pass. The GUI check exercises
actual table/cache/SQL values, LIKE/date predicates, dirty refusal, failed-filter
restoration, real dialog apply/clear/cancel, editable base filters, readonly
replacement and parameter reuse, both macro hosts, normal form validation and
named WhereCondition during an actual report Open binding. No new CTest target.
