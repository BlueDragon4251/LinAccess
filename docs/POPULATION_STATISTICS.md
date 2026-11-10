# Native population variance and standard deviation

VarP and StDevP now use the native scalar statistical accumulator alongside Var
and StDev. All use the existing stable incremental mean and second central
moment. Population results divide by the number of non-NULL inputs; sample
results retain division by n-1. Empty/all-NULL input yields Null, and a single
numeric input yields population zero while the sample functions retain Null.
Nonfinite input, result overflow and unsupported input types retain the existing
explicit errors rather than silently contributing zero.

Access SQL rewrites the real aggregate names to registered SQLite aggregate
callbacks; grouped, DISTINCT, parameterized and empty queries use that same
implementation. Result metadata identifies Double. Reports/form expressions
recognize both functions over the existing field and compound-expression paths,
using the existing requery and print aggregate caching. PDF output follows those
real calculated controls.

DVarP/DStDevP use the actual native project connection and existing validated
domain/expression/criteria path, including named Expr/Domain/Criteria arguments
and Application method calls. No business data or per-database special case is
introduced. Existing domain restrictions on unproven types, Currency/date/Boolean
statistics, explicit Null criteria and nested aggregates remain in force.

The existing access_query_model case checks independent 10/20/30 expectations,
large-offset data, grouped/singleton/all-NULL/empty input, DISTINCT parameters and
invalid types/overflow. Existing access_2003_layouts cases check real report
controls, PDF export/requery and actual domain field/compound/named/Application
calls. This is a bounded contribution to F085/F052, not full Access aggregate,
regional, event or floating-point equivalence. The fixed catalog remains
456 IDs with 386 confirmed partial scopes and 70 open full requirements.

The visual query builder also offers VarP and StDevP in its actual totals dialog.
Saved field modes survive restoration and generate the real aggregate SQL for
physical and calculated fields. The existing mainwindow_documents designer case
checks the selected modes and executes the generated statement against independent
2/4/6 expectations. Other available source fields remain hidden Where rows under
the existing design restoration rules; no source field is removed to fit a test.
