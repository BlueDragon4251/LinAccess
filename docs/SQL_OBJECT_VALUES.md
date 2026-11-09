# Live form and report values in Jet SQL

The native query compiler recognizes `Forms!document!control` and
`Reports!document!control`, including bracket identifiers, optional `.Value`, and
localized `Formulare` / `Berichte`. A connection-local SQLite scalar adapter
resolves these references against actual open Qt document controls. References
also work inside installed saved-query views; values are read at execution time,
not substituted into SQL strings or cached when a view is created.

The GUI resolver requires the GUI thread, the same database connection and
project path, the matching document kind and an open document. Missing forms,
reports and controls produce errors 2450, 2451 and 2465 respectively. A console
connection without a resolver fails with 3061. Null, text, numeric and Boolean
values preserve their scalar representation (True=-1). Binary/object values are
not fabricated. Resolver state belongs to the Qt SQL driver and survives a
close/open cycle; SQL functions are registered on each new native connection.

`PARAMETERS name Value` accepts bound scalar QVariant values, including Null,
without forcing numbers to text or numeric-looking text to numbers. Date inputs
use the existing canonical timestamp bridge. Unsupported values fail explicitly.

Verified evidence:

- `sql_object_values`: current values, mutation, Null/Boolean, localized names,
  saved-query views, close/reopen, independent connections and typed Value inputs.
- GUI object tests: actual Qt control mutation, closed-form errors and project
  isolation across two connections.
- Original Kassa regression: nine references from five unchanged query artifacts
  resolve against controls from original hash-checked form definitions. Eight
  references additionally accept synthetic unbound input changes. Bound/query
  controls keep their actual values and read-only behavior. This tests reference
  resolution, not the correctness of complete query results in Microsoft Access.
- Six original query artifacts declare parameters using `Value`; their unchanged
  declaration headers now compile. Unsupported syntax in the query bodies is
  classified separately; this does not claim all six full queries execute.

Limits: this is not a complete Access expression service. Arbitrary properties,
subform/report chains, implicit undeclared parameters, module functions, exact
Date/Currency Variants and background-thread UI evaluation remain unsupported.
The remaining three original report sources need their referenced forms open;
this package does not invent filter values or claim those report workflows run.


Bounded live scalar document/control properties now use actual native getters;
default values and explicit .Value retain their original adapter. See
[SQL_OBJECT_PROPERTIES_RUNTIME.md](SQL_OBJECT_PROPERTIES_RUNTIME.md). Existing
SQL-object/GUI/layout/original4/4 passes115.92s; arbitrary methods, deeper paths,
parameterized properties and full Access query parity remain open.
