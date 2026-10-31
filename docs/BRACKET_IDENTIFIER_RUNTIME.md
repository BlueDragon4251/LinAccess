# Bounded Access bracket names and UTF-8 statement identifiers

A bracket name in VBA expressions now uses an explicit Access host context.
The native form host reads its typed control/field value, or returns the control
object for supported member access such as `[selection].ControlType`.
A same-named procedure local does not replace that Access value. Qualified
bracket Forms/Reports collection roots use the existing native object model.
A generic host without this context raises 438. Unqualified bracket names in
an external standard module are currently rejected with 438.

This supports exact control/field names in the implemented native instance
context, not general Access Eval expressions or its complete scope rules.
Calculated controls preserve their existing typed results and read-only behavior.
The acOptionGroup constant is 103 and is tested against the actual native group.

## Tests and evidence

`core_tests` distinguishes unsupported headless evaluation from typed host values,
checks both branches and a same-named Boolean local. The native option-group test
checks both real numeric selection states, a same-named local and bracket object
access through ControlType. Actual mouse selection and SQL persistence remain
covered by the preceding option-group regression.

`real_kassa_dao_event` extracts the unchanged bracket condition from original
module `6e8e0a0e41e1a1768cff4a0b79a07205c7b0f799eb8b465af5397843ce048ed8`
and executes it in a synthetic probe over its unchanged original form layout.
The referenced original TextBox is calculated from a combo column. An initial
test wrongly tried to assign that calculated control and correctly failed with
2448; the corrected test preserves the value, verifies assignment rejection and
executes the condition with the actual typed calculated value. This isolated
probe does not execute the complete original surrounding handler. Both branches
are separately exercised with the native numeric option-group fixture.

The expression lexer already accepted UTF-8 identifier bytes, but statement
patterns for Dim, assignment, Set and object members only accepted ASCII.
Two original bang-field assignments therefore failed syntax validation on
non-ASCII field names. The statement rules now consistently accept those
spellings, including parameter/loop/module-field declarations and members.

A new `dao_execution` regression using non-ASCII procedure, variable, object
variable and field names failed before the correction. It now performs real
Edit/Update and independently reads the resulting SQL value.

The real Kassa test also executes the unchanged assignment at source line 137
of the hash-verified module identified by prefix `824977e9d7c7`. Its original
recordset variable and field spellings remain intact. Actual non-NULL imported
values and then Null pass through native DAO records on disposable tables;
independent SQL queries check both results. This isolates one original statement,
not that module's complete multi-recordset copying workflow. All temporary
schema/data changes are rolled back, and original fixture bytes remain unchanged.

The syntax inventory is now 181 accepted procedures out of 187. This is explicitly
not 181 executed procedures. Six procedures still fail syntax validation and many
accepted procedures need additional runtime APIs and end-to-end verification.

## Remaining boundaries

Full Eval grammar, expression contents inside bracket shortcuts, macro names,
standard-module scope parity and Access error-number/event oracles remain open.
UTF-8 spelling support is not complete Unicode identifier validation, Unicode
normalization or locale-aware non-ASCII case folding. Existing case-insensitive
ASCII lookup and native Qt control-name folding have not been replaced by a
complete VBA name-resolution implementation. No general class/type compatibility
or complete original application execution is claimed.
