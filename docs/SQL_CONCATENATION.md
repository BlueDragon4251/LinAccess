# Jet SQL concatenation and dynamic LIKE

The SQL compiler wraps ampersand operands in a native scalar call before
emitting SQLite SQL. Arithmetic remains inside the operands, while comparisons,
Boolean operations, argument separators and query clauses remain outside.
Parentheses/functions are scanned structurally. This avoids changing Access
precedence by substituting SQLite's higher-precedence concatenation operator.

Two NULL operands produce NULL; one NULL is treated as empty text. Text retains
UTF-8. Integers retain their full SQLite value; floating-point and integer string
conversion uses the current Qt locale without grouping. Unsupported BLOB operands
and non-finite numbers fail explicitly. Boolean/Date Variant subtype formatting,
full locale/collation fidelity and the complete Jet expression grammar remain open.

Dynamic LIKE expressions receive ANSI-89 wildcard conversion after evaluation.
Percent/underscore/backslash literals are escaped for SQLite; * and ? remain Jet
wildcards. Simple literal and declared-parameter patterns keep the existing path.
Character classes and # patterns still fail explicitly rather than pretending to
support them. SQL function errors preserve bounded native error codes.

The dedicated test verifies arithmetic/comparison/IN/LIKE precedence, parenthesized
WHERE expressions, function arguments, chained concatenation, NULL versus empty,
UTF-8, numeric locale conversion, large integers, literal wildcard characters,
parameterized and nested LIKE expressions, missing operands and rejected BLOBs.

## Original corpus measurement

All 16 tests pass. The unchanged 100-query corpus now executes 49 queries without
an Access reference oracle (previously 38). Remaining classifications are 19
missing dependencies, 6 unsupported function/compilation cases, 22 other binding
or unsupported cases, 2 execution-budget failures and 2 action queries excluded
from this read-only corpus runner.

Six of nine original report sources now bind 54 non-NULL textbox values, compared
with an independent prepared reader over a 29-table imported MDB copy. Three
sources remain blocked by localized Formulare!form!control references. They are
not reported as empty datasets. Source VBA, databases and exports remain unchanged;
private inventory output contains hashes/status/counts, not source or row values.
