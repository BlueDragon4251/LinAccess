# Checked native floating SQL arithmetic

Proven numeric declarations now participate in expression binding. Supported
unambiguous native physical type tags and consistent stored Byte/Integer/Long/
Single/Double/Boolean source declarations are used, including declared numeric
parameters. Bare ambiguous SQLite REAL/INTEGER columns do not acquire a guessed
logical type. A numeric parameter shadows an equally named unqualified field;
qualified real columns retain their actual source type.

Double +,-,*,/ paths and floating division of proven integral operands use the
same native arithmetic evaluator as VBA, with checked operands and Double result
metadata. Numeric literal division also uses floating arithmetic: 2/4 is0.5.
Arithmetic propagates through nested bound expressions and simple WHERE/HAVING
predicates; mixed Currency/Date predicate rewriting is deliberately excluded.
Source operand validation runs before the bound calculation; ordinary indexed
comparisons retain their original column predicates. Matching Null inputs
remain Null even with a zero divisor. Real DAO Double values/types are retained.

Native VBA numeric overflow now reports6 rather than generic5; zero divided by
zero reports6, ordinary division by zero11. The MS-VBAL5.6.9.3.5 Empty-divisor
exception for an original Single/Double/String/Date left value is retained where
that operand reaches this numeric path. Full coercion/Empty behavior and IEEE
exceptional-value assignment before raising an error remain unimplemented and
are not claimed as complete parity.

Read-only source type snapshots are owned by the actual Qt driver and checked
on its thread. Main schema version, current catalog presence, exact stored
definition and physical declarations are rechecked; no record values or private VBA module
source is cached. At most128 source snapshots are retained. Missing declarations
fall back to existing unproven typing; TEMP metadata shadows are refused and
TEMP source shadows do not reuse main-source hints. Real SQLite column origins
also carry proven numeric declarations through star projections without
replacing direct field expressions or losing their original headings/origins.

This does not complete arbitrary joined/derived bindings, Single-only division,
integral overflow/promotion, untyped raw SQL arithmetic, unary integral typing,
Date/Currency/Decimal arithmetic or regional/Windows equivalence. Existing
bounds remain explicit under F041/F067/F068/F071/F121/F122.

The existing core, access_query_model, dao_execution and mainwindow_documents
checks cover normative error numbers, actual tables created through DAO,
Double/integral/Boolean/Single-mixed operands, integer/floating literal division,
parameters/shadowing, NULL, predicates, real DAO types, unchanged direct headings,
star origins, catalog changes without schema-version changes and TEMP shadows.
No new CTest or catalog IDs are introduced.

The native index regression observes SQLite statement scan counters on the
actual executed query and requires zero full-scan steps for the indexed ID
predicate. Cache tests use the real importer for ambiguous REAL storage with
proven Double metadata, distinct from DAO-created physical DOUBLE declarations.
