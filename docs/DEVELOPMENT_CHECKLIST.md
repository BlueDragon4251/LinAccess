# Access compatibility implementation checklist

This tracks CODEX_HANDOFF.md and the development request. A checked item is scoped
exactly as described; it is not approval of complete Access compatibility.
Keep the repository private. Never attach original fixtures, VBA literals, database
rows, credentials, or raw tool logs to issues, PRs, releases or build artifacts.

## P1: reproducible relational import

- [x] Restore all four fixture files and verify their SHA-256 checksums.
- [x] Reproduce the published environment's failing import with MDBTools 1.0.1.
- [x] Isolate the SQLite backend's standalone primary-key clause parser failure.
- [x] Check the original code against MDBTools 1.0.0: 29/29 here; historical
  intermittent EOF failures have not been reproduced with these intact files.
- [x] Preserve exported primary keys, composite key order, required columns and indices.
- [x] Compensate for MDBTools 1.0.0's SQLite backend omitting keys/indices by
  reading structural metadata from its MySQL backend, without executing its SQL.
- [x] Full database import uses one hash-verified working copy and one transaction.
- [x] Retain all relationship metadata and enforcement/cascade flags. Enforce
  relationships marked for referential integrity between imported tables.
- [x] Compare every exported Kassa row/cell and NULL as multisets, all 29 tables,
  key positions and index counts; verify the original MDB remains unchanged.
- [x] Regression coverage for quoted identifiers, composite keys, uniqueness and rollback.
- [x] Exercise recipe foreign-key rejection and cascading deletion on a SQLite copy.
- [ ] Independently reproduce and explain the historical 0–20/29/EOF instability.
- [ ] Preserve original Jet column types, precision/scale, Currency and Decimal exactly.
- [ ] Preserve AutoNumber seeds, defaults, validation, lookups and collations exactly.
- [ ] Implement non-NULL OLE/attachment binary extraction and round trips.
- [ ] Exercise representative ACE/ACCDB fixtures and original-format write support.

## P2: VBA execution and Access object model

Verified incremental work ([details and limits](DAO_RUNTIME.md)):

- [x] Positional Sub/Function arguments, value-returning expression calls, ByVal/ByRef aliases.
- [x] Top-level label handlers, Resume retry/next/label and nested error-number propagation.
- [x] Eleven original keypad button/helper executions and typed binary form-property parsing.
- [x] Full 187-procedure, hash-verified syntax inventory, explicitly separate from execution coverage.
- [x] Native object references, CurrentDb and table-backed DAO dynasets/snapshots.
- [x] Recordset iteration, named/indexed Field.Value access and edit/insert/delete buffers.
- [x] Do While loops, multiple DAO Dim declarations and Nothing identity.
- [x] Original Kassa Qt button/unchanged handler performs a real DAO insertion.
- [x] Hidden controls remain accessible; space-containing control names bind to VBA handlers.
- [x] Empty/Null distinction, tested Boolean/bitwise rules and CLng tie-to-even rounding.
- [x] Procedure-local Resume Next/GoTo 0 and scoped Err object (partial error codes).

- [x] Typed Qt/native form/DAO field bridge, original Boolean metadata/True=-1 normalization, numeric input checks; see TYPED_FIELDS.md.
- [x] Bound-form Open/Load/Current, Dirty, control/form updates, navigation save and Unload/Close cancellation; see FORM_EVENTS.md.

Full targets remain open:

- [ ] Replace statement regex dispatch with a complete lexer/parser, AST and bound runtime.
- [ ] Implement typed Variants (Empty/Null/Error/Object/Date/Currency), coercions,
  banker rounding, Unicode, operator precedence and three-valued logic.
- [ ] Modules/classes/CodeBehind, scope/lifetime, static/module variables and references.
- [ ] Calls/returns, ByRef/ByVal, Optional/ParamArray, arrays, properties and collections.
- [ ] With, Select Case, Do/While, For Each, Exit and full conditional semantics.
- [ ] On Error/Err/Resume, nested handlers and source-mapped diagnostics/debugger.
- [ ] Application/Forms/Reports/Me/DoCmd and Access/DAO constants.
- [ ] CurrentDb/DBEngine/Workspace/Database, DAO transactions and ADO adapter.
- [ ] OpenRecordset, Fields/default members, EOF/BOF, navigation and bookmarks.
- [ ] Edit/AddNew/Update/Delete, QueryDefs/Parameters/TableDefs and Requery.
- [ ] Form record sources, filters/sort, focus and control state with real persistence.
- [ ] Open/Load/Current/Dirty/BeforeUpdate/AfterUpdate/Close/Unload ordering and Cancel.
- [ ] Mouse/keyboard/timer/combo/subform events; real Kassa UI actions verified.
- [ ] Access SetWarnings/dialog semantics without additional per-SQL consent prompts.
- [ ] Breakpoints, locals, stack frames, runtime errors and recovery.

Existing small subset execution tests remain mandatory. Static recognition of
187 procedures is not execution coverage.

## P3: queries, forms, reports and UI

Verified incremental work ([query adapter and limits](QUERY_ADAPTER.md)):

- [x] Imported QueryDefs, typed bound Parameters, temporary CreateQueryDef, SELECT/action execution.
- [x] Basic ANSI-89 strings/LIKE, single-table DISTINCTROW and TOP boundary ties.
- [x] Canonical MDB/date-literal/Date()/Now() timestamps at second precision.
- [x] Parameterless compatible query views, dependency refresh and form source exposure.
- [x] All 100 query artifacts hash-checked and classified on an isolated copy;
  37 execute without an Access result oracle, the rest remain explicitly unresolved.
- [x] SQL operation/time budgets and rejection of interrupted partial results.

Full targets remain open:

- [ ] Jet/ACE SQL parser/translator, parameters and saved query dependency resolution.
- [ ] JOIN/DISTINCT/UNION/TOP, TRANSFORM/PIVOT and Access expressions/functions.
- [ ] Date literals, wildcard semantics, Null/coercion, DDL/DML and action queries.
- [ ] Execute and classify all 100 Kassa queries against the imported database copy.
- [ ] Bind saved queries to actual form/report records; editable joins where valid.
- [ ] Navigator, datasheet/form/design views, properties, toolbox, undo/redo and SQL designer.
- [ ] Fonts/colors/geometry/default control properties and Access 2003 layout parity.
- [ ] Embedded linked subforms, tab/option groups, combos/lists and expression controls.
- [ ] Report sections/grouping/aggregates/headers/footers/pagination, report events.
- [ ] Print/PDF preview and output compared against original reference reports.
- [ ] Macros/action representation and native execution with error diagnostics.

## Corpus, platforms and delivery

- [x] Existing real Bridge regressions cover 192 artifacts and hashes, 47 forms,
  9 reports, 100 queries, 32 VBA files and 187 statically catalogued procedures.
- [ ] Per-artifact and per-procedure capability/execution classification, including failures.
- [ ] COM/ActiveX/WinAPI reference inventory and feasible native/optional Windows adapters.
- [ ] Real Windows Access 2003/365 Bridge export validation, 32/64-bit registration,
  references, encoding and failed-export diagnostics.
- [x] Local Debian 13 .deb build, exact payload check (application/desktop only),
  public install directory permissions even under umask 0077; package check wired into CI.
- [ ] Repeated Debian 12/13 and Ubuntu 24.04 remote CI builds/tests and package validation.
- [ ] Interactive Linux desktop validation and accessibility/translations.
- [ ] GitHub issue checklist and reviewed PRs; API access currently returns Forbidden.
- [ ] Full end-to-end operation of the original Kassa application, without compatibility claims
  based solely on import counts, synthetic UI actions, or static source recognition.
