# Access compatibility implementation checklist

This tracks CODEX_HANDOFF.md and the development request. A checked item is scoped
exactly as described; it is not approval of complete Access compatibility.
The 456 existing entries are frozen and indexed as F001–F456 in
[the fixed work catalog](FIXED_WORK_CATALOG.md). Add implementation details to
existing requirements; do not append more checkbox entries. The ratio of scoped
checks is not overall Access 2003 parity.
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
- [x] Reproduce 9/29 even with CLI 1.0.0 by changing only its loaded MDB library; tie all 20 rejected schemas to standalone PK clauses; current full test passes all four version/library combinations.
- [ ] Independently reproduce/explain the remaining historical intermittent EOF reports.
- [ ] Preserve original Jet column types, precision/scale, Currency and Decimal exactly.
- [ ] Preserve AutoNumber seeds, defaults, validation, lookups and collations exactly.
- [ ] Implement non-NULL OLE/attachment binary extraction and round trips.
- [ ] Exercise representative ACE/ACCDB fixtures and original-format write support.

## P2: VBA execution and Access object model

Verified incremental work ([details and limits](DAO_RUNTIME.md)):

- [x] Positional Sub/Function arguments, value-returning expression calls, ByVal/ByRef aliases.
- [x] Top-level label handlers, Resume retry/next/label and nested error-number propagation; obsolete snapshots release collection leases and DAO clones after restored scopes end.
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

- [x] Native MsgBox button groups/results, optional positional arguments and real SQL branching; unchanged Kassa missing-input handler opens its dialog and focuses the actual control. See MESSAGE_BOX_RUNTIME.md.
- [x] Bounded unqualified form-control value/object reads and writes, with existing locals/parameters taking priority; full class/scope semantics remain open.

- [x] Native late-bound DoCmd.Close for active/named forms and reports, validated record saving, Unload cancellation and recursion protection; original Kassa Close button executes unchanged. See CLOSE_COMMAND_RUNTIME.md.
- [x] ComboBox.Dropdown opens actual SQL-backed rows and preserves a native mouse selection as a typed bound value; original missing-selection handler runs unchanged.

- [x] Currency scaled-unit changes survive nested ByRef synchronization, shared aliases and handled errors; original calculated Currency value passes through a separate native nested helper exactly.

- [x] Bounded persistent CodeBehind instance fields and connection-owned standard-module fields; real original public declarations bind across native forms, with recursion, typed alias recovery, visibility and reset tests. See MODULE_STATE_RUNTIME.md.
- [x] Procedure-wide Dim binding, local/control/global lookup precedence and standard-module Me restrictions; complete compiler/class scope remains open.

- [x] Single-line If/Else/colon branches, reference-aware default invocation and scalar defaults for supported native objects; original calendar Error/Resume path and real DAO/subform SQL writes verified.
- [x] Original function-expression date-field events, native Screen/Parent getters and bounded MSCAL adapter; both unchanged original date bindings navigate/select/close and feed actual MDB records in an explicitly enabled calendar context. See CALENDAR_RUNTIME.md.

Full targets remain open:

- [ ] Replace statement regex dispatch with a complete lexer/parser, AST and bound runtime.
- [ ] Implement typed Variants (Empty/Null/Error/Object/Date/Currency), coercions,
  banker rounding, Unicode, operator precedence and three-valued logic.
- [ ] Modules/classes/CodeBehind, scope/lifetime, static/module variables and references.
- [ ] Calls/returns, ByRef/ByVal, Optional/ParamArray, arrays, properties and collections.
- [x] Native For Each/Next, Exit For, nested loops and DAO Fields enumeration; real 29-table structure regression.
- [ ] Complete With/Select Case/Do/While/Exit grammar, type/error and conditional semantics.
- [x] Imported-form Forms/Controls collections, guarded references, typed cross-form data writes and basic Access control constants.
- [x] Value/reference contexts for Set and property assignment, dynamic Me(control), DAO Field bang references; six property writes in one original Form_Load verified.
- [x] Unlinked native embedded child forms, Form references, Requery, typed edits and parent-close cancellation; original complex Form_Load and 15 actual data-driven Tags verified.
- [x] Bounded Error Variants, CVErr/IsError, ByRef/function return and calculated-control error handling.
- [ ] Complete Forms/Reports object model, properties, lifetime and subform semantics.
- [ ] On Error/Err/Resume, nested handlers and source-mapped diagnostics/debugger.
- [ ] Application/Forms/Reports/Me/DoCmd and Access/DAO constants.
- [ ] CurrentDb/DBEngine/Workspace/Database, DAO transactions and ADO adapter.
- [ ] OpenRecordset, Fields/default members, EOF/BOF, navigation and bookmarks.
- [ ] Edit/AddNew/Update/Delete, QueryDefs/Parameters/TableDefs and Requery.
- [ ] Form record sources, filters/sort, focus and control state with real persistence.
  Bounded direct-table DataEntry is implemented; writable-query/complete Windows
  semantics remain open (DATA_ENTRY_RUNTIME.md).
- [ ] Open/Load/Current/Dirty/BeforeUpdate/AfterUpdate/Close/Unload ordering and Cancel.
- [ ] Mouse/keyboard/timer/combo/subform events; real Kassa UI actions verified.
- [ ] Access SetWarnings/dialog semantics without additional per-SQL consent prompts.
- [x] Breakpoints, locals, stack frames, runtime errors and recovery in the native supported runtime; see [scope and boundaries](NATIVE_VBA_DEBUGGER_RUNTIME.md).

Existing small subset execution tests remain mandatory. Static recognition of
187 procedures is not execution coverage.

## P3: queries, forms, reports and UI

Verified incremental work ([query adapter and limits](QUERY_ADAPTER.md)):

- [x] Imported QueryDefs, typed bound Parameters, temporary CreateQueryDef, SELECT/action execution.
- [x] Basic ANSI-89 strings/LIKE, single-table DISTINCTROW and TOP boundary ties.
- [x] Canonical MDB/date-literal/Date()/Now() timestamps at second precision.
- [x] Parameterless compatible query views, dependency refresh and form source exposure.
- [x] All 100 query artifacts hash-checked and classified on an isolated copy;
  77 execute (72 ordinary + five parameterized) without an Access result oracle, the rest remain explicitly unresolved.
- [x] SQL operation/time budgets and rejection of interrupted partial results.

Full targets remain open:

- [ ] Jet/ACE SQL parser/translator, parameters and saved query dependency resolution.
- [ ] JOIN/DISTINCT/UNION/TOP, TRANSFORM/PIVOT and Access expressions/functions.
- [x] Bounded date Format/Format$ masks in SQL/VBA; actual values from ten original date columns, NULL/error/locale/DST tests.
- [x] Native SQL concatenation with arithmetic/Boolean boundaries, NULL and dynamic LIKE regression; original query execution 49/100.
- [ ] Complete date/format masks, wildcard semantics, Null/coercion, DDL/DML and action queries.
- [x] Saved ranked TOP dependency views, NULL/tie-safe bounded TOP 1 plans and qualified column labels; 77/100 original queries execute after subsequent query-origin work.
- [x] Nested SQL budget enforcement and guarded simple RIGHT JOIN adaptation with independent NULL/duplicate results.
- [x] Pure IIf projection-alias chains and leading-dot decimal/exponent literals; three additional original queries produce positive non-NULL results (77/100 execute).
- [ ] Execute and classify all 100 Kassa queries against the imported database copy.
- [x] Bounded read-only saved-query form sources and translated navigator results with explicit typed parameter prompts.
- [x] Live SQL Forms/Formulare and Reports/Berichte control values, project isolation and saved views; nine original references / five queries verified.
- [x] Typed Value parameter declarations in all six original parameterized queries.
- [ ] Complete form/report query binding, implicit parameters and editable joins where valid.
- [ ] Navigator, datasheet/form/design views, properties, toolbox, undo/redo and SQL designer.
- [ ] Fonts/colors/geometry/default control properties and Access 2003 layout parity.
- [x] SaveAsText quoted string continuation/escaping, typed Column/BoundColumn values, basic form aggregates and explicit subform references.
- [ ] Embedded linked subforms, tab/option groups, complete combos/lists and expression controls.
- [x] Bound Report_Open/Close/NoData, cancellation and read-only sources; all 9 original report sources classified, six bind 54 actual non-NULL controls.
- [ ] Report sections/grouping/aggregates/headers/footers/pagination and complete report event order.
- [x] Native current-layout single-page print/PDF preview and CommandBars adapter; three original Report_Open procedures executed explicitly.
- [ ] Full report print/PDF pagination and output compared against original reference reports.
- [ ] Macros/action representation and native execution with error diagnostics.

## Corpus, platforms and delivery

- [x] Existing real Bridge regressions cover 192 artifacts and hashes, 47 forms,
  9 reports, 100 queries, 32 VBA files and 187 statically catalogued procedures.
- [x] Per-artifact and per-procedure capability/execution classification, including failures; [instrumented paths and explicit unproven states](PROJECT_CAPABILITY_REPORT.md).
- [ ] COM/ActiveX/WinAPI reference inventory and feasible native/optional Windows adapters.
- [ ] Real Windows Access 2003/365 Bridge export validation, 32/64-bit registration,
  references, encoding and failed-export diagnostics.
- [x] Local Debian 13 .deb build, exact payload check (application/desktop/private SQLite/notice),
  public install directory permissions even under umask 0077; package check wired into CI.
- [x] Repeated Debian 12/13 and Ubuntu 24.04 remote CI builds/tests and package validation.
- [ ] Interactive Linux desktop validation and accessibility/translations.
- [ ] GitHub issue checklist and reviewed PRs; current GitHub API reads work.
- [x] Remote CI jobs start successfully again after the earlier account payment/
  spending-limit block; subsequent successful runs are linked in PROGRESS.md.
- [ ] Full end-to-end operation of the original Kassa application, without compatibility claims
  based solely on import counts, synthetic UI actions, or static source recognition.

- [x] Native ActiveMenuBar and popup Controls share real project actions; unchanged complex original Form_Open completes, with mixed numeric/String Variant comparison regressions.
- [x] Original OnCurrent completes; two original DSum controls independently match actual record sums after SaveAsText escaped-quote correction.
- [x] DSum/DCount execute native table/view aggregates with criteria, Null/count and catchable-error regressions.
- [x] Bounded Currency subtype, CCur, fixed-point operations and original calculated expression with a synthetic non-Null control input; initial eager CCur(Null) error 94 remains explicit.
- [x] Actual native activation, deferred first Current, form-switch and close-cancellation tests; original Open/Load/Activate/Current complete.
- [ ] Full focus/resize/PopUp/Modal/MDI/reentrant event semantics and Microsoft Access event oracle.
- [ ] Access oracle for the remaining initial calculated error, full Currency coercions/locales, exact Jet MONEY storage and Decimal semantics.

- [x] Jet bracketed derived tables, full-field DISTINCTROW, unique column-origin rebinding, bounded pure aliases and positive original parameter results.
- [ ] Resolve/compare the original six-column source mismatch against Access; full nested/intermediate lineage, ambiguous records and aggregate/volatile alias semantics remain open.
- [x] Native QueryDef.Parameters Count/Item/For Each, stable object identity, generation/close invalidation and original six-query collection regression.

- [x] Native Filter/FilterOn, FilterOnLoad and UI, typed editable table/readonly query filtering, validation/recovery and twelve original bound-cell checks.
- [x] Direct Me.Requery/Me.SetFocus and bound Requery Current dispatch; original unbound Requery cycle regression.
- [ ] Complete focus-specific DoCmd.Requery targeting, sorting, advanced/nested criteria, report/reentrant event semantics and designer filter persistence.


- [x] Bounded four-position OpenForm/acNormal/WhereCondition with omitted defaults,
  actual filtered rows/cells and original wizard-handler execution; typed failure
  and cancellation propagation (OPEN_FORM_RUNTIME.md).
- [x] Native OptionGroup numeric literal defaults, nested choice metadata,
  exclusive selection, Null clearing, typed getter and direct-table persistence;
  original target Open/Current initialization regression.
- [ ] Full OpenForm modeless/modal/MDI lifetimes, existing-instance reuse, named
  filters, designer/datasheet views, DataMode/WindowMode/OpenArgs and Access oracle.
- [ ] Full option-group default expressions, bound new-record defaults,
  individual child Value assignments, cancelled-update rollback and event oracle.


- [x] Bounded Access bracket control/field reads with typed values, explicit host
  context, local-shadow isolation, native object property access and original
  calculated-control condition/read-only regression.
- [x] Consistent UTF-8 identifier spelling in statement declarations/assignments,
  original Unicode DAO statement with actual imported values and Null writes;
  reproduced pre-fix DAO failure (BRACKET_IDENTIFIER_RUNTIME.md).
- [ ] Full Access Eval grammar/scope, Unicode identifier validation and locale
  case folding; full surrounding original record-copy handlers and Access oracle.


- [x] Bounded native TableDefs Count/Item/For Each, TableDef.Name and real
  OpenRecordset; original standard-module function checks all 29 tables and an
  absent name with typed Boolean results and independent record counts.
- [x] Single-line If/colon Exit For binds to the enclosing loop; nested numeric
  loop and unchanged original For Each regressions; basic function result defaults.
- [ ] Complete DAO catalog Fields/Indexes/Relations, Refresh/Append/Delete,
  linked-table metadata and immutable generation identity across drop/recreate.
- [ ] Full declared result/local/argument type enforcement and Access DAO oracle.


- [x] Bounded nested With, once-only object evaluation, relative DAO writes,
  Resume recovery and Exit For scope restoration; native CodeContextObject.
- [x] Actual native SendKeys/GoToControl with queued focus, Unicode, preflight
  rejection and bound persistence; three unchanged original keyboard handlers
  and three independent receiver/text checks (WITH_KEYBOARD_RUNTIME.md).
- [x] Typed displayed calculated-control cache, input/navigation invalidation
  and explicit Requery; unbound TextBox.Requery retains its value.
- [x] Unique generated dotted control aliases, exact-name precedence, literal
  bang/Controls.Item and ambiguity/standard-module scope regressions.
- [x] Per-module Option Compare for relational strings; native Qt folding and
  explicit generic-host non-ASCII limits; original table-existence case failure fixed.
- [x] Local SELECT INTO with parameters, actual creation/row counts, error 3010
  and nested rollback; original 14-column statement on an empty sale snapshot.
- [ ] Full initialized original sale workflow: subsequent first/second keyboard
  AfterUpdate still fails 3075/1023 without a sale ID; independent document
  lifetimes, positive original sale snapshot and real report/printing remain open.
- [ ] Full Jet/locale collation, comparison-aware Like/StrComp/InStr and complete
  SELECT INTO metadata/AutoNumber/index/default/validation propagation.


- [x] Native OpenForm continuation with persistent target, same-project unique
  instance reuse, canonical catalog Name, actual continuing Forms writes/save;
  imported Modal blocks UI without blocking VBA (OPEN_FORM_RUNTIME.md).
- [x] Immediate logical Screen.ActiveForm/ActiveControl after OpenForm/SetFocus,
  caller-tree child close/save/cancellation and guarded object invalidation.
- [x] Deferred physical window deletion until native VBA unwinds, including a
  self-close followed by a real MsgBox and nested Qt event-loop regression.
- [x] Original bound-layout WhereCondition and unchanged wizard/target opening;
  both original calendar bindings perform real asynchronous selection and
  data-query write-back in an explicitly enabled calendar context, with untouched fixture bytes.
- [ ] Independent application document lifetimes/MDI, close-time reentrant opens,
  complete active-document DoCmd/keyboard scheduling across windows, acDialog /
  other window/data modes, OpenArgs and an Access event/window oracle.


- [x] Bounded same-project active-document GoToControl, SaveRecord, Next/Last/New,
  Maximize/Requery with independent persistence; explicit Me remains instance-bound.
- [x] Application.DoCmd command paths, GoToRecord omission/constants and scalar
  argument validation; actual RunSQL CREATE/INSERT/UPDATE/DELETE/DROP, no extra
  LinAccess consent dialogs (ACTIVE_DOCUMENT_COMMANDS.md).
- [x] Per-delivery project focus for SendKeys, queued form switches, independent
  same-project windows and foreign-project save isolation.
- [x] Actual original-layout GoToControl/Last/Requery over imported MDB records,
  independent bound-cell/Null checks in the documented CodeBehind-free probe;
  unchanged original wizard/calendar/keyboard regressions remain mandatory.
- [ ] Full active-subform/navigation/end-of-record and unavailable-command parity,
  remaining command arguments, RunSQL UseTransaction, SetWarnings/Access dialogs,
  complete window event scheduling and the initialized original sale workflow.


- [x] Independent OpenForm lifetime through the project window, preserved actual
  command bars and original opener/target execution after caller Close/destruction.
- [x] Project switch/main-window close and application Quit validate native Access
  documents, save records, respect Unload cancellation, reject reentrant shutdown
  opens and clear guards on both success/cancellation (DOCUMENT_LIFETIME_RUNTIME.md).
- [x] Dynamic RowSource/RowSourceType/BoundColumn/ColumnCount/ColumnWidths with a
  typed multicolumn Qt model, hidden columns, actual mouse selection/persistence
  and failed-assignment recovery (DYNAMIC_COMBO_RUNTIME.md).
- [ ] Complete native-editor/designer shutdown, Access event/window oracle,
  full initialized sale workflow and actual report/print comparison.


- [x] Bounded native form Escape field/record undo, Shift+Enter save, F9 Recalc,
  Shift+F9 Requery, F2 editor selection, Enter focus advance and Ctrl+W/Ctrl+F4 close.
- [x] Native Form.Recalc and Form.Undo, typed NULL/numeric/duplicate-field undo,
  validation and actual SQL/focus regressions (KEYBOARD_RUNTIME.md).
- [x] Native bound KeyDown/KeyPress/KeyUp and dynamic KeyPreview with ByRef
  cancellation, ASCII replacement, modifier masks and event-order SQL traces.
- [ ] Full Access 2003 key/layout/menu/designer parity, IME/locales, complete
  key/Undo event scheduling and a Windows Access event/rendering oracle.
- [ ] Validate the original login ColumnCount/index mismatch and two renamed
  controls' missing KeyPress handlers against Access 2003; do not invent bindings.

- [x] Automatically load real multicolumn combo choices on mouse/Alt+Down popup,
  preserve typed numeric-ID search drafts, and recover from RowSource errors.
- [x] Remove the model use-after-free after QComboBox::setModel on Qt 6.4/6.8.
- [x] Force UTF-8 only for MDBTools child processes; test Unicode MDB paths and
  original table names under an unchanged C-locale parent.
- [x] Finish pinned SQLite backend, full Debian/Ubuntu regression and installed
  package verification; retain explicit Jet/Access compatibility boundaries.
- [x] Preserve control EventProcPrefix, expose bounded in-memory read/write and
  route supported native events through the explicit prefix; actual GUI/SQL
  validation, cancellation, keyboard and missing-binding regressions.
- [ ] Exact Access 2003 prefix-name edge cases and design-view source refactoring.
- [x] Native InputBox/InputBox$ expression results, optional positional slots,
  cancellation and text limits; shared dialog in both GUI hosts, actual bound
  SQL/paste/IME regression in the imported-layout host, explicit unsupported
  no-GUI/Windows Help contexts.
- [ ] Original InputBox procedure end-to-end with synthetic credential records;
  named/statement arguments, Windows Help and Access 2003 rendering oracle.
- [x] DAO clone cursor isolation, typed byte-array bookmarks, bounded Find methods/
  NoMatch, shared writes, deletion/rowid-reuse and stale/foreign token checks.
  Includes bounded Recordset.Move with Long Rows and optional native Start bookmark.
- [x] Restricted form Recordset.Clone + Form.Bookmark navigation through validation,
  shared saved values and original generated search regression: two successful
  native selections; third original missing-control reference remains error 2465.
- [ ] Full editable Form.Recordset, arrays/BLOBs, concurrent dynasets/locks/Seek,
  complete criteria/formatting and Access 2003 reference behavior.
- [x] Native Value List/Wertliste text rows and Field List/Feldliste metadata-only
  query columns; real popup/SQL persistence, quoting/Unicode, failure rollback,
  parameterized description and no-expression-evaluation regressions.
- [ ] Exact Access 2003 empty/whitespace list semantics, ColumnHeads,
  AddItem/RemoveItem, full ListBox and custom row-source callbacks.
- [x] Bounded native Eval/Application.Eval, public Function execution, Access Null
  tests, actual Forms/SQL and modal self-close lifetime regression; unchanged
  original module with 15 Eval references on synthetic native controls.
- [ ] Full Eval grammar, function resolution/coercion and Access 2003 oracle.
- [x] Read-only native ComboBox ListCount with actual lazy row-source loading,
  preserved selection/Dirty state and VBA-to-SQL/Requery regression.
- [ ] Compare Access 2003 ListCount list-header, ListBox and record-limit behavior;
  original fixtures contain only commented ListCount references.
- [x] Bounded native single-record selection/deletion, AllowDeletions, actual
  confirmation and Delete/BeforeDelConfirm/AfterDelConfirm cancellation, rollback,
  bookmark/FK/reentrancy and active-document regression; one original delete
  procedure via explicit invocation and bound native click on synthetic rows.
- [ ] Complete clipboard interoperability, multi-record selection/deletion, complete writable
  query sources, full deletion focus/event ordering and Access 2003 oracle.
- [x] Preserve 21 original AutoNumber flags and native AUTOINCREMENT generation
  for the 16 single-primary-key counters, retaining actual imported values and
  verifying no-reuse on isolated copies (AUTONUMBER_IMPORT.md).
- [x] Generate the five non-primary-key counters with persistent atomic defaults,
  native form/DAO/clipboard use, reopen and concurrent-connection regression.
- [ ] Preserve stored source seeds, random/replication modes, full counter property
  semantics and exact Jet allocation/overflow behavior.

- [x] Native single-record Copy/PasteAppend and selected-record Ctrl+C, typed
  clipboard, real SQL/bookmark/active-target tests and original manual-key conflict
  via direct invocation and native click on synthetic rows.
- [ ] Complete insertion event pipeline, multi-record/external clipboard formats
  and Access 2003 reference behavior; unsupported contexts fail explicitly.

- [x] Bounded joined action UPDATE: native single-table INNER JOIN execution,
  frozen source, typed deterministic duplicates, ambiguity rollback, parameters,
  physical identity checks and unchanged original query on synthetic schema copies.
- [ ] General Jet joined-query updateability, outer-join targets and Access oracle.

- [x] Active-form GoToRecord first/last/next/previous/absolute/new with offsets,
  explicit acActiveDataObject, save cancellation and new-row boundaries.
- [x] Named GoToRecord acDataForm targets, including hidden forms and project/lifetime checks.
- [ ] Native table/query/report GoToRecord targets and datasheet navigation.

- [x] Named-control DoCmd.Requery on active form, distinct from Me.control.Requery;
  native SQL row-source reload and bound-control form reload with validation.

- [x] Bounded native-control BeforeInsert/AfterInsert with Cancel/error/Undo,
  programmatic-assignment distinction and validation/constraint retry tests.
- [ ] Access BeforeInsert/Change/Text/Value timing oracle and PasteAppend events.

- [x] Imported native text/combo OnChange, EventProcPrefix dispatch, programmatic
  Value suppression and KeyDown/KeyPress/BeforeInsert/Change/KeyUp regression.
- [x] Bounded native Text/Value buffers, focused Text setters, control-update
  visibility, cancellation/retry, SQL persistence and finite cascading Change.
- [ ] Full input/focus/popup timing, arbitrary update reentrancy, Undo events/other
  control kinds, multiline/IME and complete combo validation parity.

- [x] Focused native editor SelStart/SelLength/SelText with replacement,
  Change/SQL effects, Null/no-selection, validation and UTF-16 regression.

- [x] Native form TimerInterval/OnTimer, Load initialization, hidden execution,
  stop/restart, close/cancel/lifetime guards and asynchronous SQL regressions.
- [x] Original clock Timer handler through its original binding on a prepared
  project copy; two timestamps independently checked (not full startup).
- [ ] Access modal/timer scheduling and original clock display-locale oracle.

- [x] Single LEFT/RIGHT joined UPDATE to the preserved direct table, including
  unmatched NULL sources, Is Null filtering, affected counts and atomic rollback.
- [ ] Nullable-side/FULL joined updates and writable saved-query targets.

- [x] Native DoCmd.OpenReport preview with WhereCondition, imported Report VBA,
  failure/cancellation and independent lifetime; original preview handler and
  bound button open the original report with independently checked fields.

- [x] Bounded UTF-16 Len/Left/Right/Mid and dollar substring variants, shared SQL
  Len, embedded NUL, malformed-text validation, Null/Long error semantics and
  scalar getter/procedure/alias regressions; split surrogate strings stay unsupported.

- [x] Trim/LTrim/RTrim and dollar variants: end spaces only, preserved non-space
  whitespace, scalar defaults, Variant Null versus String 94 and VBA error tests.

- [x] Bounded InStr/StrComp: UTF-16 positions/binary order, Option Compare,
  scalar defaults, Null/errors and native VBA tests; locale matching remains open.

- [x] Bound OldValue through form AfterUpdate, text/combo control Undo with
  other-edit preservation, BeforeUpdate self-change rejection and draft retry.

- [x] Bounded project DBEngine/Workspace transactions with real savepoints,
  nested commit/outer rollback, isolated-reader SQL/VBA/native-button evidence
  and rollback/connection-generation guards; broader DAO scope remains open.

- [x] SaveAsText NotDefault Boolean semantics, native Enabled/Visible persistence,
  TextBox Locked, TabStop and bounded flat TabIndex traversal/reordering, including
  original layout keyboard evidence and preserved read-only input protection.

- [x] Native TextBox Password mask with real masked echo, unchanged VBA/SQL
  value, dynamic updates, Null/Unicode/navigation and original-layout probe;
  generic masks remain unsupported.

- [x] Bounded flat form Cycle 0/1 with native Tab/Shift-Tab, validated focus,
  actual record SQL/navigation, cancellation/retry, empty-new-row handling,
  handler focus override and close/MsgBox lifetime regression.

- [x] Bounded multiline Select Case lists/ranges/relations/Else, one-time selector,
  nested conditional/loop stacks, Resume regression and native Qt/SQL branches.

- [x] Bounded Do Until, postcondition Loop While/Until, unconditional Do/Loop,
  While/Wend and Exit Do with nested frame unwinding, error retry, budget limits
  and native transaction-button SQL evidence.

- [x] Bounded static control-font import/templates and native Qt text rendering,
  including editable combo editor/popup, real SQL/navigation persistence and
  original form/report layout probes; dynamic/exact visual parity remains open.

- [x] Bounded native form VBA FontName/Size/Weight/Bold/Italic/Underline properties,
  with preflight errors, scalar conversion, preserved masked text/selection/SQL,
  read-only-view styling and runtime font persistence across requery/navigation;
  documented >=700 Bold, 8/10-point defaults and unsupported Report OnPrint preflight.

- [x] Bounded multiline ElseIf first-match chains with Null/inactive/side-effect
  handling, nested loop exits, error retry and displayed native SQL mode selection.

- [x] Bounded Form.AllowEdits import/VBA/user-input policy, programmatic dirty
  exception/save/Undo, unbound/new-row handling, deletion preflight, checkbox and
  group-only Click, callback lifetime and five original flag-retention probes.

- [x] Bounded form-control DefaultValue metadata/templates and literal new-row/
  unbound initialization, no spurious events/inserts/ID allocation, native VBA/
  SQL/empty-NULL/mirror/error recovery evidence and 34/15 original metadata probes.

- [x] Bounded native TextBox/Label TextAlign metadata/templates and VBA styling,
  typed General/Left/Center/Right, pending-input/SQL/read-only preservation and
  24 original layout probes (208 Labels / 226 TextBoxes); distributed/report
  runtime setters and complete visual parity remain open.

- [x] Bounded native ForeColor templates/RGB/system text colors and VBA RGB(...),
  native pixel evidence, masked buffer/SQL/read-only/popup preservation and
  45 original layout probes (271 RGB / 277 system-color controls); complete
  OLE palettes, backgrounds and Windows visual equivalence remain open.

- [x] Bounded native TextBox/Label BackColor/BackStyle templates and VBA updates,
  real opaque/transparent composite pixels, masked-buffer/SQL/permission/
  navigation preservation and 55 original layout probes (326 Normal / 587
  Transparent / 414 colors); other adapters and full Windows visuals remain open.

- [x] Shared native VBA/SQL RGB primitive with real query/DAO fields, parameters,
  quoted identifiers, alias/view reuse, error transport and atomic direct/joined
  updates including invalid channels and conflicting source colors.

- [x] Shared native VBA/SQL QBColor classic 16-entry palette, all-index and
  conversion/error tests, actual Qt ForeColor/Eval effects, parameter/alias
  queries and atomic DAO updates through a compiled palette view.

- [x] Managed independent imported form/report opening from the actual main
  window/file-dialog/explorer route, singleton reuse, modal metadata, Open/
  Unload cancellation, SQL saving and deferred self-close/MsgBox lifetime;
  legacy generated-form modal migration and complete MDI/F11 remain open.

- [x] Bounded Access 2003 F11 explorer activation/restored splitter, native
  KeyPreview/control cancellation, modal blocking, logical database-window
  context, stale-form command guards, explicit-owner saving and Screen/
  activation/early-Current regression through the actual main-window GUI.

- [x] Bounded native DoCmd.SelectObject for real table/query/form/report explorer
  targets and open native forms/reports, metadata grouping/identity, hidden
  restoration, omitted names, scalar preflight, validation/modal/error and
  unsaved SQL preservation; other open-window/object kinds remain unsupported.

- [x] Bounded Ctrl+O/Ctrl+N real project open/create dialogs from the main window
  and nonmodal native documents, KeyPreview/validation/modal/Unload cancellation,
  overwrite preservation and an executing-VBA project-switch guard; direct
  MDB/ACCDB project editing and complete File-menu/keyboard parity remain open.

- [x] Bounded Ctrl+F6/Ctrl+Shift+F6 native form/report/database-window ring with
  focus restoration, validation/KeyPreview/modal guards, hidden/closed/foreign
  exclusion, minimized restoration and target-close lifetime regression; full
  MDI/window-order and table/query/editor cycling remain open.

- [x] F9 real row-model refresh for focused native combo adapters, KeyPreview
  cancellation, draft/focus/dirty/SQL preservation, query-failure recovery and
  F4 popup/cancellation regression; full list-box/popup/Windows timing remains open.

- [x] Bounded read-only native ComboBox ListIndex/ItemData over the actual typed
  bound-column model, lazy loading, scalar/index/error/report/draft boundaries,
  BoundColumn=0 and corrected NULL-vs-empty selection; ListBox/multi-selection
  and exact Windows pending-buffer/error/subtype identity remain open.

- [x] Persistent native single-selection ListBox QTreeView with shared real typed
  SQL/value/field rows, hidden columns, defaults, bound/unbound/NULL values,
  mouse/keyboard/control validation, SQL save/Undo/Locked/report/failure-recovery
  tests and retained MultiSelect diagnostics; complete headers/multi-selection/
  ListBox API/Windows visual and event parity remain open.

- [x] Bounded native ColumnHeads templates/VBA/real header views for query/value
  ComboBox and single-selection ListBox models, documented header-aware ListCount,
  value-row separation and SQL/scalar/error/report tests; field captions/Field
  List headers/exact Windows extra-index/visual context remain open.

- [x] Bounded native Simple/Extended ListBox selection and Selected/ItemsSelected/
  ItemData/ListIndex getters, real mouse/keyboard and indexed VBA setters, scalar/
  collection/lifetime/validation tests and NULL-value field-save/mirroring/new-row
  safety; Simple selected requery, ListIndex writer, exact Windows events/types/
  extra-header indices and full VBA array/class semantics remain open.

- [x] Native ComboBox ListRows template/default/VBA range and actual popup cap,
  measured 3/5-row viewport, 40-row scrolling/SQL selection, retained draft/model
  identity and scalar/error/report limits; full Windows header/theme/screen-edge/
  open-popup timing and declared Integer identity remain open.

- [x] Bounded Integer/Long module fields with zero defaults, 16/32-bit rounded
  assignment/range/error preservation, private/public/project lifetime and actual
  SQL tests; full numeric subtype/arithmetic/ByRef/array/class identity remains open.

- [x] Bounded checked Integer/Long locals, parameters, function results and integral
  For counters, known ByRef conflicts/temporaries/copyback and pre-effect invalid
  scalar-object/enumerator guards with real SQL tests; full subtype/promotion/
  Variant-reference/loop/compile/error timing and other types remain open.

- [x] Checked Boolean/String locals, parameter/result kinds and known-reference
  conversion/copyback, pre-effect invalid scalar loop/object guards and actual
  typed DAO Unicode/Boolean persistence; full Variant/object-default/locale/type/
  compiler/reference semantics and other primitive types remain open.

- [x] Declared SQL BIT/BOOLEAN/LOGICAL/LOGICAL1/YESNO parameter aliases, strict
  finite Boolean binding, NULL preservation and real DAO UPDATE/error atomicity
  tests; full Parameter metadata, assignment timing and Variant/locale identity open.

- [x] Native VBA CBool with Boolean kind/type identity, Empty/numeric/Currency/
  text/object conversion, strict argument/error checks and actual DAO persistence;
  full locale/Date/SQL conversion and other numeric conversion functions open.

- [x] Integer/Long value tags for checked fields/locals/parameters/results/defaults,
  Variant copies, native CInt/tagged CLng, VarType/TypeName and VbVarType constants
  with actual DAO persistence; full literal/arithmetic/Variant-storage/Len semantics open.

- [x] Checked Single/Double module/procedure storage, native CSng and conversion
  dispatch, type identity, float rounding/overflow and actual DAO persistence;
  Single underflow/For, full formatting/promotion/Variant/locale semantics open.

- [x] Declared Currency fields/locals/parameters/results/defaults, exact checked
  conversion/reference paths and real DAO decimal TEXT transfer at the upper
  boundary; full Jet Currency columns/SQL, mixed promotion and loops remain open.

- [x] Correct 32-bit SQL INTEGER and numeric parameter aliases, rounded/finite
  Single bindings, declared DAO Parameter.Type/enum constants and read-only/
  lifetime checks with real SQL tests; inference/Size/full coercion remain open.

- [x] Native VBA/Access-expression integer division and rounded Mod, correct
  precedence/sign/null/overflow behavior and fixture-independent Qt calculated-
  control/SQL/save tests; SQL rewriting and full numeric promotion remain open.

- [x] Native VBA/Access-expression power, documented left-to-right chains, unary
  precedence/Double/null/domain/overflow handling and actual Qt/VBA/SQL tests;
  SQL rewriting and full underflow/rounding/type/error identity remain open.

- [x] Native Xor/Eqv/Imp Boolean/null/bitwise execution, logical precedence and
  bounded subtype/range handling with actual Qt/VBA/SQL audit tests; SQL rewriting
  and full Variant/promotion/error/And-Or edge semantics remain open.

- [x] Native Fix/Sgn and checked Abs/Int null/subtype/overflow behavior, exact
  Currency truncation and actual Qt/VBA/SQL calculated-control regressions; full
  Variant/Byte/Date/locale/promotion/SQL function semantics remain open.

- [x] Native Ctrl+apostrophe previous-record TextBox value replacement through
  real edit/validation/undo events, typed NULL/empty, safe guards and SendKeys/
  Qt/SQL tests; exact Access edit-mode selection and other adapters remain open.

- [x] Native Ctrl+Plus/Ctrl+Shift+Equals new-row navigation and documented
  RunCommand 28/65-68, KeyPreview/validation/policy/no-INSERT/SendKeys Qt/SQL tests;
  full keyboard/layout/navigation/MDI/query semantics remain open.

- [x] Native Ctrl+Minus/SendKeys deletion through actual validation/callback/
  confirmation/savepoint behavior, Qt No/Yes/default-No and foreign-key tests;
  batch/datasheet/query/complete keyboard and Windows event identity remain open.

- [x] Long native TextBox values avoid Qt default truncation, with real SQL
  unchanged-field/VBA append/keyboard copy/Unicode-newline round trips; complete
  Memo multiline UI, formatting and Access field-size limits remain open.

- [x] Native modeless literal form search via Ctrl+F/acCmdFind, field/match/case/
  advance modes, draft/validation/lifetime/active-context and real Qt/SQL tests;
  full Find/Replace/wildcard/format/source/locale/MDI/UI identity remain open.

- [x] Native Shift+F4/acCmdFindNext/SendKeys repeat the last document-local literal
  search after closing its real dialog; validation, active context, wrapping and
  SQL regression pass; shared/global and complete Find/Replace identity remain open.

- [x] Checked native VBA Byte/CByte subtype, module/local/default/parameter/result/
  known-ByRef handling and real Qt/DAO Byte parameter SQL regression; full numeric
  promotion/Byte arrays/counter/locale/Windows identity remain open.

- [x] Typed native Not bit inversion for Byte/Integer/Long, Boolean/Null/Empty,
  scalar default values and checked numeric conversion; reproduced core/Qt/SQL
  regression fixed; complete mixed/text/locale/Jet SQL identity remains open.

- [x] Integer subtype for native VarType results, including nested inspection,
  Byte-array code, actual Qt calculated controls/Variant copies/SQL regression;
  general object inference/full Variant and array identities remain open.

- [x] Native VarType scalar-default capability for controls/DAO/Err, recursive/
  single-read/error propagation and actual Qt/SQL type-code/EOF/closed-object
  regression; full COM/class/default/type/lifetime/Windows identities remain open.

- [x] Native ASCII Asc/Chr and supported UTF-16 AscW/ChrW, checked errors/types,
  actual Qt/DAO/SQL and three unchanged original KeyPress procedures (384 direct/
  285 Qt cases); ANSI/DBCS/full BSTR/SQL/Windows/business identity remain open.

- [x] Restore declared Byte/Integer/Long field subtypes and bounded Field.Type
  from unambiguous import metadata, Qt/VBA/SQL bounds/pending/save tests and all
  127 original INT columns/471904 values; raw SQL/full Jet/type identity remains open.

- [x] Shared native VBA/SQL bankers Round, documented half cases, explicit fixed
  Currency, actual Qt/DAO/SQL/pending/save/rollback regression; Single/high
  precision/full IEEE/locale/Jet money/Windows identity remain open.

- [x] Bounded native Form.Visible hide/show with retained lifetime/drafts, hidden
  active-context/default-command protection, actual Qt/VBA/SQL/Unload/self-close
  regression; initialization/modal/dirty-hide/full Windows/MDI identity remains open.

- [x] Native Ctrl+Page Down/Up record navigation through validation/save, actual
  text/ComboBox focus, KeyDown/SendKeys/boundaries/Current self-close regression;
  multi-page/other views/full Windows keyboard identity remain open.

- [x] Native VBA/calculated Replace with UTF-16 suffix/count/compare, scalar-once
  checked arguments, actual Qt/DAO parameter/SQL pending/save/Undo regressions;
  full Windows collation/BSTR/error identity and Jet SQL Replace remain open.

- [x] Restore native Long results for supported Len/InStr and Integer for StrComp,
  checked range/Null, actual Qt/VBA/DAO parameter/SQL type-code/pending/save/Undo
  regression; numeric/UDT Len/full Windows/SQL subtype identity remains open.

- [x] Native IsObject with object-preserving parser context, no default getters,
  Nothing/closed-reference identity, explicit getter errors and actual Qt/VBA/DAO/
  SQL pending-source regression; full class/COM/ADO/Windows identity remains open.

- [x] Preserve Currency/Single/Double source labels from MDBTools Access schema on new
  imports, checked ambiguous mappings and persisted metadata; original 43 Currency/
  57 Double columns verified. Exact money/Single storage/SQL/runtime semantics remain open.

- [x] Bounded native direct-table Currency fields/dbCurrency with lossless REAL
  encoding inside +/-2^39, exact Qt text/DAO/pending/save/OldValue/Undo/copy and
  independent 121613 original values; full-range/SQL/query/Windows money remains open.

- [x] Bounded Currency/Money SQL parameters with fixed Value/dbCurrency, numeric
  binding, checked rounding/ranges/Null/retention/lifetime and actual Qt/VBA UPDATE;
  full-range/SQL money/LIKE/result typing/Windows identity remains open.

- [x] Explicit bounded native SQL CCur through shared conversion, decimal rounding/
  ranges/errors/aliases and real DAO atomic UPDATE/Qt/VBA requery regression;
  full query Currency subtype/SQL arithmetic/locale/Windows identity remains open.

- [x] Checked numeric Currency SQL Forms/Formulare/Reports/Berichte references,
  actual Qt/VBA/DAO numeric/pending/no-save/range regression; full object grammar/
  SQL Screen/locale/Windows identity remains open.

- [x] Proven root CCur/Currency-parameter query types through DAO, query-table
  display and readonly Qt controls/empty/filter/TOP/DISTINCT/saved queries;
  general SQL/view/alias/aggregate/arithmetic/Windows typing remains open.

- [x] Exact native sums of proven Currency expressions/parameters with grouped/
  Null/cancellation/overflow/range/Qt report regression and 43 independent original
  column sums; implicit field/view/complex/full SQL money/Windows semantics remain open.

- [x] Verified live Currency fields and exact implicit SUM in simple physical-
  table SELECT/DAO/readonly Qt paths, stale-plan/schema/filter safeguards and all
  43 original sums without casts; general/view/join/action/Windows typing remains open.

- [x] Exact proven Currency SELECT +,-,* with precedence/associativity/rounding/
  Null/parameter/SUM/DAO/Qt report/range/overflow regression; mixed/division/
  complex/action/full-range/Windows arithmetic remains open.

- [x] Simple physical-table Currency UPDATE target conversion and proven RHS
  arithmetic, parameter/alias/qualified/multi-assignment/NULL/range/atomic-error
  regression through actual DAO/VBA/Qt; joined/append/complex/full Jet semantics
  remain open.

- [x] Simple single-row physical-table Currency INSERT VALUES target conversion,
  explicit/positional/reordered columns, parameter/exact-expression/NULL/error/
  DAO/VBA/Qt regression; append/select/complex/full Jet semantics remain open.

- [x] Bound native Form.CurrentRecord as read-only Long through Me/Forms/Screen/
  unqualified VBA/calculated controls, navigation/filter/new-row/Current-event/
  pending-edit/lifecycle regression; unbound/design/full Windows behavior open.

- [x] Bounded physical-source Currency INSERT SELECT append with target
  conversion, frozen positional source, exact Currency self-append, parameters
  and whole-statement rollback; complex/full Jet append semantics remain open.

- [x] Shared native ISO/Gregorian DateDiff and Year/Month/Day/Hour/Minute/Second/
  Weekday in VBA and SQL, checked types/calendar boundaries/Null/errors and
  existing-test coverage; Date Variants/numeric/locale/full calendar semantics open.

- [x] Native Form.OrderBy/OrderByOn field-list sorting, table identity preservation,
  query ordering, imported defaults, validation/state rollback and existing
  navigation/filter/CurrentRecord regression; expressions/collations/full sorting open.

- [x] App table editor uses checked source types, exact Currency line editors,
  full long-Memo text, read-only/generated AutoNumbers and native defaults with
  real Qt/editor/persistence regression; full datasheet/constraints/formatting open.

- [x] App SQL editor uses the Access compiler/model, typed declared parameters,
  DAO action execution/affected rows and saved actions, verified through its real
  dialog; complete Jet/designer/editor parity remains open.

- [x] Native app Boolean table checkboxes with Space, pending edits and checked
  -1/0 persistence; full Yes/No defaults/Null/visual identity remain open.

- [x] Native table Ctrl+S completes/validates a live cell before transactional
  manual save, retaining invalid editor text; full per-record/undo/key parity open.

- [x] Pending app datasheets participate in tab/window/project/VBA-Quit close
  preflight, preserving cache/live editor and calling form; unchanged/reverted
  states close normally; full autosave/events/undo/window semantics remain open.

- [x] Native datasheet Ctrl+W/Ctrl+F4 guarded close, page-wide Ctrl+S and deferred
  shortcut-sender deletion with actual key regression; full window/key parity open.

- [x] Checked DAO Long RecordsAffected/RecordCount and Integer Fields.Count,
  zero DDL counts without stale SQLite changes, preserved DML/make-table counts
  and real editor feedback; full cursor/metadata/Windows semantics remain open.

- [x] Bounded simple SQL writes check verified Byte/Integer/Long/Boolean/date
  targets, shared rounding/ranges/canonical storage/Null/atomic late-row failure
  and unmarked-SQLite preservation; complex/full Jet coercion remains open.

- [x] Native one-dimensional Variant Array construction, Option Base, checked
  indexing/element replacement, bounds/type inspection and whole-array value/
  ByRef/ByVal semantics with real VBA-to-DAO persistence; full arrays remain open.

- [x] One-dimensional Variant ReDim/Preserve/Erase, Option Base and negative
  lower bounds, grow/shrink/value-copy preservation; typed/full array parity open.

- [x] Native Split/Join/Filter with String arrays, typed element replacement,
  zero-based bounds, limits/compare/Unicode and real VBA-to-DAO list persistence.
- [x] Bounded array For Each with Variant control, named sources/direct
  constructors, live later-element writes and value copies; full alias parity open.
- [x] Native VBA and SQL InStrRev/StrReverse, UTF-16 positions, defaults,
  projection aliases and checked errors; full BSTR/locale/Jet typing remain open.

- [x] Trailing Optional scalar parameters with explicit literal defaults,
  checked per-call binding, preserved supplied ByRef/ByVal and real DAO effects;
  Missing/named/positional-hole/full declaration parity remain open.

- [x] Native VBA Like wildcard/digit/list/range/negation matching with comparison
  precedence, Option Compare, Null and error 93; full locale/BSTR semantics open.

- [x] Native final ParamArray Variant packing, empty/multiple args, bounds,
  indexing/For Each/Join, ByVal forwarding and DAO effects; slot/reference writes,
  named/hole/Optional combinations and complete binder semantics remain open.

- [x] Optional typed defaults without literals, distinct Missing/IsMissing,
  metadata/Empty/Null separation, argument holes and non-aliasing missing forwarding.
- [x] Native VBA named Sub/Function calls with written evaluation order,
  reordered ByRef references, Optional holes and checked duplicate/unknown names.
- [x] Signature-bound named native DoCmd arguments with actual SQL, filtered
  child opening/control read/close; unsupported options and other methods stay open.

- [x] Native form/report text OpenArgs before opening events, Null/empty-text
  storage, instance/collection reads, independent lifetime and reopen; full
  Variant/mutability/data/window-mode semantics remain open.

- [x] Native DoCmd.RunSQL UseTransaction (named/positional/default), scoped
  action savepoint/cleanup, preserved outer DAO transaction and Null SQL errors;
  complete Jet partial-failure/transaction semantics remain open.

- [x] Native new-form acHidden opening, Forms/data/OpenArgs lifetime, initial
  Open/Load/Current without synthetic Activate and named close; full modes remain open.

- [x] Selected native builtin named signatures with defaults/overload/error
  handling and real DAO use; other signatures and full binder parity remain open.
- [x] Reference-backed Nz(Null) VBA Empty versus query zero-length String,
  context arithmetic/concatenation and explicit fallback preservation.

- [x] Native new-form acDialog caller wait through close/hide and cancelled
  Unload, with deferred document lifetime; full Quit/desktop semantics remain open.
- [x] Native initial hidden reveal and guarded re-hide through Visible/OpenForm;
  full pending-edit/lifecycle/Windows visibility behavior remains open.
- [x] Imported/native Modal and PopUp reads with acDialog overrides; runtime
  mutation and complete window-frame/property semantics remain open.
- [x] Native acIcon and active-document Minimize/Restore/Maximize Qt states,
  including a caller distinct from the target; full window/desktop parity open.

- [x] Accepted native application Quit stops active nested VBA/expression/dialog
  frames without On Error bypass; rejected Quit remains handleable (2501).
- [x] Native construction scope prevents later event/show work after startup
  Quit in Open/Load, including UI opening without an outer procedure.
- [x] Native Variant/Byte array function results retain typed values rather than
  forced legacy scalar text; full array/SCODE/return-binding parity remains open.
- [x] Converted-form opener in mainwindow uses independent native nonmodal opening,
  retaining source Modal input blocking without wrongly pausing normal VBA calls;
  complete standard-module application/DoCmd object coverage remains open.


## Additional bounded native runtime increments (2026-11-03)

These entries record specific implemented/tested subsets, not equal-sized parts
of complete Access parity. Existing broader open items remain open.

- [x] Named arguments for implemented DAO Database/QueryDef/TableDef/Find methods;
  unsupported parameter families remain explicit.
- [x] Native Recordset.Move, AbsolutePosition, EditMode and Updatable on supported
  cursors, with checked position/change-state behavior and real DAO regressions.
- [x] Cursor-owned LastModified and Integer primary-key rekey propagation through
  supported DAO/form snapshots and bookmarks; external/trigger edges remain open.
- [x] Read-only DAO result Field.OrdinalPosition with actual projection order;
  full persistent catalog mutation remains open.
- [x] Direct-table DataEntry sessions and fresh acFormAdd/acFormEdit/acFormReadOnly
  opening policies, real inserts and NewRecord Integer semantics; queries open.
- [x] Native command-button Default/Cancel keyboard dispatch, focused editor undo
  and runtime uniqueness; full Access keyboard/event oracle remains open.
- [x] Imported attached-label mnemonics and guarded control SetFocus validation;
  broader mouse/accessibility/focus transition parity remains open.
- [x] Preserving existing String/Byte arrays in Variants, zero-based VBA.Array
  and optional Let assignment; full typed/multidimensional arrays remain open.
- [x] Bounded hexadecimal/octal literals and checked 32-bit numeric And/Or;
  full decimal/subtype/Null operator matrix remains open.
- [x] Native character/reverse-search SQL bridge, corrected InStrRev string Null,
  and bounded Space/String/Hex/Oct VBA/SQL functions; ANSI/full BSTR remain open.
- [x] Native VBA Sqr/Exp/Log/Sin/Cos/Tan/Atn with Double results and runtime errors;
  full Windows floating-point oracle remains open.
- [x] Native datasheet Ctrl+Plus/Ctrl+Minus through validated model operations,
  real-key SQL persistence and invalid-editor guards; auto-save semantics open.

- [x] Native Val decimal prefixes, stripped ASCII spaces, signed hexadecimal/octal
  patterns and checked suffixes in VBA/SQL; full suffix/locale matrix remains open.
- [x] Bounded LCase/UCase and dollar variants with correct Null, scalar and SQL
  behavior; non-ASCII Windows locale conversion remains explicitly unsupported.
- [x] Gregorian ISO DatePart intervals, explicit weekdays and Jan-1 week numbering
  in VBA/SQL; Date Variants, locale/Hijri and other week rules remain open.
- [x] Eager native Switch/Choose selection, typed VBA result preservation, Null,
  rounded indexes and SQL scalar execution; full object/coercion edges remain open.

- [x] Bounded native Date Variant, CDate, Date locals/module fields/parameters/
  results, type inspection and Gregorian calendar integration; full locale/OLE open.
- [x] Native Date clocks and supported bound-field/DAO Date parameter roundtrip,
  with physical-column Date lineage in query results; full expression typing open.
- [x] Calculated Date display and typed unbound Date storage; original calculation
  regression retained, complete formatting/precision/edit-event oracle remains open.
- [x] Guarded Window menu next/previous/database/Close All and numbered visible
  form/report list with stale-entry protection; full MDI/datasheet list remains open.

- [x] Native DateSerial/TimeSerial and DateAdd with checked arguments, relative
  normalization/month clamping and typed SQL projections; region/Hijri remain open.
- [x] Typed Date SELECT addition/subtraction and numeric Date-difference SUM,
  with metadata/Null/nesting checks; full SQL clause/operator coverage remains open.
- [x] Date Variant overflow promotion, numeric bound-Date assignment and explicit
  subsecond storage refusal without mutation; full Date coercion matrix open.
- [x] Native unbound calendar Date values and original focused-target popup
  transfers, plus civil Date input across system DST gaps; full ActiveX open.

- [x] Native procedure colon chains and qualified implemented VBA intrinsics,
  preserving Date literals, named arguments and local VBA object precedence.
- [x] Native DoEvents dispatch and Integer-zero result with retained document
  lifetime and nested Quit noncontinuation; full Windows message ordering open.

- [x] Native VBA Timer wall-clock Single result, qualified/bare/parenthesized
  calls and arity checks; Windows tick-resolution equivalence remains open.

- [x] Native ordered Collection construction, typed local/default/keyed access,
  Add/Remove/Count/named insertion, Variant/object identity and leased For Each;
  regional keys, cyclic lifetime and complete class semantics remain open.

- [x] Typed Collection module persistence, required ByVal/ByRef parameters and
  function results, shared mutation/local replacement and caller rebinding;
  Optional object parameters and complete class/reference semantics remain open.

- [x] Native form Ctrl+semicolon / Ctrl+Shift+colon date/time insertion through
  edit buffers, Change, validation and real Date persistence; regional masks open.

- [x] Native editable datasheet text-cell date/time shortcuts, starting edit or
  replacing a live selection, with real Ctrl+S persistence; regional masks open.
- [x] Existing keyed datasheet cell restored after save/model reset and sort
  movement, with composite-key comparison and bounded search; new/no-key open.

- [x] Optional Collection Nothing defaults, typed IsMissing and supplied/named/
  omitted ByRef arguments; generic Optional object binding remains open.
- [x] Native main-window text editing menu with actual undo/redo/clipboard
  commands and protected/password states; shared MDI/range operations open.

- [x] Shared acyclic nested Collection/Variant-array graphs and nested For Each,
  with indirect cycle refusal before mutation; cyclic object lifetime stays open.

- [x] Native Unicode-string LenB UTF-16 byte counts, Null/Empty, qualified calls
  and bound SQL/alias execution; numeric/UDT storage layout remains open.

- [x] Native datasheet row/column/all keyboard selection and checked rectangular
  text copy, full bounded fetch, Currency/number precision, quote handling and
  rejection before clipboard replacement; typed MIME/range paste remain open.

## Bounded datasheet range paste and retryable Save (2026-11-03)

- [x] Strict quoted TSV into an existing clean manual-submit rectangle, with typed preflight, bounded fetching and no automatic SQL submission.
- [x] Transactional native table Save retains every draft after a later SQL constraint failure; correction/retry includes updates, inserts and deletes.
- [ ] Clipboard record append, paste over existing drafts, typed binary MIME, regional masks and complete Access range undo.

See DATASHEET_CLIPBOARD_RUNTIME.md for implementation bounds and verification.


## Native datasheet field navigation and Zoom (2026-11-03)

- [x] F2 switches native table cells between caret editing and field navigation, retaining invalid input until correction.
- [x] Shift+F2 native multiline text Zoom for table cells and their text editors; Cancel, untouched NULL/line endings, readonly display and manual-submit drafts.

See KEYBOARD_RUNTIME.md; complete Access keyboard/display parity remains open. No complete product coverage follows from checklist counts.


- [x] Native TextBox Zoom preserves form edit buffers, Change, readonly display and cancelled BeforeUpdate before actual SQL saving.
- [x] DoCmd.RunCommand acCmdZoomBox=179 resolves the active native form control, with error 2046 when unavailable.

- [x] Application.RunCommand forwards the existing implemented native command subset with named Command binding and original error propagation.


## Native table data integrity (2026-11-03)

- [x] A successful commit followed by metadata/query refresh failure retires drafts, reports committed status and never repeats an already persisted insert on Save retry.
- [x] REAL/FLOAT/DOUBLE text editors preserve supported Double precision; untouched nullable REAL/Currency/Date cells remain SQL NULL.

See DATASHEET_CLIPBOARD_RUNTIME.md; exact Decimal and complete Windows formatting
remain open. Both data-loss paths were reproduced on synthetic working data.


- [x] Native Ctrl+End fetches complete supported datasheets before last-record navigation; Ctrl+Home returns to the first field.
- [x] Native Tab/Backtab/Shift+Tab/Return/Enter retain invalid input and validate successful edits into the cache without SQL autosubmit.


- [x] Rejected native table input survives normal editor-close/mouse field changes, restoring current cell and focus; Escape still discards explicitly.
- [x] Native datasheet context menu dispatches checked Copy/Paste/Zoom and record/column/all selection with actual enabled-state guards.


- [x] Native Window menu lists and activates the actual open table tabs alongside native Form/Report documents.
- [x] Shared project Ctrl+F6/Shift+Ctrl+F6 ring includes real table pages, with source validation, modal guards, tab-click refusal and retained manual-submit drafts.


- [x] Reopening a native table reuses the existing page/model and preserves pending drafts without SQL submission.
- [x] DoCmd.SelectObject acTable/name/False activates an actually open native table while retaining the caller's Me and independent lifetime; unsupported closed/query contexts stay explicit.


- [x] Native DoCmd.OpenTable opens actual project datasheets in default/acEdit and acReadOnly modes, preserving explicit unsupported view/mode boundaries.
- [x] Read-only table model guards block native writes/insert/delete/paste/submission and checkbox mutation; dirty policy changes preserve drafts.
- [x] OpenTable native method binding supports named/omitted arguments and distinguishes Missing defaults from explicit Empty enum coercion.


- [x] DoCmd/Application RunCommand SaveRecord commits the actual active native table through the shared UI transaction/selection helper, including specialized interpreter dispatch.
- [x] A form initialization Save keeps its own context and leaves a dirty background table/SQL unchanged until an explicit native table switch/save.

- [x] Explicit/default native table DoCmd.Close validates fields and transactionally saves record drafts, including acSaveNo; failures retain the real page/editor/cache.
- [x] Parent-hidden active table editors remain checked, and reentrant table Close terminates its owning native context menu safely.

- [x] Native table acCmdSelectRecord validates the active field and selects the real row while preserving readonly access and caller Me.
- [x] Native table DoCmd/Application acCmdCopy uses the actual checked TSV selection, including readonly and dirty cached rows without SQL submission.

- [x] Native table acCmdZoomBox uses checked typed cached edits and readonly viewing; Cancel completes without mutation and SelectRecord preserves its active field.

- [x] Bounded named/default table GoToRecord and native First/Previous/Next/Last use actual view order and fetch SQL rows beyond the initial Qt batch.
- [x] Table navigation transactionally saves a sole changed existing current record with primary-key restoration; SQL failure retains the draft for retry.
- [ ] General table navigation with unrelated multirow drafts, inserted/deleted current records, and Access addition-row/error equivalence.

- [x] Native table RunCommand SortAscending/SortDescending sorts the actual field, preserves primary-key/current-column identity and readonly access, saves a sole current update and protects unrelated drafts.

- [x] Native table literal Find handles visible text-field scope, part/whole/start and case modes, real SQL fetching and retained dirty-cache boundaries.
- [x] Ctrl+F, Shift+F4, native Edit menu and DoCmd/Application Find/FindNext share actual table search paths, including the cell-editor route.
- [x] Modeless table Find reactivates its owning page safely and is disposed with its document; checked view errors propagate accurately to VBA.
- [ ] General table formatted/wildcard/locale/all-field search, remaining typed coercion/partial matching, Replace and DoCmd.FindRecord equivalence.

- [x] Bounded DoCmd.FindRecord searches actual current native table text fields with named/Missing/Empty argument binding, first/next origin and normal no-match completion.
- [x] Native form FindRecord uses the actual active bound text field, readonly access and checked validation cancellation (2501), preserving caller Me and pending drafts.

- [x] Native table/form Find dialogs and FindRecord support Up/Down without wrap, SearchAll with wrap, preserved FindNext direction and explicit first/next origins.

- [x] Native table/form dialog and FindRecord exact numeric searches retain adjacent Double distinctions without display-string comparison.
- [x] Native table/form Currency searches compare checked four-decimal units through a lossless typed invocation and real field conversion.
- [x] Native table/form Date searches use checked second-resolution canonical Date values, including negative OLE aliases and explicit subsecond refusal.
- [x] Native table/form Boolean search and FindNext skip SQL NULL; readonly typed dialogs and invalid-input retention use the real data path.

- [x] Native table SendKeys resolves actual cell/Find/Zoom focus per key, retains project identity and ordered modal/queued delivery without old-caller fallback.
- [x] DoCmd.SendKeys shares global native delivery, named/omitted Wait binding, DoEvents queue draining and checked Text/Value/SQL edit boundaries.

- [x] Native table DoCmd/Application acCmdSelectAllRecords shares bounded full-fetch Ctrl+A selection, validates the active editor and preserves the current field without submitting drafts.

- [x] Native table acCmdPaste routes real clipboard TSV through atomic typed preflight/manual-submit cache; invalid later cells, readonly targets and existing drafts retain checked errors without SQL writes.

- [x] Active-table DoCmd.Requery and Shift+F9 reload the actual SQL source, save a sole current update through checked transactions and protect unrelated drafts.
- [x] Native table Shift+Return and Shift+keypad Enter save through active-editor validation and the transactional Save path without losing the current primary-key/field selection.

- [x] Native table acCmdRecordsGoToNew creates a real insert draft after checked sole-current saving, including empty tables; existing insert/delete/unrelated drafts are protected.
- [x] Native table acCmdDeleteRecord uses checked single-row deletion; shared button/Ctrl+Minus refuse multirow selections before mutation, with actual transactional DELETE on Save.
- [x] Successful native INSERT captures actual generated SQLite primary keys inside the transaction and restores the saved keyed row/field after model reset without guessing IDs.

- [x] Named/default native table GoToRecord acNewRec shares checked record insertion, including named-page activation, empty tables, readonly protection and existing-draft refusal.
- [x] Native table insertion fetches all existing source batches within the explicit row limit before appending its draft; SQL remains unchanged until Save.

- [x] Native table acCmdPaste resolves actual owned QLineEdit cell/Find focus, pastes into the real text buffer and preserves checked field/SQL boundaries, including readonly-source Find queries.

- [x] Native acCmdCopy/acCmdCut use the actual owned cell/Find QLineEdit selection, preserve buffer/SQL boundaries and readonly-query editing, and protect missing selections/masked text/clipboard limits.
- [x] Native table acCmdUndo/acCmdRedo share the focused QLineEdit history with actual shortcuts, preserving buffered field validation and explicit missing-history/readonly/grid boundaries.

- [x] Native table grid Ctrl+Insert/Shift+Insert share checked readonly TSV copy and typed draft paste with Ctrl+C/Ctrl+V; text editors retain native Qt shortcuts.

- [x] Native CurrentProject/Application.CurrentProject expose real readonly Name/Path/FullName, per-connection identity and independent caller lifetime with checked project-generation invalidation and preserved user-name lookup.

- [x] CurrentProject.AllForms/AllReports query saved native/imported catalogs with checked Count/Item/named lookup/For Each, readonly descriptor names, live scope and explicit ambiguity/version rules.
- [x] Project catalog IsLoaded tracks actual hidden/modal/embedded previews and visible native runners/designers, excludes closed owners, and invalidates independent item references on project replacement.
- [x] Native report Visible getter/setter uses checked Qt document visibility, preserves loaded readonly reports and executes real activation/deactivation VBA/SQL with continued caller code.

- [x] CurrentData/Application.CurrentData expose real main-schema user tables and saved query catalogs through checked Count/Item/named lookup/For Each and readonly AccessObject metadata.
- [x] CurrentData catalogs track background/readonly/closed native datasheet tabs and actual named query dialogs, retain independent held references and invalidate them on accepted project replacement.

- [x] Explicit readonly normal-view DoCmd.OpenQuery opens real native/imported SELECT result tabs with checked named/Missing arguments, typed parameter prompts and no silent editable/action-query substitution.
- [x] Native query tabs outlive their opener, retain actual CurrentData IsLoaded/window focus, reuse existing documents and close by named/default DoCmd.Close or Ctrl+W/Ctrl+F4.

- [x] Active readonly query Requery/Shift+F9 reload saved SQL and typed parameters, preserve old results on failure/cancel and prevent closed-document resurrection and same-query recursive execution.

- [x] Readonly query documents support actual default/named acDataQuery record navigation, checked bounds/offsets, first/previous/next/last commands and readonly new-record rejection.
- [x] SelectObject acQuery activates existing named native query tabs case-insensitively with document validation, independent caller identity and no silent reopen/requery.
- [x] Query Copy/Ctrl+C/Ctrl+Insert uses real typed result cells and shared bounded TSV quoting; hidden/disjoint/binary/unrepresentable selections retain existing clipboard contents on failure.
- [x] Native query SelectRecord/SelectAllRecords and scoped SendKeys address actual selected result cells and focused controls without falling back to the calling form.

- [x] Native readonly query Find/FindNext/Ctrl+F/Shift+F4 and current-field FindRecord search real typed results, preserve unknown-type boundaries and focused search text commands, and support immediate VBA search after OpenQuery.

- [x] Native query Edit-menu Copy/Select All forwards actual view commands with readonly Cut/Paste disabled, shared checked formatting and selection limits before per-cell allocation.
- [x] Query acCmdZoomBox/Shift+F2 opens an actual readonly Qt cell Zoom dialog with checked text, no SQL/model mutation and guarded target lifetime.

- [x] Three-part saved-query/relation/field references bind through proven direct view projections and ranked wrappers to exact output labels, distinguishing same-table input views and refusing ambiguous/compound/overqualified origins.

- [x] Native multidimensional Variant values support checked indexing/assignment, dimension-aware Long bounds, value copies and first-dimension-fastest For Each without flattening Join/Filter inputs.
- [x] Variant ReDim supports up to 60 dimensions with Option Base per omitted bound; Preserve grows/shrinks only the last upper bound and retains existing values/copy isolation.
- [x] Explicit-positive-NumRows DAO GetRows returns real zero-based typed field/record matrices, including native binary cells, with next-unread cursor advance and conversion-failure position preservation.

- [x] ReDim As primitive type in existing Variant containers retains actual typed defaults/coercion/metadata and Preserve type guards, while full typed declarations and unsupported intrinsic array-kind paths remain explicit.

- [x] Set on Variant/Object-array cells preserves actual native references/Nothing, independent container copies and shared object identity; ReDim As Object defaults and Preserve use checked native binding.
- [x] Erase of Byte arrays inside Variant variables releases the container to Empty with supported ByRef propagation while independent byte copies retain contents/bounds.

- [x] Default/acEdit OpenQuery supports real keyed single-table SELECT-star datasheets, record save/validated close/requery and independent caller lifetime.
- [x] Editable query documents retain query identity and reject competing writable source tabs, changed saved definitions and source-schema/shadow changes without discarding drafts.

- [x] Editable single-table queries retain bounded compiled WHERE predicates in actual SQL refresh/save, including checked disappearance after a committed value leaves the filter.
- [x] Editable query source keys require structural non-null identity, distinguishing real INTEGER rowid aliases from nullable or DESC indexed primary keys.

- [x] Local declared dynamic VBA arrays retain primitive/Object/Variant type through unallocated state, checked ReDim/Preserve, compatible copies and Erase; ByRef/fixed/class boundaries remain explicit.
- [x] Declared dynamic module arrays persist real typed values across procedure invocations and retain their declared type after deallocation.

- [x] Required declared ByRef array parameters perform real matching-type caller writeback for element edits, Preserve resize and Erase, with checked mismatch refusal.
- [x] Whole-array VBA a() expressions/calls preserve array reference metadata; direct declared-array ByVal arguments refuse while explicit Variant copies remain independent.

- [x] Native VBA FileLen reads actual UTF-8 absolute/relative Linux file sizes as checked Long values, with real sparse-file overflow and path error handling.
- [x] Native VBA FileDateTime returns typed actual last-write Date values in local time, with named/qualified calls and timezone-checked real file metadata.

- [x] FileLen/FileDateTime execute through real parameterized native SQL callbacks with volatile metadata refresh, canonical integer/date projection types and checked filesystem error propagation.
- [x] DAO recordset Fields expose actual SQL file metadata as Long/Date values and dbLong/dbDate types; proven DATE/DATETIME Field.Type returns dbDate.

- [x] Literal-bound local/module fixed VBA arrays allocate real typed dimensions with Option Base and retain bounds/type when Erase resets cells to native defaults.
- [x] Fixed array binding survives whole-array ByRef argument forwarding, allowing element/reset writeback while guarding ReDim, incompatible shape and whole-array replacement.

- [x] Native scalar local/module Const declarations compose checked pure expressions, retain typed immutable bindings and pass temporary values through supported ByRef calls.
- [x] Fixed array dimensions fold integral constant expressions and prior module/local constants without executing runtime APIs.
- [x] Native VBA character constants include Windows-target vbNewLine/CRLF and checked empty/NUL/control-character String values without claiming null-BSTR interoperability.

- [x] Optional scalar parameter defaults fold pure expressions and prior same-module constants into immutable metadata, preserving named overrides, explicit Empty/Null/Missing distinctions and independent ByRef default values.

- [x] A bounded per-thread native VBA translation LRU reuses source-only plans with exact-source/context keys, preserves fresh host/module/argument binding and leaves large plans uncached.

- [x] Native Scripting.Dictionary construction supports real ordered key/value operations, bounded comparison/key families and cross-container cycle guards without generic COM activation claims.
- [x] Dictionary Keys/Items expose independent owned snapshots and object For Each retains its enumerated owner through source-variable clearing and loop exit.
- [x] Indexed object default writes support Dictionary values/references and real DAO Recordset/Fields pending edits committed only through Update.

- [x] Scalar module Object fields initialize to Nothing and retain actual native object identity across calls, with checked reference binding, visibility metadata and independent host state.

- [x] Native As Object argument/Set/collection-loop binding checks actual references; ByVal/ByRef replacement and independently omitted Optional Nothing parameters execute through real native objects.

- [x] Native DAO identities support checked qualified module/local/argument/result bindings and Nothing defaults, retaining actual SQL cursor/object state across calls and refusing mismatched replacements.

- [x] Explicit procedure-local Static declarations retain native scalar/object/dynamic/fixed-array state by host/procedure/source through repeated and recursive calls; the broader Static-procedure modifier remains explicitly unsupported.

- [x] Native datasheet Ctrl+H/menu/acCmdReplace opens real whole-text-field replacement; transactional Save/rollback drafts, readonly and primary-key/source/record validation guards apply.

- [x] Native datasheet Esc/context action discards only current-record pending changes, preserves other drafts, cancels pending inserts and leaves saved SQL intact; first editor Esc retains existing model drafts.

- [x] Qualified native Scripting.Dictionary local/module/parameter/function-result bindings retain actual checked references and optional Nothing defaults without a COM/type-library dependency.

- [x] Declared native object value/Let assignments route actual Field/Parameter/Control scalar defaults to checked Value setters while preserving references and real buffered SQL/form update behavior.

- [x] Native For Each enumerates packed/multidimensional Byte arrays as actual Byte Variants, preserving live supported element edits, copy isolation and source-allocation/replacement guards.

- [x] Multi-target Erase retains one source-mapped instruction, ordered native dynamic/fixed/Variant reset behavior, constant preflight and statement-scoped On Error Resume Next behavior.

- [x] Multi-target ReDim/Preserve uses one source-mapped native instruction, publishes successful module/ByRef bindings before dependent later bounds and retains checked type/copy/error/constant behavior.

- [x] Bare Up/Down in native datasheets use the existing validated current-record Save/navigation path, preserve columns and retain current drafts/selection after actual SQL failure.

- [x] Tab/Backtab to adjacent existing datasheet rows, including delegate editor completion, use checked record Save/navigation and retain field destinations or pending drafts after SQL failure.

- [x] Native Ctrl+Home/Ctrl+End use checked boundary record navigation and visible-field order, preserving pending edits for same-known-record field movement and retaining failed SQL drafts.

- [x] Native loaded-page PageUp/PageDown pass actual viewport record distances through checked Save/navigation, retain failed SQL drafts/current rows and avoid SQL Save on zero-distance page movement.

- [x] Native Next counter lists preserve matched loop partners, zero/Exit For and mixed For Each execution, source-group Resume Next boundaries and released abandoned enumeration state.

Literale Tabellen-Standardwerte werden jetzt aus dem MDB-Export erhalten und
in SQL, DAO-AddNew und gebundenen Formular-Neudatensätzen verwendet:
[SOURCE_LITERAL_DEFAULTS.md](SOURCE_LITERAL_DEFAULTS.md). Diese Erweiterung gehört
zum bestehenden offenen Umfang F015 und erhöht weder456 IDs noch die384 begrenzt
bestätigten Einträge. Allgemeine Defaults/Eigenschaften bleiben offen.

Native scalar SQL/VBA standard-module calls now execute real SELECT/UPDATE work
with shared scalar module/Static state and checked source/connection generations.
Private/class/ambiguous calls, GUI/DAO host context, existing object/array state and
unproven transports remain refused. See NATIVE_VBA_SQL.md. F040/F068/F121 are still
open; no new checkbox or ID. Catalog456/384/72 unchanged. Production code full48/48
(87.35 s), additional actual UPDATE DAO1/1 (0.74 s), package/backend/four hashes pass.

Native compiler declarations now drive supported SQL/VBA Currency/Date return
transport and stale-signature guards. Actual DAO types/values, typed arithmetic/
INSERT and side-effect-free metadata preparation passed. General Variant/nested/
compound typing and typed-function make-table metadata remain open. Full48/48
(85.94 s), fresh core1/1 (2.36 s), final label DAO/query/original3/3 (8.36 s), updated
package/backend/four hashes pass. No new checkbox or ID;456/384/72 unchanged.

Typed make-table milestone (2026-11-05): full current desktop48/48 passed
(87.98 s); Debian payload/build-ID/pinned SQLite3.54.1 and all four original
fixture hashes passed. GUI-free core runtime unchanged from fresh1/1 (2.36 s).
Original-query coverage remains78/100 without Access result oracle. No new
CTest target or catalog ID; fixed456/384/72. Native SQL scope: NATIVE_VBA_SQL.md.

Native unqualified FileCopy now copies real binary files and replaces existing
regular destination files. Missing-source failure preserves the destination;
identical files/final symlinks/Windows drive mapping refuse3251. No full Windows
sharing/error/metadata equivalence claim. Scope: docs/NATIVE_DIRECTORY_STATEMENTS.md.
Existing desktop core1/1 (1.77 s), fresh headless core1/1 (1.69 s), latest Debian/
build-ID/SQLite3.54.1/four fixture hashes pass. Last full typed make-table48/48
(87.98 s); no new target or fixed ID,456/384/72 unchanged.

Value List ComboBox/ListBox AddItem(Item,Index)/numeric RemoveItem(Index) now
mutate real Qt models, including named VBA calls, preserving clean record/bound
value. Stable Form context only; active popup/selection, headers/ordinal binding,
selected combo-row removal, text-key removal and empty/Null-only edge cases remain
bounded refusals. Full scope: docs/COMBO_LIST_PROPERTIES_RUNTIME.md. Final existing
layout1/1 (10.66 s), current Debian/build-ID/SQLite/four hashes pass; headless runtime
unchanged FileCopy1/1 (1.69 s), last full48/48 (87.98 s). F162 stays open;456/384/72.

SQL/VBA registry now reserves source aliases of registered linaccess_jet_ adapters
as well as actual SQLite names. This prevents public VBA FV As Currency from
providing false Currency metadata to builtin SQL FV. Actual builtin result20
and absent false Currency proof pass in existing DAO/query/original3/3 (7.79 s);
Debian/build-ID/SQLite/four hashes pass. Core runtime unchanged FileCopy1/1
(1.69 s), last full48/48 (87.98 s). Complete name-priority oracle remains open.
Scope: docs/NATIVE_VBA_SQL.md; fixed456/384/72 unchanged.

Native NPer(Rate,Pmt,Pv,[Fv],[Type]) now computes real fractional financial
periods in VBA and SQL with qualified/named VBA calls, zero-rate balance and
stable logarithmic nonzero-rate calculation. No finite unique solution5;
bounded rate/timing3251; full Windows numeric/error oracle remains open. Scope:
docs/NATIVE_FINANCIAL_FUNCTIONS.md. Existing core/DAO/query3/3 (2.76 s), fresh
headless1/1 (1.77 s), updated Debian/build-ID/SQLite/four hashes pass. Last
broader full48/48 (87.98 s); no new target/ID,456/384/72 fixed.

Native financial expansion (2026-11-05): NPV(Rate,Values) and
MIRR(Values,Finance_rate,Reinvest_rate) execute real one-dimensional Double arrays
with arbitrary lower bounds, mixed cash-flow signs and bounded finite rates.
NPV first-flow period1, MIRR first-flow period0; no array flattening or SQL array
transport. Other element/rank/empty/single-sign NPV semantics stay bounded3251.
DDB executes real VBA/SQL declining-balance depreciation with default factor2
and actual salvage cap, whole life/period up to32767. No full fractional/asset
range or Windows error/numeric oracle claim. Scope: docs/NATIVE_FINANCIAL_FUNCTIONS.md.
Current full desktop48/48 (85.94 s), fresh GUI-free core1/1 (1.67 s), updated
Debian/build-ID/pinned SQLite3.54.1/four original fixture hashes pass. Original
corpus78/100; no new target or ID. F068 stays open, fixed456/384/72.

Native Rate/IRR now share an actual balance/analytic-derivative root kernel with
default guess0.1, bounded twenty iterations, residual validation and explicit
nonconvergence5. Rate supports VBA/SQL whole annuity periods1..32767; IRR supports
actual one-dimensional Double arrays with documented Values/Guess names and
positive/negative flows. No direct SQL array transport or full Windows root/
error/precision/selection oracle claim. Scope: docs/NATIVE_FINANCIAL_FUNCTIONS.md.
Existing core/DAO/query3/3 (2.52 s), fresh headless1/1 (1.73 s), updated Debian/
build-ID/SQLite/four hashes pass. Last broader full is preceding financial
milestone48/48 (85.94 s). No new target/ID; F068 open,456/384/72 fixed.

Native IPmt/PPmt now calculate actual interest/principal portions in VBA/SQL,
sharing stable annuity factors with existing FV/PV/Pmt and advancing balances
through earlier periods. Beginning/end timing, first beginning-payment interest0,
zero-rate principal and actual SQL credit rows pass. Whole periods1..NPer up to
32767, rate greater than-1 and timing0/1 only; full Windows rounding/error/domain
oracle stays open. Scope: docs/NATIVE_FINANCIAL_FUNCTIONS.md. Final existing
core/DAO/query3/3 (2.68 s), fresh headless1/1 (1.79 s), updated Debian/build-ID/
SQLite/four hashes pass. Last full is fb63120 financial milestone48/48 (85.94 s),
preceding roots/components. No new target/ID; F068 open,456/384/72 unchanged.

Calculated query headers now preserve native source-expression labels before
scalar/concat/backend helper rewriting. Explicit/implicit/plain/quoted aliases,
mixed-star projections and first-leg UNION names stay supported. Internal helper
names are removed from supported generated headers. No Windows Expr1000, duplicate
numbering or complete alias-grammar claim. Scope: docs/QUERY_COLUMN_LABELS.md.
Final full current desktop48/48 (86.70 s), updated Debian/build-ID/pinned SQLite/
four original hashes pass. Qt-only label change leaves GUI-free runtime unchanged
from fresh payment-components1/1 (1.79 s). Original78/100; no Access oracle.
No new target or catalog ID; F068/F080/F148 open,456/384/72 fixed.

Native table Ctrl+H now offers real atomic Alle ersetzen for literal whole-field
text in a visible non-key column of a keyed writable table. Current filtered/
sorted rows are fetched; Qt Unicode case matching, NULL preservation and current
primary-key restoration are implemented. Later SQL failure rolls back the batch
and clears only its own staged replacements; qualified base revertRow bypasses
the saving guard that previously blocked revertAll. Unrelated drafts are refused
and preserved. No full key/query/wildcard/locale/concurrency/Undo/Access oracle
claim. Scope: docs/TABLE_REPLACE_ALL_RUNTIME.md. Final existing mainwindow1/1
(2.92 s), current Debian/build-ID/SQLite/four hashes pass. Last broader full is
preceding labels48/48 (86.70 s), unchanged fresh headless runtime1/1 (1.79 s).
No new target/ID; F373/F160/F168 stay open,456/384/72 fixed.

Native Ctrl+Minus/DeleteRecord now stages multiple completely selected stored
table records, including noncontiguous rows. Partial-cell selection, mixed drafts,
keyless/new rows and reentrant/context changes stay refused. Persistent anchors
and weak view/model guards preserve identity/lifetime. Existing Save commits the
set atomically; a later FK failure retains all SQL rows and pending choices.
Native VBA SelectAllRecords/DeleteRecord now executes the supported selection.
No full Access confirmation/event/Undo/concurrency/selection oracle claim. Scope:
docs/TABLE_MULTI_DELETE_RUNTIME.md. Final targeted mainwindow1/1 (2.80 s), full
current desktop48/48 (84.19 s), updated Debian/build-ID/SQLite/four hashes pass;
core unchanged fresh GUI-free1/1 (1.79 s). No new target/ID; F168/F148 open,
fixed456/384/72.

Native Name sourceExpression As destinationExpression now uses actual Linux
renameat2/RENAME_NOREPLACE for files/directories within one filesystem. No target
overwrite fallback; existing target58. String/bracket/parenthesis-aware As
parsing and local keyword-named assignments are preserved. Source links/special
files, current-directory ancestors, cross-filesystem/Windows mapping and full
sharing/error oracle stay bounded. Scope: docs/NATIVE_DIRECTORY_STATEMENTS.md.
Final desktop core1/1 (1.78 s), fresh GUI-free core1/1 (1.69 s), updated Debian/
build-ID/SQLite/four hashes pass. Last broader full is preceding multi-record
48/48 (84.19 s); no new target/ID, F040 open,456/384/72 fixed.

Native Scripting.FileSystemObject now supports real FileExists, FolderExists,
BuildPath and GetAbsolutePathName through local CreateObject/New and qualified
typed locals/parameters/results. TypeOf uses explicit native identity; wrong
classes13 leave the existing target intact. Scripting type validation/defaults
share a closed canonical family rather than guessing identity from TypeName.
No retained host/database handle, fabricated filesystem state, general COM/
reference resolution, drive/UNC/wildcard mapping or full FSO contract. SQL global
object-state guards remain strict. Scope: docs/NATIVE_FILE_SYSTEM_OBJECT.md.
Final full desktop48/48 (85.42 s), fresh headless1/1 (1.69 s), updated Debian/
build-ID/SQLite/four hashes pass. Original78/100; no Access result oracle.
No new target/ID; F040/F042/F091 open,456/384/72 fixed.

Native FSO GetFileName/GetBaseName/GetExtensionName/GetParentFolderName now
perform actual lexical extraction for bounded Windows/Linux paths, including
UTF-8 and nonexistent paths, without filesystem access or drive mapping.
Unvalidated root/dot/trailing-separator edge cases remain explicit3251. Scope:
docs/NATIVE_FILE_SYSTEM_OBJECT.md. Desktop core1/1 (1.73 s), fresh headless1/1
(1.72 s), current Debian/build-ID/SQLite/four hashes pass. Last broader full
48/48 (85.42 s) precedes this increment. No new target/ID; F091 remains open,
fixed456/384/72.

Native FSO GetFile(FileSpec) returns an independent real File object with
normalized absolute Path, Name and live Size. Actual regular Linux files only;
no host/stream/database retention. Long sizes and larger exact Double sizes
are distinguished; above exact range6. Symlinks, mutations, typed File binding,
full sharing/race/Windows metadata contract remain bounded/unimplemented.
Scope: docs/NATIVE_FILE_SYSTEM_OBJECT.md. Actual binary/UTF-8 metadata and a
retained object reading a resized2-GiB file pass desktop core1/1 (1.82 s) and
fresh GUI-free core1/1 (1.72 s). Current Debian/build-ID/SQLite/four hashes pass.
Last broader48/48 (85.42 s) predates both FSO increments. No new target/ID;
F091 open, fixed456/384/72.

Native Scripting.File binding now uses explicit native identity for qualified
locals/parameters/results. Path is the scalar/default value: CStr(file) reads
the path, VarType(file) is8 and TypeName(file) remains File. DateLastModified
reads actual filesystem modification time through the same shared local-time
conversion as FileDateTime, including UTC/EST changes. Property mutations
remain explicit3251 rather than claiming a writable implementation. Scope:
docs/NATIVE_FILE_SYSTEM_OBJECT.md. Targeted core/DAO2/2 (2.62 s), fresh
headless1/1 (1.66 s), current Debian/build-ID/SQLite/four fixture hashes pass.
Last broader48/48 (85.42 s) predates the FSO increments. No new target/ID;
F091 open, fixed456/384/72.

Native unqualified Kill pathnameExpression now unlinks an actual regular Linux
file only in active VBA branches; a local Kill identifier retains assignment
semantics. No shell/recursive deletion. Missing file53 vs parent76, Null94,
NUL52, read-only/denied70, wrong-kind75; wildcard/final-link/Windows sharing
semantics remain bounded. Scope: docs/NATIVE_DIRECTORY_STATEMENTS.md.
Targeted core1/1 (2.21 s), fresh headless1/1 (1.83 s); final full current
desktop48/48 (84.89 s), current Debian/build-ID/SQLite3.54.1/four original
hashes pass, covering all preceding FSO increments. Original query inventory
remains78/100 without Windows Access result oracle. No new target/ID; F040/F091
remain open, fixed456/384/72.

Native FSO GetFolder(FolderSpec) now returns a real typed Folder with Path/Name
and scalar Path default; trailing separators normalize. Files and SubFolders
produce native typed Files/Folders collections with fresh bounded Count and
For Each snapshots of actual regular files/directories. File.ParentFolder and
non-root Folder.ParentFolder return existing parent Folder objects. No open
handle/host retained; links/special entries, >10000 entries, root parents, Item,
recursive size, mutations and full Windows collection semantics stay bounded.
Scope: docs/NATIVE_FILE_SYSTEM_OBJECT.md. Expanded desktop core1/1 (1.86 s),
fresh GUI-free core1/1 (1.79 s), current Debian/build-ID/SQLite/four hashes pass.
Last broader48/48 (84.89 s) precedes this increment. No new target/ID; F091 open,
fixed456/384/72.

Native FSO CreateFolder(Path) creates one actual directory, returning a typed
Folder. Existing target58 preserves both directory and caller reference; a held
SubFolders collection observes the newly created UTF-8 folder. No recursive
creation, fabricated parent, overwrite or Windows ACL/sharing contract. Named
FSO parameters now follow IFileSystem IDL: GetFile FilePath, GetFolder FolderPath,
CreateFolder/GetFileName/GetAbsolutePathName Path. These supersede earlier
descriptive help placeholders; positional calls are unchanged. Scope:
docs/NATIVE_FILE_SYSTEM_OBJECT.md. Final desktop core1/1 (1.87 s), fresh
GUI-free core1/1 (1.77 s), current Debian/build-ID/SQLite/four fixture hashes
pass. Last broader48/48 (84.89 s) precedes folder/create increments. No new
target/ID; F091 remains open, fixed456/384/72.

Native FSO CopyFile/MoveFile/DeleteFile and File.Copy/Move/Delete/writable Name
now perform actual bounded single-file operations. Shared native_file_operations
keeps VBA FileCopy/Name/Kill and object methods on the same binary-copy, atomic
nonreplacement rename and unlink implementation. Overwrite=False58, live path
updates after File.Move/Name, directory destinations ending /, and read-only
Force policy are real. Trailing separators no longer hide a terminal folder
link from GetFolder or destination checks. Wildcards, links/special files,
cross-filesystem moves, full Windows ACL/sharing/metadata/race/error contracts
stay bounded. Scope: docs/NATIVE_FILE_SYSTEM_OBJECT.md. Final full desktop48/48
(86.15 s), fresh GUI-free1/1 (1.92 s), current Debian/build-ID/SQLite3.54.1 and
all four original hashes after the full run pass. A private native manual probe
confirms Force=False70 preserves a read-only binary test file and Force=True
unlinks it. Original queries78/100; no Windows result oracle. No new test
target/ID; F040/F091 open, fixed456/384/72.

Native read-only FSO OpenTextFile/File.OpenAsTextStream now return an owned
Scripting.TextStream with real Read/ReadLine/ReadAll/AtEndOfStream/Line/Column/
Close. Descriptor validation refuses terminal links/special files; ASCII only,
explicit non-ASCII/Unicode/write/create/count limits3251 rather than invented
conversion. A weak device/inode registry guards shared native copy-destination,
rename-source and unlink operations while a native reader is held; Close releases
it. No host/database retention or full Windows sharing/encoding/EOF oracle.
Scope: docs/NATIVE_TEXT_STREAM.md. Actual native VBA CRLF/LF/position/EOF/NUL/
held-delete/close/reopen workflow passes core/DAO2/2 (2.67 s), fresh headless1/1
(1.68 s), current Debian/build-ID/SQLite/four fixture hashes. Private manual
probe confirms hard-link alias mutation guard, Close release, EOF62 and failed
encoding-read position recovery. Last broader48/48 (86.15 s) precedes this
increment. No new target/ID; F040/F091 remain open, fixed456/384/72.

Native ASCII TextStream writing/appending is now implemented through FSO
CreateTextFile and OpenTextFile/File.OpenAsTextStream modes1/2/8. Real Write/
WriteLine/WriteBlankLines output, CRLF, append, Create and atomic no-overwrite
creation are supported. Payload validation rejects non-ASCII before writing a
prefix. The weak inode registry distinguishes readers/writers; descriptor
validation and sharing registration precede truncation, preserving bytes on
refused overlap. No O_TRUNC occurs before these guards. Unicode/locale codecs,
Folder.CreateTextFile, full Windows sharing/error/mode/position semantics and
filesystem transaction guarantees remain open. Scope: docs/NATIVE_TEXT_STREAM.md.
Final full current desktop48/48 (87.77 s), fresh GUI-free1/1 (2.23 s), both
builds zero compiler warnings; current Debian/build-ID/SQLite3.54.1 and four
original hashes after full tests pass. Current native hard-link/close/EOF/
encoding-recovery probe passes. Original inventory78/100; no Windows result
oracle. No new CTest target/ID; F040/F091 open, fixed456/384/72.

Native TextStream Format=-1/CreateTextFile Unicode=True now implement explicit
UTF-16LE BOM reading, writing and appending. Existing strict UTF converters are
shared in native_utf.hpp with the VBA string runtime, preserving scalar checks.
Actual output bytes include one BOM, correct non-BMP surrogate pairs and CRLF;
append validates rather than duplicates BOM. Read/column counts use UTF-16 units;
a split surrogate read3251 restores position. BOM-less/big-endian/odd-byte/default
code-page and Unicode read/Create=True remain explicit bounds, not guessed
conversion. Scope: docs/NATIVE_TEXT_STREAM.md. Final targeted core/DAO2/2
(2.77 s), full current desktop48/48 (89.99 s), fresh GUI-free1/1 (1.85 s), both
builds zero compiler warnings; current Debian/build-ID/SQLite3.54.1/four original
hashes pass. Private malformed-BOM append preservation, malformed-surrogate
position recovery and refused-create/no-side-effect probe pass. Original78/100;
no Windows result oracle. No new CTest target/ID; F040/F091 open, fixed456/384/72.

Native TextStream Skip/SkipLine now reuse the real checked read/line paths;
AtEndOfLine observes actual next CR/LF without consuming it. Its EOF contract
remains explicit3251. Folder.CreateTextFile now creates an actual plain child
name through the shared atomic-create/guarded-truncate/ASCII-or-UTF16LE pipeline;
path escapes remain3251. Actual VBA line-end/skip/NUL and Folder-created Unicode
readback cases pass core/DAO2/2 (2.69 s), fresh GUI-free1/1 (1.92 s), both builds
zero compiler warnings; current Debian/build-ID/SQLite/four original hashes pass.
Scope: docs/NATIVE_TEXT_STREAM.md. Last broader48/48 (89.99 s) precedes this
increment. No new CTest target/ID; F040/F091 open, fixed456/384/72.

Native Ctrl+P in active top-level forms now follows KeyPreview/KeyDown and can
be cancelled by KeyCode=0 before opening the genuine print dialog. The existing
form toolbar and report print command share a native execution/reentry/active-
window/modal/close guard; pending form controls are confirmed before the dialog.
Closed/cancelled documents are not rendered after acceptance. Report page setup
and the actual layout/report renderers remain in use. Embedded-form contexts,
complete form recordset/page/confirmation and Windows printer/event oracle
remain open. Scope: docs/FORM_PRINT_SHORTCUT_RUNTIME.md. Actual Qt/VBA form
cancellation and genuine form/report dialogs (rejected without physical jobs)
pass targeted layout/mainwindow2/2 (13.00 s); final full current desktop48/48
(86.11 s), current Debian/build-ID/SQLite/four hashes pass. GUI-free core is
unchanged since navigation1/1 (1.92 s). No new CTest target/ID; F148/F135 remain
open, fixed456/384/72.

Native table/query datasheet printing and atomic PDF output are implemented as
bounded work under existing F080/F148/F087; scope and evidence are recorded in
DATASHEET_PRINT_RUNTIME.md. The full requirements remain open, 456/384/72 unchanged.


## QueryDef.Fields without executing query bodies (2026-11-06)

Supported prepared SELECT definitions now expose native DAO.Fields/Field metadata:
Count, zero-based/name Item, enumeration, Name and OrdinalPosition, plus Type/Size
only where the existing native type proof establishes them. Repeated references
preserve identity while held. Preparation does not step rows or invoke native VBA
function bodies; the existing DAO regression checks the actual function-call
counter and unchanged input rows. Action/make-table definitions and ambiguous
field names are refused3251. Query field DefaultValue remains explicitly
unproven3251; unknown types/sizes and other metadata are not guessed.

Collections/fields hold their QueryDef owner but retain no live cursor. SQL
replacement, closure, main-schema changes and workspace rollback invalidate
held metadata3420. Fresh preparation also compares result names/types, rather
than caching UDF declarations forever. Unsupported dependency/context preparation
is still an explicit boundary; full Jet metadata/error precedence and a Windows
Access2003 oracle remain open under F055/F121/F122. No catalog ID or checkbox
changes:456 total,384 scoped confirmations,72 open full requirements.

Nachweis dieser QueryDef.Fields-Etappe: bestehendes DAO-Ziel1/1 (0,96 s), anschließend vollständige Debian13-Desktop-Suite48/48 (98,29 s). Aktuelles Debian-Paket/Build-ID/SQLite3.54.1 und vier Originalfixture-Hashes bestanden. Keine neue Distributionsmatrix/GitHub-CI behauptet.

Metadatenanbindung: vorhandenes DAO-Ziel1/1 (0,99 s) einschließlich tatsächlichem typisiertem VBA, benannten Zugriffen, Refresh und fremdem Thread bestanden; aktuelles Paket/Build-ID/SQLite3.54.1 und vier Originalhashes bestanden. Die Vollsuite48/48 (98,29 s) gehört zum vorherigen a00eb29; kein neuer Vollsuite-/Matrixlauf für dieses kleine Inkrement behauptet.

Parameterfortsetzung: bestehendes DAO-Ziel1/1 (1,09 s), aktuelle App und Debian-Paket/Build-ID/SQLite3.54.1/vier Originalhashes bestanden. Named Index/Refresh und Werterhalt durch echtes VBA geprüft; letzter Vollsuite-Lauf bleibt a00eb29 mit48/48 (98,29 s). Kein neuer Katalogpunkt, keine vollständige Parameter-/Jet-Parität.

TableDef.RecordCount: DAO/originales Kassa-Ziel2/2 (53,71 s) bestanden, einschließlich unabhängiger Zählung aller29 Originaltabellen. Aktuelle App/Paket/Build-ID/SQLite3.54.1 und vier Originalfixture-Hashes bestanden. Bestehende unveränderte Syntaxinventur separat187/187 akzeptiert (1,12 s), keine Ausführungsaussage. Kein neuer Vollsuite-/Matrixlauf dieses Inkrements.456/384/72 unverändert.

Binär-GetChunk: bestehendes DAO-Ziel1/1 (1,23 s), aktuelle App/Paket/Build-ID/SQLite3.54.1 und vier Originalfixture-Hashes bestanden. Nur gespeicherte positive native Binärbereiche; AppendChunk/OLE/Memo/Null/Randsemantik bleiben offen. Letzte Vollsuite a00eb29:48/48 (98,29 s), nicht diesem Inkrement zuschreiben.456/384/72 unverändert.

Native binäre Field.AppendChunk-Fortsetzung implementiert: Byte()-Chunks im Edit/AddNew-Puffer, erster Chunk ersetzt, folgende hängen an; je Feld getrennt, 16MiB, Update/CancelUpdate/Session-Reset und geschützte Mischverwendung. Genauer begrenzter Umfang in docs/DAO_BINARY_FIELDS.md. F016/F055/F121/F122 bleiben offen,456/384/72 unverändert.

Binär-AppendChunk-Nachweis: finales bestehendes DAO-Ziel1/1 (1,38 s), danach vollständiger aktueller Debian13-Desktop48/48 (98,95 s), aktuelles Paket/Build-ID/SQLite3.54.1 und vier Originalhashes bestanden. Byte-Value-Zuweisung bleibt separat geprüft; Mischverwendungsfehler bewahren den Puffer. Kein neuer Test/Katalogpunkt.
Unveränderlicher629b3c9-Folgecheck: Debian12 DAO1/1 (1,49 s), Ubuntu24.04 DAO1/1 (1,27 s), jeweils aktuelle App/DAO neu gebaut, passende Pakete erzeugt/tatsächlich installiert, SQLite3.54.1 und vier Originalhashes bestanden. Diese629-Nachweise enthalten noch nicht die anschließenden TableDef.RecordCount-/Binärchunk-Inkremente; keine neue Vollmatrix/GitHub-CI behaupten. Letzte früheren vollständigen lokalen Distributionssuiten bleiben e2d64c9 mit48/48 je Plattform.456/384/72 unverändert; kein Gesamtabschluss.

Recordset.Fields-Fortsetzung: benannte/geprüfte Indizes, kanonische Ergebnisnamen, gehaltene Collection-/Field-Identität auch bei Enumeration, Null94/fehlend3265/Arity450 und explizite Doppelspalten-Verweigerung3251. Bestehendes DAO-Ziel1/1 (1,01 s), aktuelles Paket/Build-ID/SQLite3.54.1/vier Originalhashes bestanden; Details docs/DAO_RUNTIME.md. Kein neuer Vollsuite-Lauf dieses Inkrements.
Unveränderliche41b193c-Folgeläufe bestanden Debian12 DAO1/1 (1,61 s) und Ubuntu24.04 DAO1/1 (1,21 s), jeweils aktuelle App/Paket/Installation/SQLite3.54.1/vier Originalhashes. Diese Belege enthalten TableDef.RecordCount/GetChunk/AppendChunk, nicht die danach folgenden Field.Name-/Recordset.Fields-Inkremente. Letzte komplette Host-Suite41b193c:48/48 (98,95 s).456/384/72 unverändert; GitHub-CI weiter gesperrt.

Datenblatt-Seiteneinrichtung: funktionierender gemeinsamer Qt-Geometriedialog, Sitzungslayout und explizite versionierte Projektpersistenz, native RunCommand PageSetup32, Papier/Ränder für Vorschau/Druck/PDF. Conditional Savepoint-Speicherung bewahrt konkurrierende Änderungen und Aufrufertransaktionen; ungültige gespeicherte Geometrie blockiert PDF bis zur expliziten Reparatur. Details docs/DATASHEET_PRINT_RUNTIME.md. Ziel1/1 (3,64 s), vollständiger aktueller Desktop48/48 (89,86 s) inklusive Originalrezepte2 Tabellen/39 Zeilen, Paket/Build-ID/SQLite3.54.1/vier Originalhashes bestanden. Keine neue Distributionsmatrix/GitHub-CI;456/384/72 unverändert, F080/F087/F148 bleiben offen.

Druckexemplare: reale Datenblatt-PDFs mit1..100 Kopien, sortierter/unsortierter und vorwärts/rückwärts Seitenfolge; native Qt-Treiberkopien getrennt von manuellen Kopien. Vorhandenes Ziel1/1 (3,99 s), aktuelle App/Paket/Build-ID/SQLite/vier Originalhashes bestanden. Letzte Vollsuitee8375e8:48/48 (89,86 s), dessen GUI-freier Kern1/1 (2,20 s). Unveränderlichee837-Folgeprüfungen Debian12/Ubuntu24.04:2/2 (6,65 s /5,65 s), passende Paketinstallation/Backend/vier Hashes; noch ohne Kopien-Inkrement. Details docs/DATASHEET_PRINT_RUNTIME.md.456/384/72 unverändert, keine physische Drucker-/Windows-Parität behaupten.

Auswahldruck: vollständige markierte sichtbare Datensätze als eingefrorener eigener Text-Snapshot, reale Tabellen-/Readonly-Abfrage-PDFs mit unabhängig geprüfter Ein-/Ausschlussmenge. Leere/teilweise Zellmarkierungen verweigern3251 ohne Ganzdatenblatt-Fallback. Ziel1/1 (4,15 s), aktuelle App/Paket/Build-ID/SQLite3.54.1/vier Originalhashes und drei Uploadhashes bestanden. Details docs/DATASHEET_PRINT_RUNTIME.md. Kein neuer Vollsuite-/Matrixlauf dieses GUI-Inkrements;456/384/72 unverändert.

DAO-DDL-Identität: beobachtete eigene Database/QueryDef.Execute CREATE/DROP und SELECT INTO führen monotone Main-Tabellengenerationen; alte TableDefs/Fields/physische Recordsets/Clones verweigern3420 statt neu erstellte gleichnamige Daten zu ändern. Gleiches SQL/Rootpage/Rowid, alter Schreibpuffer, frische Wrapper, No-op/Fehl-DDL, ALTER und Workspace-Rollback unabhängig geprüft. Externes/unbeobachtetes SQL, TEMP-only und volle Jet-Locks/Lebensdauer bleiben offen. Vollständiger aktueller Desktop48/48 (89,70 s), App/Paket/Build-ID/SQLite/vier Originalhashes bestanden; schließt Kopien/Auswahldruck und Originalrezepte ein. Details docs/DAO_RUNTIME.md.456/384/72 unverändert.

Jet-Rohbinärimport: ausdrücklich gewählter MDBTools-Hexcodec (kein Raten aus CSV), exakte BLOB-/NULL-Werte und Byteanzahl-Vorschau im realen Teil-/Gesamtimport. Laufend begrenzte Child-Ausgabe20MiB/Diagnostik64KiB,30s; großer Export bricht atomar ohne Quelldateiänderung ab. Technische echte Jet4-MDBs mit Jackcess5.0.4 unabhängig geschrieben/zurückgelesen, keine Geschäftsdaten und keine neue App-Java-Abhängigkeit. Decoder-/GUI-/DAO-Bytes/Typ/Chunks/Hash/Rollback geprüft. Ziel2/2 (12,26 s), vollständiger aktueller Desktop48/48 (91,47 s), neue Kassa-Kopie Import/Zellwerte und78/100 Abfragen, App/Paket/Build-ID/SQLite/vier Originalhashes bestanden. Details docs/JET_BINARY_IMPORT.md; keine OLE-Anzeige/Anhänge/Originalformat-Schreib- oder Windows-Parität behaupten.456/384/72 unverändert; kein neuer CTest.

DAO.Relations: lesender tatsächlicher Importkatalog mit Name/Table/ForeignTable/Attributes und geordneten Fields.Name/ForeignName; Integer Count, benannte Indexargumente, echte VBA-Typen/TypeOf/Enumeration, gehaltene Identität und kontrollierte Invalidierung. Erzwungene Beziehungen werden mit allen tatsächlichen Main-Fremdschlüsseln und CASCADE-Aktionen abgeglichen; unbekannte Schlüssel/uneindeutige Metadaten verweigern3251 statt Teilkatalog. Bestehende Ziele2/2 (11,87 s), zusätzlich neue Kassa-Uploadkopie mit acht Originaldefinitionen/acht Feldpaaren bestanden. Details docs/DAO_RELATIONS_RUNTIME.md; Append/Delete/Indexes/Linked und vollständige Jet-/Windows-Parität bleiben offen.456/384/72 unverändert, kein neuer Testkatalogpunkt.

Nicht-Primärindexrichtungen: ursprünglicher Parserverlust von DESC behoben; vollständige geordnete ASC/DESC-Feldlisten für CREATE/ALTER INDEX, ungültige Zusätze verweigern statt Normalisierung. Native Indexanlage erhält Richtung, Reihenfolge und Unique in derselben Importtransaktion. Optionale echte Jet4-Referenz ohne Datensätze mit zwei unabhängig zurückgelesenen Indizes, tatsächlicher Gesamtimport und PRAGMA index_xinfo geprüft; keine App-Java-Abhängigkeit. Details docs/SOURCE_INDEX_DIRECTIONS.md. Kein echter Primärindexname aus generierten MDBTools-Namen behauptet; DAO.Indexes/Primärindexrichtungen/Locale/Originalformat weiterhin offen.456/384/72 unverändert.

Indexrichtungsnachweis: vollständiger aktueller Debian13-Desktop48/48 (89,51 s) bestanden, einschließlich tatsächlicher Jet-Binär-/Indexreferenzen, Originalrezepte für Datenblatt-PDF, Originalkassenhandler und aller originalen Kassa-Indexfeldreihenfolgen/-richtungen. Schließt DAO.Relations und Integer-Count ein. Keine neue Distributionsmatrix/GitHub-CI für diese Inkremente;456/384/72 unverändert.

DAO-Single-Werte: bestätigte physische Felder mit DAO-Typ6 liefern nun bei Value/Bang/Clone und im Edit/AddNew-Puffer echte Single-Werte; Schreiben verengt vor Puffermutation und persistiert binary32-Wert, Null/Empty und explizite Konvertierungsfehler bleiben geprüft. Echte technische Jet4-Float-Referenz mit unabhängig zurückgelesenen Werten, tatsächlicher Import/Metadaten/Update/Requery/Quellhash und VBA VarType/CSng im bestehenden Ziel2/2 (1,94 s) bestanden. Keine automatische Query-/Formularlineage und keine verlustfreie allgemeine Single-Importbehauptung: MDBTools1.0.0 %.8g kann einzelne Bitmuster verändern, belegtes Gegenbeispiel in docs/DAO_SINGLE_FIELDS.md. F014/F041/F055/F105/F122 bleiben offen,456/384/72 unverändert. Letzte Vollsuite4548ce7:48/48 (89,51 s), nicht diesem Typinkrement zuschreiben.

Originale Indexmetadaten: separater lesender nativer MDBTools-Helfer statt generierter CLI-Primärnamen. Strikt versionierte Namen/Kind/Flags/geordnete Felder/Richtungen, gegen tatsächliche Tabellenspalten und ursprünglichen Primärschlüssel validiert; atomare main._la_access_indexes-Persistenz. Neue Importe verwenden echte Nichtprimärnamen. Zusammengesetzte und einzelne DESC-Primärschlüssel bleiben echte native DESC-Indizes; DESC-AutoNumber und unbewiesene IgnoreNulls/Required-Konfigurationen verweigern explizit. Bestehende GUI/Import/Kassa-Ziele3/3 (15,27 s), abschließende erweiterte technische Jet-Referenz1/1 (0,96 s);20 echte Primär-/48 andere Originalnamen in29 Tabellen nachgewiesen. Details docs/SOURCE_INDEX_METADATA.md. Noch kein vollständiger DAO.Indexes-/Windows-Nachweis;456/384/72 unverändert. Build benötigt mdbtools-dev/pkg-config, Paket enthält geprüften lesenden Helfer.

Indexmetadaten-Grundlage vollständig verifiziert: aktueller Debian13-Desktop48/48 (85,91 s), neue echte Jet-Referenzen einschließlich einzelner und zusammengesetzter DESC-Primärindizes, Originalrezepte, Originalkassenhandler und Originalindexnamen bestanden. Aktuelles Paket/Build-ID für App und Metadatenhelfer, tatsächliches Laden des entpackten Helfers, SQLite3.54.1 sowie vier Originalfixture-/drei Uploadhashes bestanden. Kein neuer Katalogpunkt und keine vollständige DAO.Indexes-/Windows-Parität;456/384/72 unverändert. Lokale Debian12-/Ubuntu24.04-Folgeprüfung der neuen Build-Abhängigkeit folgt separat.

DAO.Indexes: lesender originaler Indexkatalog mit nativen Typen/TypeOf, Integer Count/Item/named Index/Refresh/Enumeration, echten Namen und Boolean Primary/Unique/Required/IgnoreNulls sowie geordneten Fields.Name/Long Attributes (dbDescending). Jeder Zugriff vergleicht den vollständigen tatsächlichen Main-Indexkatalog, Primärschlüssel/Rowid-Alias, Unique, Felder/Richtungen und BINARY-Collation. Fehlende Quellen, logische kind2/IgnoreNulls/unbekannte oder zusätzliche native Indizes verweigern3251 ohne Teilkatalog. Gehaltene Definitionen/Fields invalidieren kontrolliert nach Schema/Metadaten/Rollback, TEMP bleibt wirkungslos. Bestehende Ziele3/3 (11,69 s), zusätzlich neue Kassa-Kopie mit20 Primär-/48 anderen DAO-Indexidentitäten, aktuelle App/Paket/Build-IDs/Helfer/SQLite3.54.1/vier Originalfixture-/drei Uploadhashes bestanden. Details docs/DAO_INDEXES_RUNTIME.md. Keine Mutation/Foreign/Clustered/DistinctCount/Windows-Parität;456/384/72 unverändert. Letzte Vollsuite bleibt74d75c7:48/48 (85,91 s).

Unveränderliche74d75c7-Folgeprüfung abgeschlossen: Debian12/GCC12/Qt6.4 und Ubuntu24.04/GCC13/Qt6.4 jeweils5/5 gezielte Ziele (30,64 s /27,39 s), neues mdbtools-dev/pkg-config, Paketinstallation, Backend3.54.1, Originalhashes und tatsächlich installierter Helfer mit originalen Jet-Namen/Richtungen unter MDBTools1.0.0 bestanden. Diese Snapshots enthalten die Metadaten-Grundlage, noch nicht die folgende DAO.Indexes-Anbindung; keine GitHub-CI- oder Vollsuitebehauptung für die Matrix.

Exakter Single-Import: realer MDB_FLOAT-Rohwertzugriff im vorhandenen lesenden Helfer (--single-values), IEEE754-Bitmuster statt verlustbehafteter %.8g-Speicherung. Source-SHA256 vor/nach Scan und atomarem Import, echte SINGLE-Spalten/REAL-Affinität, CSV-Zeilen/NULL/8g-Abgleich,1M-Zellen/16MiB-Grenze. Rohwerte bleiben außerhalb des angezeigten SQL-Schemas. Gegenfallbits509022793, minimale Subnormale, maximaler endlicher Single, Brüche/+0/NULL nach unabhängiger Jackcess-Jet-Erzeugung im tatsächlichen Gesamt-/Teilimport und DAO bitgenau nachgewiesen. NaN/Infinity/negatives Null oder inkonsistente CSV-Daten verweigern atomar. Bestehende relevante Ziele5/5 (27,76 s), finaler Affinitätsguard im Importziel1/1 (1,05 s); aktuelle App/Paket/Build-IDs/SQLite3.54.1/vier Originalfixture-/drei Uploadhashes und tatsächlich entpackter Helfer mit bekannten Originalbits/Quellhash bestanden. Details docs/SOURCE_SINGLE_PRECISION.md. Legacy-CSV, Double/Decimal/Currency, generische SQL-/Default-/Query-Semantik und Windows bleiben offen;456/384/72 unverändert. Letzte Vollsuite74d75c7:48/48 (85,91 s), neue Matrix noch nicht behaupten.

Double-Import bitgenau: vorhandener lesender Helfer --double-values, binäre64-Bit-Quelle statt %.16g-Rückwandlung. Gemeinsamer streng validierter access_floating_metadata-Codec für Single/Double, Classic-Locale-Formatprüfung ohne Double-Overflow bei der Dezimaldarstellung von MAX_VALUE. Original-Rohwerte aus der bereits geschützten Gesamtimport-Snapshotdatei; strukturleere Anlage scannt Daten nicht doppelt. Low-level API unterscheidet ausdrücklich caller-owned Csv (kein Dateizugriff aufgrund eines Loglabels) von NativeMdbBits; reale Gesamt-/Inspector-Kopien und Originalprüfungen wählen NativeMdbBits, fehlender Nachweis fällt niemals still auf CSV zurück. Dadurch bleiben technische Tests mit echten Schemas und eigens bereitgestellten CSV-Zeilen korrekt getrennt.

Unabhängiger Jackcess-Auszug beider Kassa-Testkopien:13 Tabellen/57 Double-Spalten,69405 Nicht-NULLwerte/514 NULLs, kein negatives Null/NaN/Infinity. Der bisherige Dezimalexport reproduzierte2404 numerische Zeilentupel in zwei Tabellen nicht exakt; neue native Kopien stimmen bitgenau mit allen unabhängigen Tupeln überein. Echte Jet-Extremwertreferenz, Gesamt-/Teilimport, DAO Type7/Size8, Herkunfts-/CSV-Guards und atomare Negativnull-Verweigerung geprüft. Aktueller korrigierter Debian13-Desktop48/48 (143,10 s), aktuelle neue Uploadkopie erneut29/29 und unabhängiger Auszug exakt; dort weiterhin78/100 Originalabfragen ohne Windows-Ergebnisorakel. App/Paket/Build-IDs/SQLite3.54.1/vier Originalfixture-/drei Uploadhashes und tatsächlich entpackter --double-values-Helfer mit bekannten Originalbits/SHA bestanden. Details docs/SOURCE_DOUBLE_PRECISION.md. Zuvor entdeckte technische CSV-Herkunftsfehler behoben, keine Schutzprüfung für Originale entfernt.456/384/72 unverändert; keine neue Matrix für Double behaupten.

Nach einmaligem lesendem Resume derselben Codex-Call-Anfrage3ea2a1b8-75c6-411c-8f85-3cb76c4f3571 wurde status=unresolved und eine normale finale LiveKit-Teilantwort „Genau“ zur Windows-Referenzfrage geliefert; keine konkrete Umgebung/Protokolle/Zugriffsinformation und keine confirmed_decisions. CI-Frage bleibt unbeantwortet. Beide externen Blocker bleiben offen, keine Freigabe erfinden oder neue Anrufanforderung daraus ableiten. Privates Ergebnis resumed-result.json im bestehenden Call-Verzeichnis.


AutoNumber-Kern, 2026-11-08: Quellstände und verwaltete signed32-Vergabe für neue
Importe/neue Designertabellen, DAO-AddNew-Reservierung vor Rückgabe, verbrauchte
Cancel-/Updatefehler-/Workspace-Rollback-Nummern, negatives Wrap, tatsächliche Long-
Typen, NULL-/Änderungsschutz und atomare Lifecycle-Grenzen implementiert. Gemeinsame
bestehende Integration6/6 in101,55s; nachgezogene betroffene Layout-/Identitätsfälle
1/1 in11,02s und1/1 in0,23s. Kein Vollsuite-/GitHub-CI-Nachweis. F171 bleibt für die
in AUTONUMBER_IMPORT.md ausdrücklich beschriebenen breiten Restanforderungen offen;
456/384/72 unverändert.

Nativer CreateTableDef/CreateField/Fields.Append/TableDefs.Append-Schemaaufbau
und TableDefs.Delete implementiert; gemeinsame DDL/Counter/Typ-Metadaten-Atomizität,
echter VBA-Durchlauf und gezielte DAO/Import2/2 (1,25s), App gebaut. Grenzen in
DAO_TABLEDEF_FIELDS.md; F055/F121 offen,456/384/72 unverändert.

Detached DAO-Indexbuilder für eigene neue Tabellen samt realen Primär-/Unique-/
ASC-/DESC-Constraints und Kollisions-Rollback/Retry implementiert. DAO/Import2/2
(1,22s), App gebaut. DAO_INDEXES_RUNTIME.md; F055/F121/F130 offen,456/384/72 gleich.

DAO CreateIndex/Indexes.Append/Delete für Nichtprimärindizes vorhandener geprüfter
Tabellen implementiert. Unique-Fehleratomizität, ASC/DESC, Rollback/Identität im
bestehenden DAO/Import2/2 (1,31s) geprüft, App gebaut. DAO_INDEXES_RUNTIME.md;
F121/F130 offen,456/384/72 unverändert.

Grafische Tabellenanlage auf gemeinsamen DAO-Schemaweg umgestellt; native Typen,
Text255/Memo, Primärindex und anschließende Indexanlage auf tatsächlicher UI-
Tabelle im bestehenden Mainwindow/DAO2/2 (4,95s) geprüft. App gebaut; Grenzen in
DAO_TABLEDEF_FIELDS.md.456/384/72 unverändert.

Grafische Indexverwaltung auf gemeinsamem DAO-Weg umgesetzt. Tatsächliche modale
Anlage/Löschung mit Richtungen und unveränderten Datensätzen im bestehenden
Mainwindow/DAO2/2 (5,11s) geprüft, App gebaut. DAO_INDEXES_RUNTIME.md;
456/384/72 unverändert.

Native primitive Speichergrenzen und deklarierte Required-/Attributes-Eigenschaften
umgesetzt. GUI/DAO2/2 (5,00s), abschließende kombinierte primitive DAO-Prüfung1/1
(1,21s), App gebaut. Vier Fixture-/drei Uploadhashes unverändert; Herkunfts- und
Konvertierungsgrenzen DAO_TABLEDEF_FIELDS.md.456/384/72 unverändert.

Grafische Textgröße1..255 mit vorgeschalteter Größen-/Defaultprüfung und
Entwurfserhalt umgesetzt; tatsächlicher GUI-Save/DAO-Größenrücklesen im
bestehenden Mainwindow1/1 (3,91s), App gebaut.456/384/72 unverändert.

Normale DAO-Feldanlage/-löschung an vorhandenen Tabellen umgesetzt; tatsächliches
VBA/named args/Bestandswerte/Metadaten/Fehleratomizität/Abhängigkeiten/Rollback im
bestehenden Mainwindow/DAO2/2 (5,38s) geprüft, App gebaut. DAO_TABLEDEF_FIELDS.md;
456/384/72 unverändert.

Grafische Feldbearbeitung samt Bestandswerten/Defaults/Bestätigung/Abhängigkeiten
und O(1)-DAO-Cursorinvalidierung bis Requery umgesetzt; reale GUI/Clone-Puffer-
Prüfung Mainwindow/DAO2/2 (5,11s), App gebaut. Details DAO_TABLEDEF_FIELDS.md.
456/384/72 unverändert.

Kontrolliertes Refresh gehaltener TableDef.Fields/Indexes nach eigener DDL/
Rollback implementiert; Iteminvalidierung und gleichnamiger Tabellenersatz im
bestehenden DAO1/1 (1,23s) geprüft, App gebaut.456/384/72 unverändert.

Detached DAO-Relationbuilder für explizite DontEnforce2-Beziehungen samt
Persistenz/Abhängigkeits-/Generations-/Rollbackschutz umgesetzt. Bestehendes
DAO/Import2/2 (1,33s), App gebaut; tatsächliche RI-Migration weiterhin offen.
DAO_RELATIONS_RUNTIME.md,456/384/72 unverändert.

Echte native numerische Fremdschlüssel/Cascades und Entfernung mit verifiziertem
Schemaumbau/Datenerhalt implementiert. Bestehendes finales DAO1/1 (1,33s), App
gebaut; Counter-/View-/Trigger-/Index-/Kindtabellen-/Regelerhalt geprüft. Grenzen
DAO_RELATIONS_RUNTIME.md.456/384/72 unverändert.

Echte grafische Beziehungsverwaltung mit RI/Cascade-/DontEnforce-Modus und
bestätigter Entfernung umgesetzt. Reale modale FK-/Bestandswerte-Prüfung im
bestehenden Mainwindow/DAO2/2 (5,33s), App gebaut. DAO_RELATIONS_RUNTIME.md;
456/384/72 unverändert.

Explizite native1:1-Relationen mit echten Unique-Schlüsseln und Löschabhängigkeits-
schutz sowie GUI implementiert. Bestehendes Mainwindow/DAO2/2 (5,64s), App gebaut.
DAO_RELATIONS_RUNTIME.md;456/384/72 unverändert.

Vorhandene native Feldproperties Required/DefaultValue/Text.Size jetzt über DAO
und grafischen atomaren Batch-Save änderbar. Geprüfter Schema-/Bestandswert-/
Countererhalt, Fehlerrollback und Cursorinvalidierung bis Requery; Original-/
Counter-/unbekannte Schemaänderungen bleiben begrenzt. Unabhängige unbenannte
Fremdschlüssel blockieren den Feldumbau nicht; betroffene unbekannte Schlüssel
bleiben geschützt. Bestehendes GUI/DAO2/2 (5,92s), App gebaut. Details/Grenzen
DAO_TABLEDEF_FIELDS.md.456/384/72 unverändert. Weitere offene Anforderungen
fortlaufend implementieren; kein Zwischenabschluss oder Abschlussanruf.

Native Nicht-Counter-Felder sind jetzt per DAO Field.Name und grafisch atomar
umbenennbar, einschließlich Index-/Beziehungsmetadaten, echten FK-Verweisen und
Digestprüfung der Werte/rowids. Access-Abfrageartefakte/VBA/Formularbindungen
werden nicht refaktoriert. Echtes GUI/DAO2/2 (5,66s), App gebaut. Umfang/Grenzen
DAO_TABLEDEF_FIELDS.md.456/384/72 unverändert. Weiter native Feldvalidierung
und weitere offene Anforderungen implementieren, keine Zwischen-Finals.

Native AllowZeroLength für Text/Memo ist bei DAO-Anlage und vorhandenen Feldern
sowie in Neuer-Tabelle-/Feld-Dialogen implementiert, echte SQL-Prüfung/3315,
Null separat Required. Physische DAO-/Clonefelder lesen explizite Required-/
AllowZeroLength-Metadaten schreibgeschützt. Atomarer Umbau/Fehlerrücknahme;
Original-/unbewiesene Eigenschaften nicht erfunden. GUI/DAO2/2 (5,99s), App
gebaut; Details DAO_TABLEDEF_FIELDS.md.456/384/72 unverändert. Nächste Arbeit:
Primärindizes vorhandener nativer Tabellen wirklich anlegen/entfernen, samt
Daten-/Counter-/Abhängigkeitsschutz und grafischer Bedienung; dann weiter.

Primärindizes vorhandener nativer Tabellen sind jetzt per DAO und grafisch
wirklich anleg-/entfernbar: echter PK/NOT NULL, atomarer Daten-/Counter-/
Regelerhalt, Abhängigkeits-/Cursor-/Transaktionsschutz. GUI1/1 (4,30s), finales
DAO1/1 (1,64s), App gebaut; keinen frischen2/2-/Gesamtlauf behaupten. Details
DAO_INDEXES_RUNTIME.md.456/384/72 unverändert. Als Nächstes Required-Indizes
vorhandener nativer Tabellen mit echten Nullregeln auch auf bisher nullable
Feldern sowie korrekter Regelentfernung umsetzen; danach fortlaufend weiter.

Required-Indizes vorhandener nativer Tabellen erzwingen jetzt echte NOT-NULL-
Regeln auch auf bisher nullable Feldern. Vorhandene Nulls verweigern atomar3314;
Löschen entfernt nur implizite Indexregeln, andere Pflichtregeln bleiben erhalten.
Entfernte physische Indizes werden nicht versehentlich wiederhergestellt. GUI/
DAO2/2 (5,73s), App gebaut; DAO_INDEXES_RUNTIME.md.456/384/72 unverändert.
Nächste Arbeit: sichere physische DAO-Zeilenidentität bei normalen Nutzerfeldern
namens rowid/_rowid_/oid; danach andere offene Anforderungen ohne Zwischenfinal.

Physische DAO-Recordsets verwenden jetzt einen nachgewiesenen unverschatteten
rowid/_rowid_/oid-Alias für Lesen, Update/RETURNING und Löschen; normale gleich
benannte Nutzerfelder können keine falschen Zeilen mehr adressieren. Familie/
Clone/Requery übernehmen den Alias. Alle drei Aliasnamen oder WITHOUT ROWID
verweigern3251. SQLite3.54.1-Alias-ALTER-Grenze wird ebenfalls ausdrücklich3251,
keine vollständige Schemaunterstützung behaupten. Final DAO1/1 (1,58s), App
gebaut, DAO_TABLEDEF_FIELDS.md.456/384/72 unverändert. Nächste größere native
Katalogaufgabe: TableDef.Name für vorhandene eigene Tabellen samt tatsächlicher
Counterdefault-/Index-/Beziehungsmetadaten-Umbindung; dann weiter ohne Zwischenfinal.

TableDef.Name und grafische Tabellenumbenennung für eigene native Tabellen
implementiert: tatsächliche Counterdefault-/Frontier-/Schritt-/Index-/
Beziehungs-/Seiteneinstellungs-Umbindung, physische FK/View/Trigger-Verweise,
Digest-/Generations-/Rollbackschutz. Gespeicherte Access-Abfragen/VBA/Formulare
werden nicht refaktoriert; weitere Grenzen DAO_TABLEDEF_FIELDS.md. GUI/DAO2/2
(5,82s), App gebaut. Originaluploads3/3 und Fixtures4/4 Hashes unverändert.
456/384/72 unverändert. Nächste Aufgabe: echte Verbindungs-Schreibfähigkeit und
readonly DAO-Fehler/Metadaten sowie nachgewiesene lokale TableDef-Metadaten;
dann weitere offene Anforderungen fortlaufend, kein Zwischenabschluss.

DAO-Schreibfähigkeit prüft jetzt echtes SQLite-readonly und query_only statt
lediglich Recordsettyp. Database/TableDef/Recordset.Updatable, readonly Counter-
Metadaten und3027-Verweigerung vor Daten-/Countermutation implementiert; Cancel
bleibt möglich. Explizit eigene lokale TableDefs haben Attributes0/Connect=""/
SourceTableName=""; keine erfundenen importierten Linkdaten oder externen Provider.
Final DAO1/1 (1,66s), App gebaut, technischer Readonly-Dateivergleich unverändert.
DAO_TABLEDEF_FIELDS.md;456/384/72 gleich. Weitere offene DAO-/Schemaanforderungen
fortlaufend implementieren; keine Zwischen-Finals/Abschlussanrufe.

Native Single-Felder6/Size4 jetzt im DAO-/Tabellen-/Feldbau verfügbar, mit echter
quantisierter REAL-Speicherung/Prüfung. Gemeinsame Single-Konvertierung für DAO,
Datenblatt und begrenzte direkte SQL-Schreibpfade/physische Append-Selects,
Literaldefaults/ältere Defaultinserts und direkte Query-Typweitergabe. Negative
Null/Unterlauf bleiben3251 statt stiller Normalisierung. Final GUI/DAO2/2
(6,12s), App gebaut; vorher Import1/1 (0,06s), kein frischer3/3-Lauf. Grenzen
DAO_TABLEDEF_FIELDS.md/SOURCE_SINGLE_PRECISION.md; insb. komplexe/joined Writes,
Arithmetik/Locale/Access-Oracle und erzwungene Single-FKs noch offen.456/384/72
unverändert. Andere offene Anforderungen fortlaufend weiterprogrammieren.

Echte Single-Fremdschlüssel/Cascades jetzt für nachgewiesene eigene SINGLE_REAL-
Felder beider Seiten, einschließlich GUI-Auswahl. Nicht nachgewiesene Legacy-/
importierte REAL-Single-Schlüssel bleiben verweigert. GUI/DAO2/2 (6,26s), App
gebaut; DAO_RELATIONS_RUNTIME.md. Nur der Titelguard des GUI-Fehlerhelfers wurde
danach eingeengt, keine Produktionsänderung nach Prüfung.456/384/72 unverändert.
Fortsetzung: native Date-/Now-/Time-Defaults und Datumsnormalisierung prüfen/
implementieren, danach weitere offene Anforderungen ohne Zwischen-Final.

Native Datumsdefaults Date()/Now()/Time() jetzt für eigene Date-Felder im DAO-,
Tabellen-/Feldentwurf, manuellen Datenblatt und direkten Tabellenformular.
Neue Puffer speichern den initialisierten Uhrwert; explizit Null überschreibt,
ungebundene Datumsdefaults werden mitgespeichert. Neue Date-Felder erzwingen
kanonische ISO-Speicherung; alte Definitionen bleiben bei Migration erhalten.
Nur begrenzte eigene Uhrfunktionen, keine beliebigen SQL-/VBA-Defaults. App
gebaut, GUI/DAO2/2 (6,24s). SOURCE_LITERAL_DEFAULTS.md/DATE_RUNTIME.md;
F014/F015/DAO-/Formular-Vollanforderungen bleiben offen,456/384/72 unverändert.
Fortsetzung: eigene kanonische Date-Beziehungsschlüssel samt GUI und tatsächlichen
FK-/Kaskadenwirkungen; danach weitere offene Anforderungen, kein Zwischen-Final.

Eigene kanonische Date-Schlüssel jetzt für tatsächliche referenzielle Integrität,
Änderungs-/Löschkaskaden und Beziehungsdialog freigegeben. Beide Seiten benötigen
nachgewiesene eigene Date-Metadaten und DATETIME-Speicherung; ältere/importierte
Date-Schlüssel bleiben3251. Bestehende Uhrdefaults und binäre Nutzwerte bleiben
bei der Migration erhalten. GUI/DAO2/2 (6,11s), App gebaut. Keine vollständige
Jet-Datums-/Locale-/Untersekundenparität;456/384/72 unverändert. Fortsetzung:
native numerische Gültigkeitsregeln im DAO/Feldentwurf mit tatsächlicher
Speicherprüfung und atomarer Bearbeitung, danach weitere offene Anforderungen.

Numerische native Field.ValidationRule im DAO/Feldentwurf mit echten CHECKs
implementiert: explizite Vergleiche mit And/Or oder Between, Null separat von
Required,3317 bei Verletzung. Atomare Regeländerung prüft vorhandene Werte;
Rollback, fehlerhafte Mehrzeilenänderung, erhaltene DAO-Puffer, Löschen der Regel,
Feld-/Tabellenrename und Byte/Integer/Long/Single/Double/Boolean geprüft.
GUI/DAO2/2 (7,08s), App gebaut; DAO_TABLEDEF_FIELDS.md. Allgemeine Expressions,
ValidationText und volle Referenzsemantik bleiben offen;456/384/72 unverändert.
Fortsetzung: eigene Validierungstexte und physische Recordset-Feldmetadaten,
danach übrige offene Anforderungen ohne Zwischenabschluss.

Native ValidationText jetzt im DAO/Feldentwurf für unterstützte numerische Regeln,
mit wirklichem3317-SQL-Fehlertext; UTF-8/Umlaute, Großschreibung und Quotes bleiben
erhalten. Leeren/Bearbeiten/Migration unterstützt; Begrenzung255 UTF-16-Einheiten,
kein NUL. Physische Recordset.Fields lesen eigene ValidationRule/ValidationText
schreibgeschützt, mit bestehenden Clone/Requery-/Generationsguards. GUI/DAO2/2
(7,32s), App gebaut. Allgemeine Ausdrücke und genaue Windows-Formular-/Fehler-
Ereignissemantik weiterhin offen.456/384/72 unverändert. Fortsetzung: tatsächliche
Feldreihenfolge/OrdinalPosition im DAO und Feldentwurf mit atomarem Werterhalt.

Native TableDef.Field.OrdinalPosition verschiebt jetzt echte physische Spalten,
auch Counter, atomar mit Feld-/Indexkatalog und GUI-Positionssteuerung. Vorhandene
Werte/Rowids, Counterfrontier, Single-/Uhrdefaults, Regeln/Texte, Indizes, FK,
Trigger und Views bleiben erhalten; physische DAO-Puffer/Clones3420 bis Requery.
Unveränderte Position ohne Schemaänderung, außerhalb vorhandener Positionen3251;
keine ursprünglichen/unklaren Schemas oder aktiven Aufrufertransaktionen umbauen.
GUI/DAO2/2 (7,15s), App gebaut. DAO_TABLEDEF_FIELDS.md;456/384/72 unverändert.
Fortsetzung: Gültigkeitsregeln für eigene kanonische Date-Felder einschließlich
Uhrgrenzen im DAO/Feldentwurf, dann übrige offene Anforderungen fortlaufend.

Gültigkeitsregeln jetzt auch für eigene kanonische Date-Felder: ISO-#Datum#-
Konstanten (ganze Sekunden) und Date()/Now()/Time() mit Vergleichen, And/Or,
Between. Tatsächliche CHECK-Durchsetzung/3317 und ValidationText, atomare
Bestandsprüfung, Rename/Migration/GUI; alte/importierte Date-Speicherung bleibt
verweigert. Regionale Konstanten, Untersekunden/allgemeine Ausdrücke offen.
GUI/DAO2/2 (7,08s), App gebaut;456/384/72 unverändert. Fortsetzung: verlustfreie
numerische Datentypänderungen im Feldentwurf mit Bestands-/Defaultprüfung;
DAO Field.Type nach Append bleibt readonly. Danach offene Anforderungen weiter.

Feldentwurf unterstützt jetzt verlustfreie numerische Typänderung für eigene
Nicht-Counter: Boolean/Byte/Integer/Long/Single/Double mit Bestands-/Default-
Konvertierungsprüfung vor Umbau. Verlust/Rundung verweigert3251, Überlauf6;
Null, Indizes, Regeln/Texte und Counterfrontier geschützt. Numerische Werte
werden bei beabsichtigtem Speicherklassenwechsel bitgenau verglichen; andere
Werte/Rowids bleiben typgenau. Relationkeys vorher lösen; DAO Field.Type nach
Append weiter383. GUI/DAO2/2 (6,94s), App gebaut; DAO_TABLEDEF_FIELDS.md.
456/384/72 unverändert. Fortsetzung: doppelte Hauptfenster-Kompilierung von App
und GUI-Ziel über gemeinsame statische Desktopimplementierung beseitigen, dann
weitere offene Entwurfs-/Access-Anforderungen ohne Zwischenabschluss.

App und Hauptfensterintegration teilen jetzt dieselbe statische Desktopbibliothek
(MainWindow/Inspector/BridgeImporter und Qt-Metaobjekte), ohne zusätzliche Runtime-
DSO. Hauptfenster wird im gemeinsamen Build nur einmal kompiliert. GUI1/1 (5,25s),
GUI-freier Kern1/1 (1,93s), App mit SQLite3.54.1/JOIN-/Null-/Counter-Backendcheck
bestanden; kein frischer Vollsuite-/Distributionslauf.456/384/72 unverändert.
Fortsetzung: ISO-#Datum#-Defaults und führendes= bei begrenzten Literaldefaults
im Tabellen-/Feldentwurf und DAO; danach weitere offene Anforderungen fortlaufend.

Native Datum-Standardwerte akzeptieren jetzt eindeutige ISO-#Datum#-Literale
(optional Uhrzeit mit ganzen Sekunden) und führendes=. Begrenzte allgemeine
Literaldefaults akzeptieren ebenfalls= und äußeren Leerraum; Textinhalt bleibt
unverändert. DAO-Anlage/Bearbeitung, Tabellen-/Feldentwurf und Datumsregeln nutzen
gemeinsamen Date-Parser; ungültige Tage/regionale Formen verweigern3251. GUI/DAO2/2
(7,98s), App gebaut; erster Compilefehler war fehlender Prüfcode-Include und ist
behoben. Kein Importparser-/Windowslexik-/Vollsuiteclaim;456/384/72 unverändert.
Fortsetzung: Table-Recordsets/Index/Seek für belegte native numerische Indizes,
mit tatsächlicher Indexreihenfolge und klaren unbewiesenen Grenzen.

Native physische Table-Recordsets1 jetzt explizit/dbOpenTable sowie als Default
für Main-Tabellen/TableDefs; SQL-/QueryDef-Dynasets2/Snapshots4 bleiben erhalten.
Geprüfte native nichtnullable numerische Indizes liefern tatsächliche Reihenfolge,
Index/Seek mit kompletten Schlüsseln, ASC-Bereichen, DESC-Gleichheit, NoMatch/
Bookmarks, benannten Argumenten und kontrollierten Clonegrenzen. NativeSnapshot-
Suche, kein B-tree-/Live-Jet-Claim; Date/Text/Null/Partial/DescendingRange und
unabhängiger Clone-Indexwechsel noch offen. DAO_RECORDSET_INDEX_SEEK.md mit
aktuellen Microsoft-API-Referenzen, kein2003-Oracle. Core/GUI/DAO/originalesKassa-
Ereignis4/4 (122,64s; davon Originalimport113,62s), App gebaut. Originaluploads3/3,
Fixtures4/4 Hashes unverändert. Keine erneute breite Originalprüfung ohne neuen
betroffenen Umfang.456/384/72 gleich. Fortsetzung: eigene kanonische Date-Index-
Schlüssel und typed Seek, danach offene Anforderungen ohne Zwischenabschluss.

Index/Seek jetzt auch für eigene kanonische nichtnullable Date-Felder und gemischte
Date-/Zahlenschlüssel. Gebundene Date-Variants und VBA-#Datum# bleiben typisiert;
String-Datekeys verweigern13. Datebereiche vergleichen kanonische ganze Sekunden
entsprechend tatsächlicher Indexreihenfolge. DAO1/1 (2,20s), App gebaut; keine
Wiederholung des unverändert betroffenen Originalimports. Voller Jet-/OLE-/Locale-
Oracle bleibt offen,456/384/72 gleich. Fortsetzung: dokumentierte zusammengesetzte
Teilschlüssel-Suche mit fehlenden hinteren Schlüsseln als Null für die bereits
strukturell nichtnullen unterstützten Indizes; danach weitere offene Anforderungen.

Seek unterstützt jetzt ausgelassene hintere Schlüssel zusammengesetzter
nichtnuller eigener Indizes: Suffix wird gemäß dokumentierter Access-Suche als
Null behandelt; Gleichheit NoMatch, aufsteigende Präfixbereiche passend. Numerik/
Date, explizit Missing am Suffix, Bookmark-Rettung geprüft. Explizite Nullwerte,
Lücken und absteigende Bereiche bleiben3251. DAO1/1 (2,25s), App gebaut;
456/384/72 gleich. Fortsetzung: eigene Feld-Caption/Description im Entwurf mit
wirklichen Tabellenköpfen und Formular-/Berichtslabels, dann übrige Anforderungen.

Eigene Feld-Caption/Description jetzt tatsächlich im Projekt, Entwurf (auchCounter),
Datenblattkopf/Hinweisen und kopierten Formular-/Berichtslabels. Physische Namen/
Werte/Counter bleiben gleich. DAO.CreateProperty Text10 sowie begrenztes eigenes
Field.Properties Item/Append/Delete/Refresh, Value/defaultValue, echte native
Property/Properties-Typbindung und readonly Recordset-Properties implementiert.
Persistente Property-IDs verhindern Wiederbelebung nach Delete/Recreate; keine
vollständige Count-/Ordinal-/Enum-Liste erfunden. Schema-/Rollbackguards bleiben.
Details NATIVE_FIELD_DISPLAY_PROPERTIES.md. GUI1/1 (5,65s); danach fehlende
Typregistrierung/TypeOf-Registrierung korrigiert und zu gemeinsamer Registry
vereinheitlicht: Core/DAO2/2 (4,66s), App gebaut. Kein frischer3/3-/Vollsuiteclaim.
456/384/72 unverändert. Fortsetzung: weitere tatsächlich vorhandene native Feld-
Property-Metadaten über Properties statt nur direkte Member anbinden; vollständige
originale Properties-/Windowsparität bleibt offen. Kontinuierlich weiterarbeiten.

Native Field.Properties stellt jetzt zusätzlich vorhandene intrinsische Metadaten
über echte Property-Objekte bereit: Name/Type/Size/Attributes/DataUpdatable/
OrdinalPosition/Required/AllowZeroLength/DefaultValue/ValidationRule/-Text, jeweils
nur soweit direkter Getter belegt. Kontext bleibt erhalten (Counter TableDef17,
Recordset49); Value-Schreiben nutzt vorhandene atomare Setter, readonly bleibt383.
Eigene Property-Familie folgt zulässigen Größen-/Pflicht-/Namensänderungen; kein
Count/Enum/Vollkatalog erfunden. DAO1/1 (2,46s), App gebaut;456/384/72 gleich.
Fortsetzung: tatsächliche Caption/Description-Spaltenproperties aus MDBTools
lesend übernehmen; vorhandene mdb_col_get_prop-API geprüft, keine VBA-Extraktion.

Echte Jet4-Caption/Description-Spaltenproperties werden jetzt aus vorhandener
lesender mdb_col_get_prop-API im selben Metadatenlauf übernommen (keine extra
Prozesse proTabelle). Versionierter begrenzter Abschnitt, vollständige geordnete
Spaltenbindung, Format-/SHA-/Textprüfung; fehlende Komponente bleibt unbekannt.
Import speichert belegte Werte/Abwesenheit, leere Strings und eigene Lease-IDs;
Datenblatt/Form-/Berichtserzeugung und readonly DAO-Properties lesen sie.
Andere Formate/alte Projekte/Vererbungsflag/Vollkatalog nicht geraten. Echte
unabhängige Jackcess-Jet4-Referenz ohne Datensätze bestätigt Unicode/Quotes,
Description-Zeilenumbruch, leer/fehlend, MDB→SQLite/DAO/Kopf und Quellhash.
Import/DAO2/2 (2,57s), App/lesender Helfer gebaut. Java nur Prüfdaten-Erzeugung,
keine Runtime-Abhängigkeit. SOURCE_FIELD_DISPLAY_PROPERTIES.md;456/384/72 gleich.
Fortsetzung: weitere belegbare Eigenschaften und allgemeine Entwurfsanforderungen
fortlaufend umsetzen; keine Zwischenabschlüsse oder Fertiganrufe.

Jet4 Required/AllowZeroLength werden jetzt im vorhandenen lesenden Metadatenlauf
als ausdrücklich vorhandene boolesche Properties übernommen. Tatsächliche
MDBTools-Repräsentation yes/no gegen unabhängig wieder geöffneten technischen
Jet4-Behälter geprüft; fehlende Properties bleiben unbekannt. Arbeitskopien
setzen NOT NULL und Text-/Memo-Leerstring-CHECK tatsächlich durch; DAO liest
Originalwerte readonly. GUI verwendet ausdrückliches nativeOwned statt aus
bekannten Getterwerten auf Bearbeitbarkeit zu schließen; eine Metadatenabfrage
pro Feldliste, importierte Entwurfsänderungen deaktiviert. Keine Originalmutation.
Technische Importprüfung1/1 (0,16s), anschließend bestehende DAO/GUI/echte
Kassa-Tabellenkopie3/3 (95,41s; davon Quellkopie87,28s); App/Helfer gebaut.
SOURCE_FIELD_DISPLAY_PROPERTIES.md. Keine neue Testregistrierung oder F-ID;
456/384/72 bleiben unverändert. Fortsetzung: eigene DAO-QueryDefs dauerhaft
umbenennen, Katalog-/View-/Generationseigenschaften mit berücksichtigen.

Eigene persistente DAO-QueryDef.Name lässt sich jetzt atomar umbenennen, auch bei
reiner Groß-/Kleinschreibungsänderung. Gesamte gespeicherte Definition bleibt
bytegleich; SQL, Parameterbindungen und aufrufereigene Field-/Parameterreferenzen
bleiben erhalten. Katalogcache folgt dem neuen Namen, verwaiste alte Einträge
geben nach Neuanlage nicht versehentlich die umbenannte Abfrage zurück. Andere
alte Namenshandles werden über Namespace-Generation3420 ungültig. Vorhandener
View-Installer läuft im selben Savepoint; Fehler stellt SQL und fünf Registrywerte
zurück. Kollision/ungültiger Name/readonly/aktive äußere Transaktion verweigert;
importierte/temporäre Namen und automatische Referenzkorrektur bleiben offen.
DAO1/1 (2,52s), App gebaut; keine Quellkopie-/Gesamtsuite unnötig wiederholt.
DAO_QUERYDEFS_CATALOG.md; F055/F121 offen,456/384/72 unverändert. Weiterarbeiten.

Datenbankfenster bietet jetzt Ansicht→Abfrage umbenennen für eigene Abfragen.
Offene Abfragedatenblätter vorher schließen; importierte Abfragen bleiben
unverändert/geschützt. Live-QueryDef wird über Namensdialog gehalten, damit
zwischenzeitliches Löschen/Neuanlegen nicht auf ein anderes Objekt umleitet.
Explorer/Selektion folgen tatsächlichem Namen. Tabellen-/Abfrageumbenennung
nutzen gemeinsamen Core-Helfer für vorhandene Druckseiteneinstellungen:
Schema geprüft, Layout unverändert verschoben, Zielkollision bricht gesamten
Savepoint ab. Kein QtGui im DAO-Backend. Bestehende DAO/GUI2/2 (7,84s) prüfen
auch tatsächliches Öffnen/Ausführen, Seitenbindung und atomare Teilfehler-Rücknahme.
App gebaut; langsame Originalkopie nicht ohne Importänderung wiederholt.
DAO_QUERYDEFS_CATALOG.md; F055/F080/F121 offen,456/384/72 unverändert. Weiterarbeiten.

DoCmd.Rename führt jetzt über tatsächliche eigene TableDef-/QueryDef.Name-Setter
Tabellen-/Abfrageumbenennung aus, einschließlich neu geordneter benannter
Argumente. ObjectType und OldName derzeit ausdrücklich erforderlich; Missing
bleibt auch bei Namen/Kommalücken Missing statt versehentlich acTable(Empty).
Null94, unbewiesener Kontext/Typ3251; geschlossene/beschäftigte/modale Projekte
und passende offene Dokumente werden geschützt. Aufrufer über skalare Auswertung
und Dispatch abgesichert. Reale VBA-Integrationsprüfung bestätigt beide Objekttypen,
resultierendes SQL, Seitenbindung und Fehlerfälle. DAO/GUI2/2 (7,77s), App gebaut.
Große DAO-/GUI-Prüfprogramme verwenden unter GNU/Clang Release jetzt -O1, um
Wiederübersetzungen zu verkürzen; Produkt-/Desktop-/Datenbibliotheken bleiben
nach tatsächlich geprüften flags.make bei -O3. Keine neue Testregistrierung.
DOCMD_RENAME_RUNTIME.md; F052/F055/F080/F121 offen,456/384/72 unverändert.
Fortsetzung: tatsächliches Entfernen eigener Datenbankobjekte samt abgeleiteten
Ansichten und Seitenmetadaten prüfen/umsetzen; kontinuierlich weiterarbeiten.

Eigene Tabellen/Abfragen können jetzt über Ansicht→Ausgewähltes Objekt löschen
(mit Nein als Vorgabe und über Dialog gehaltenem DAO-Handle) oder ausdrückliches
DoCmd.DeleteObject entfernt werden. Benannte Argumente/Missing/Null und passende
offene Dokumente geschützt; importierte Objekte im Fensterpfad deaktiviert.
QueryDefs.Delete entfernt Definition, vorhandene Seitenmetadaten und verwaltete
Views im selben Savepoint; Fehler stellt fünf Registrywerte wieder her, eigene
Workspace-Rücknahme bleibt erhalten. Native DROP/TableDefs.Delete räumt ebenso
Seitenmetadaten auf, damit Neuanlage keine Altzustände erbt. Unbekannte Katalogform
bricht gesamte Mutation ab; Roh-SQL-Rename mit Seitenzustand verweigert ohne
Rebinding. Keine automatische Löschung/Umschreibung abhängiger gespeicherter
Bindings; Relationenschutz bleibt. Bestehende DAO/GUI2/2 (8,06s), App gebaut:
echter VBA-Aufruf, Offen-/Missing-/Null-Schutz, GUI Nein/Ja, Views/Seiten/Objekte
wirklich entfernt, Namenswiederverwendung, Teilfehlerrücknahme, andere Daten gleich.
DELETE_OBJECT_RUNTIME.md; F052/F055/F080/F121 offen,456/384/72 unverändert.
Fortsetzung: vorhandene echte Access-Eval-Beobachtungen für allgemeine typisierte
Ausdrücke nutzen; VBA-/Eval-Kontext unterscheiden, keine pauschale Oracle-Parität.

Native Decimal-Variants (VarType14/CDec) implementiert: geprüfter96-Bit-Koeffizient,
Skala0..28, exakte Dezimalarithmetik/-Vergleiche, Vorzeichen/Abs/Fix/Int/Round,
Ganzzahl- und Currencykonvertierung ohne Double-Zwischenrunde; Variant-Arrays,
einmalige Skalarauflösung, homogene Berichtssummen/-Extrema/-Selektoren sowie
verlustfreie ungebundene Steuerelemente. Boost.Multiprecision nur Header>=1.74;
Buildabhängigkeit README/Workflow ergänzt, keine neue Shared-Library/Java.
Normales VBA, berechnete Controls und tatsächliches Application/Host.Eval bleiben
getrennte Kontexte. Eval übernimmt beobachtete Long-/Integer-/Decimal-Literale,
Typen, Präzedenz/Nz und Fehlerresultate. Acht Fehlerausdrücke über exaktenSHA256-
Fallabgleich mit sämtlichen112 technischen Collector-Ausdrücken wiedergewonnen;
keine Eingaben erfunden. CDec in beobachtetem Eval-Kontext verweigert2439,
Switch/Error-Variant6; gewöhnliche VBA-CDec-Konvertierung bleibt verfügbar.
Bereinigte echte Access11.0/DAO3.6/de-AT-Beobachtungen:104Werte plus8zugeordnete
Fehler, ohne private Quellmetadaten/Pfade/VBA. Bestehender DAO-Prüflauf bestätigt
56Zahlen-/Typbeobachtungen und8Fehler;48andere Werte bleiben ausdrücklich offen.
Core/DAO/GUI3/3 (9,49s), App/Backend3.54.1/JOIN/NULL/Counter geprüft; GUI-freie
Konfiguration erfolgreich, gelaufener Core-Binärpfad lautldd ohneQt/Boost-DSO.
Kein neuer CTest-Punkt/Katalogpunkt, keine neue CI-/Matrix-/Vollparitätsbehauptung.
NATIVE_DECIMAL_RUNTIME.md; exakte dbDecimal-Felder/Jet-SQL-Speicherung, vollständige
Locale/IEEE-/Automation-Identität/AST weiteroffen,456/384/72 unverändert.
Fortsetzung: die48weiteren echten Ausdrucksbeobachtungen und allgemeine
Sprach-/Datenbankanforderungen fortlaufend umsetzen; keine Zwischenabschlüsse.

Allgemeine Formatfunktionen erweitert: kombinierte Datum-/Zeitmasken, Monats-/
Wochentagsnamen, AM/PM, Literale/Datum-Escapes, feste/optionale Dezimalstellen,
Gruppierung/Prozent sowie FormatNumber/FormatCurrency mit benannten/ausgelassenen
Optionen und TriState. AktuelleQLocale liefert Zahlenkultur/regionale Erkennung;
Currency-Affixabstände werden als Legacy-Leerzeichen normalisiert, numerische
Gruppierungszeichen bleiben erhalten. MathematischeZahlenformatierung nutzt
Decimal, SQLiteINTEGER wird ausdrücklich ohne Double-Umweg gelesen (2^53+1 geprüft).
NativeSQL und Desktop teilen Formatter; reineCore-Hosts bekommen keinenQt-Zwang.
VBA.CDec unterstützt jetzt benanntesExpression-Argument. Eval-FormatNull liefert
beobachteten Leerstring, andere bestehendeVariant-/SQL-Nullregeln bleiben getrennt.
Gesamte vorhandene Ausdruckssuite abgeglichen:56Zahlen-/Typbeobachtungen+8Fehler
imDAO-Ziel sowie48weitereString-/Datums-/Locale-Beobachtungen über öffentliche
Application.Eval-Schnittstelle des tatsächlichenForm-/App-Objektmodells unterde-AT.
Insgesamt104beobachteteWerte+8Fehler; Doublefinanzwerte weitermitrelativerToleranz,
KalenderwertemitunabhängigerQt-Tages-/Sekundenberechnung. KeinepauschaleAST-/
OriginalVBA-/Locale-/Gesamtparität. Core/DAO/GUI/Formatter4/4 (9,49s), Appgebaut.
NATIVE_FORMAT_RUNTIME.md beschreibt Grenzen: komplexe/skalierende/wissenschaftliche/
Stringmasken, volleWindowsRegionaleinstellungen, negativeCurrencyprofile,
Originalfeld-Formatproperties und exakteNativeDecimal-/Currency-Datenbankfelder
weiteroffen. F014/F041/F068/F071/F118/F164;456/384/72 unverändert.
Kontinuierlich an allgemeinenDatenbankanforderungen weiterarbeiten, keinZwischenfinal.

Präzisionslücke in Str/Str$ geschlossen: Decimal-/Currency-Skalare werden einmal
aufgelöst und ohne Double-Zwischentext ausgegeben, einschließlich maximalem
96-Bit-Decimalwert und kleinster28erNachkommastelle. Currency-Konvertierung aus
Double wertet nun53-Bit-Mantisse/Exponent über ganzzahlige625/2er-Skalierung aus;
keine Abhängigkeit von plattformspezifischerlong-double-Präzision. Geprüft sind
exakte Halbwerte, direkt benachbarte IEEE-Werte, Vorzeichen, Subnormale sowie
signed64-Grenze/Nichtfinite/Überlauf. Die nachgewiesene REAL-Speichergrenze bleibt
unverändert; nativevolleCurrency-Datenbankfelder werden damit nicht vorgetäuscht.
Core/DAO/GUI/CurrencySQL/-Felder5/5 (10,03s), Appgebaut; bestehende112Ausdrucksfälle
weiterimDAO-/GUI-Prüfverbund. KeineOriginalkopie ohnebetroffeneImportänderung
wiederholt. Core-Prüfharness jetztwieDAO/GUI unterGNU/ClangRelease -O1; Runtime
lautflags.make weiter-O3. Keine neueTestregistrierung/F-ID;456/384/72 gleich.
Fortsetzung: Currency-Parameterwerte von begrenzterSQL-REAL-Kodierung trennen und
an allgemeinenexaktenDatenbanktypen fortlaufendweiterarbeiten.

Currency-Parameterwerte sind jetzt von der begrenzten SQL-REAL-Kodierung getrennt.
Value/defaultValue behält exakteCurrency-Werte einschließlich beider signed64-
Extrema undNull; ungültige Zuweisung ersetztvorigenWert nicht. ErstBind/Execute
prüft weiterdiebelegteREAL-Grenze undverweigertunsupportedWerte3251 ohnesieaus
demParameterzulöschen. Skalarstandardwerte werden einmalaufgelöst, danachwird
Generation/Lebensdauererneutgeprüft: einSQL-ändernderRückruf kannnichtdurchalte
ParameterinsneueSQLschreiben. VorhandeneParameter-/DAO2/2 (2,40s), Appgebaut;
Eingabe/Lesen/Überlauf-/SQL-Verweigerung/Null/Rückruf-/CloseundechteDatensätzegeprüft.
CURRENCY_PARAMETERS_RUNTIME.md. KeinexakterDatenbank-Currency-Speicherbehauptet;
456/384/72gleich. Fortsetzung: sichere reineCurrency-Parameterprojektionenohne
REAL-Umweg undallgemeineexakteDatenbanktypenweiterumsetzen.

Reine tabellenfreie Currency/Money-Parameterprojektionen liefern jetzt beide
signed64-Grenzwerte ohne REAL-Umweg: ganzzahlige Zehntausendstel mit eindeutigem
privatem Ergebnisdeskriptor. DAO behält Currency/Type5/Size8 und Null; Datenblatt
und Zwischenablage zeigen den exakten Betrag. Explizite Ausgabenamen verhindern
Verwechslungen mit SQLite-Parameternamen, auch bei Alias=Parametername.
SQL-Ausdrücke/FROM/WHERE/mehrdeutige Spaltennamen bleiben an der bisherigen
REAL-Grenze; veränderte Pläne sowie textuelle Filter/DAO-Find auf Tick-Ergebnissen
werden vor falschen Einheitenvergleichen verweigert. Bestehende Parameter-/DAO-
Prüfziele gemeinsam 2/2 bestanden (2,69s), Anwendung gebaut. Keine neue CTest-
Registrierung, kein neuer Katalogpunkt, keine native Currency-Felderstellung
oder allgemeine exakte Currency-SQL-Parität behauptet. 456/384/72 unverändert.
Details: CURRENCY_PARAMETERS_RUNTIME.md. Fortlaufend allgemeine native
Datenbank-/VBA-Anforderungen umsetzen; kein freiwilliger Zwischenabschluss.


Declared Currency For/Next counters now use exact signed64 ticks for the start,
frozen end/step and checked increment. Positive/negative four-decimal steps,
full-range boundaries, zero iterations, explicit zero step with Exit For,
body-written counters and the final post-loop value retain Currency kind.
Increment overflow raises6 before replacing the prior counter. Existing structured
loop stacks, Exit For and error recovery remain shared; no Double conversion is
used for this counter family. The existing execution budget still bounds zero-
step loops. Core/DAO checked together2/2 (3.99s), application built. Actual Access
2003 loop-oracle coverage, mixed Variant counter promotion and complete numeric
For/Next semantics remain open; no full parity assertion. Existing F041/F043/F045/
F105 only, fixed456/384/72 unchanged. Continue implementation without a final.

Native Zahlenmasken unterstützen jetzt bis zu drei Abschnitte für positive/
negative Zahlen und Zahlenwert0, leere Ersatzabschnitte, zitierte/escaped
Präfixe/Suffixe, negative Klammern und Prozent an beiden Seiten. Tausender-
Skalierung durch abschließende Kommata prüft exakte Rückrechnung und verweigert
Präzisionsverlust3251, statt still zu runden. .00 bleibt unterstützt. Gemeinsamer
Code für VBA/SQL/Formausdrücke; Masken-/SQL-Fälle und vorhandene GUI-Ausdrucks-
referenzen gemeinsam2/2 (6,01s), App gebaut. Keine zusätzlichen CTest-Ziele.
Wissenschaftliche/bedingte/Farb-/vierteNull-/interleavte Masken und vollständige
Windows-Formatgleichheit bleiben offen. Neue technische Prüfungen sind kein
zusätzlicher Windows-Access-Nachweis. NATIVE_FORMAT_RUNTIME.md; bestehende
F041/F068/F071/F118/F164,456/384/72 unverändert. Direkt weiterentwickeln.

FormatPercent ist jetzt als echte native VBA-/SQL-Funktion mit allen fünf
benannten/ausgelassenen Standardargumenten verfügbar. Die gemeinsame geprüfte
Maskenstrecke berücksichtigt Nachkommastellen, führendeNull, negative Klammern,
Gruppierung und lokale Prozentzeichen. Die benannte Format-Maske Percent liefert
zwei Nachkommastellen ohne Gruppierung. SQL-Integer bleiben vor der Skalierung
exakte Decimal-Werte; 9007199254740993 wird ohne Double-Verlust formatiert.
Null94, ungültige Optionen5 und Arity450 sind geprüft. Bestehender Formatter und
GUI gemeinsam2/2 (5,83s), App gebaut; native VBA-Positional-/Named-/Missing-Fälle
und vorhandene48 GUI-Referenzwerte enthalten. Keine neue Windows-Beobachtung
oder zusätzliche CTest-Registrierung. Decimal96/28-Grenze und vollständige
regionale Prozentprofile bleiben offen;456/384/72 unverändert.
NATIVE_FORMAT_RUNTIME.md; kontinuierlich allgemeine Datenbankfunktionen umsetzen.

Eigene native Tabellenfelder besitzen jetzt eine persistente DAO-Format-Text-
Eigenschaft mit derselben Detached-/Append-/Value-/Delete-/Identitätslogik wie
Caption/Description. Feldentwurf kann Masken setzen/löschen. Datenblatt zeigt
Formatwerte; SQL und EditRole bleiben unverändert. Zwischenablage behält den
bestehenden rohen typisierten Text. Druck/PDF nutzt eigene geprüfte formatierte
Werte: ungültige Masken zeigen #Fehler im Raster und brechen die Ausgabe vor
Erzeugung ab. Native Datum-/Zahlenmasken sind über echte DAO-/Designer-/Raster-
und PDF-Textextraktion geprüft, einschließlich beibehaltener SQL-Werte, Property-
Identität/Delete/Recreate und Rename/Schemaumbau. Fehlendes Include im GUI-
Prüfziel gezielt behoben; danach vorhandene DAO/GUI gemeinsam2/2 (8,00s), App
gebaut. Originale Format-Eigenschaften werden noch nicht als bekannt behandelt
(3251 statt erfundener Abwesenheit); Query-Lineage und Formular-/Berichtserbung
bleiben offen. NATIVE_FIELD_DISPLAY_PROPERTIES.md;456/384/72 unverändert.
Direkt als Nächstes proveniente Jet4-Formatübernahme mit technischer MDB prüfen.

Jet4-Format-Eigenschaften werden jetzt durch die bestehende schreibgeschützte
MDBTools-Grenze gelesen und mit eigener expliziter Format-Provenienz importiert.
DAO liest vorhandene Werte readonly383 bzw. meldet nachgewiesene Abwesenheit3270;
alte Metadaten ohne neue Format-Provenienz bleiben unbekannt3251. Datenblatt-
Formatierung verwendet diese tatsächlich gelesenen Masken, ohne Bearbeitungs-
oder SQL-Werte zu verändern. Die neue reine Entwicklungs-Javaquelle erzeugt eine
unabhängig wieder geöffnete technische Jet4-MDB mit null Zeilen/ohne VBA; nur die
importierte SQLite-Testkopie erhält Zahlen-/Datumsprüfwerte. Bestehende Import-/
DAO-Ziele gemeinsam2/2 (2,70s), zusätzlich frühere Caption-/Required-/AllowZeroLength-
Referenzdateien im selben Importziel; Herkunft, Typen, readonly, bekannte Abwesenheit,
Legacy unbekannt, ungültige Provenienz, Anzeige/Edit/Copy und unveränderter MDB-Hash
geprüft. App/Helper gebaut; vier Fixture-Hashes und drei Upload-Originalhashes OK.
Keine MDB oder Java-Abhängigkeit ins Produkt/Git übernommen. Keine Windows-
Render-/Vollparität;456/384/72 unverändert. SOURCE_FIELD_DISPLAY_PROPERTIES.md.
Kontinuierlich weitere allgemeine Datenbank-/VBA-Anforderungen umsetzen.

FormatDateTime unterstützt jetzt native VBA-/SQL-Aufrufe mit Date und optionalem
NamedFormat sowie vbGeneralDate/vbLongDate/vbShortDate/vbLongTime/vbShortTime.
Default und explizit ausgelassener NamedFormat wählen General Date; reine Zeit
und reines Datum werden getrennt dargestellt, sonst kombiniert. Lange Zeit nutzt
das tatsächliche QLocale-Profil (einschließlich US-PM), ohne Zeitzonentext; kurze
Zeit bleibt24-stündig. No-host-Syntaxprüfung darf geschachteltes CDate nicht
vorzeitig als Null konvertieren: gezielt behoben. CLDR-Schmal-/NBSP vor AM/PM
werden im nativen Legacy-Profil als Leerzeichen ausgegeben. Vorhandene Formatter/
GUI gemeinsam2/2 (5,57s), App gebaut; echte VBA Named/Missing/qualified-Fälle,
SQL1..2Argumente, Null94/ungültigeOption5/Arity450/ungültigesDatum13 und deutsche/
US-Zeitdarstellung geprüft. Keine zusätzliche CTest-Registrierung oder Windows-
Beobachtung. Volle regionale Datums-/Subsekunden-/Windows-Gleichheit weiter offen.
NATIVE_FORMAT_RUNTIME.md;456/384/72 unverändert. Kontinuierlich weiterentwickeln.

TextBox.Format wird jetzt aus SaveAsText übernommen und für berechnete Formular-
Textfelder sowie Berichtstextfelder tatsächlich gerendert. Getrennte typisierte
Werte bleiben maßgeblich: Anzeige-Fehler ändern Value nicht; Currency/Datum und
Long-Untertypen bleiben auch beim Löschen einer Maske erhalten. Native Berichte
erben proveniente direkte Feldmasken aus dem Berichtsentwurf und drucken sie
wirklich ins PDF. Bestehende gespeicherte Überschreibungen bleiben unabhängig.
Runtime-Formatlesen/-setzen/-löschen unterstützt diese Kontexte; ungültige Masken
lassen vorherige Eigenschaft stehen. Ausgabe verweigert Formatfehler vor falschem
PDF-Erfolg. Bearbeitbare Formularfelder behalten Masken nur mit explizitem3251-
Hinweis; deren Fokus-/Eingabepipeline ist nicht vorgetäuscht. CONTROL_FORMAT_RUNTIME.md.
GUI-Ziel bestanden (5,70s); anschließend Layoutziel einzeln bestanden (11,56s).
Der breitere Layoutverbund enthielt noch die alte Currency-For-Verweigerung:
gezielt durch echten Schleifenlauf mit zwei SQL-Schreibwirkungen und exaktem
Currency-Endzähler ersetzt. Keine neue CTest-Registrierung, keine frische gesamte
Suite/Windows-Gleichheit. App gebaut;456/384/72 unverändert. Ohne Zwischenfinal
weiter an den allgemeinen offenen Anforderungen programmieren.

General Number und Scientific sind jetzt native benannte Zahlenformate. General
Number behält Currency-/Decimal-Text exakt, trimmt überflüssige Nachkommastellen
und gruppiert nicht. Double/Single verwenden native15/7-signifikante Stellen.
Wissenschaftliche Einzelmasken 0.00E+00 bzw. e/E und +/- unterstützen optionale
Nachkommastellen, Exponentenpolsterung und exakte halbe Decimal-Rundung samt
Mantissenübertrag. Auch endliche Double-Werte außerhalb Decimal96 (1e100 und
Subnormale) können so formatiert werden. Vorhandene Formatter/GUI gemeinsam2/2
(5,79s), App gebaut; Max-Currency ohne Double-Umweg, größter Decimal, sehr kleine/
große Double, Rundung, Locale und echte SQL-Integer geprüft. Bisherige gezielte
Ausgabefehlerfälle verwenden weiterhin unsupported Farbmasken statt nun gültiger
wissenschaftlicher Masken. Keine neue Windows-Beobachtung/CTest-ID; kombinierte
wissenschaftliche Abschnitts-/Literal-/Engineering-Masken, vollständige regionale
Coercion und Access-Gleichheit bleiben offen. NATIVE_FORMAT_RUNTIME.md;456/384/72
unverändert. Ohne Zwischenabschluss direkt weiterimplementieren.

Textmasken unterstützen jetzt @ (mit Leerstellen), & (ohne Leerstellen), ! für
Linksfüllung, native Groß-/Kleinschreibung mit >/<, zitierte/escaped Literale und
einen zweiten Abschnitt für leere Strings/Empty. Null-Konventionen bleiben
getrennt. Tabellenanzeige und geerbte Berichtsformate verwenden diese Strecke;
Value und rohe Zwischenablage behalten den gespeicherten Originaltext. Native
Formatter/GUI gemeinsam2/2 (5,72s), App gebaut; echte SQL-Fälle, Füllrichtungen,
überzählige Zeichen, Leerwerte, Tabellen-/Berichtsmaske bei unveränderten Daten,
Unicode und vor unzulässigem Surrogat-Split verweigerte Ausgabe geprüft. NUL/
überlange Ausgabe sowie längenändernde Case-Konvertierung mit Platzhaltern werden
explizit verweigert; vollständige Windows-/Unicode-/Maskengleichheit bleibt offen.
Keine neue Windows-Beobachtung oder CTest-ID. NATIVE_FORMAT_RUNTIME.md;
456/384/72 unverändert. Nächste Umsetzung: Fokus-/Eingabestrecke formatierter
bearbeitbarer Formulartextfelder, ohne Anzeigen als Daten zurückzuparsen.

Bearbeitbare Formular-TextBoxen nutzen jetzt echte Formatdarstellung außerhalb
des Fokus und kanonische, ungerundete Eingabetexte im Fokus. Maßgeblich bleiben
typisierte Werte; nur tatsächliche Bearbeitung verwirft die Wertkopie. Formate
werden nicht als Daten zurückgeparst. Reine Darstellungswechsel erzeugen keine
Change-/Dirty-/SQL-Wirkungen. Doppelte Bindungen erhalten unabhängig ihre Masken;
verschiedene Werte mit identischer gerundeter Anzeige werden dennoch gespeichert.
Native Formularerstellung übernimmt proveniente Feldmasken. Bestehende Validierung,
Value/Text-Puffer, Undo und AfterUpdate bleiben wirksam; Maskenänderungen während
laufender Eingabe/BeforeUpdate verweigern3251. Skalarstandardwerte werden einmal
aufgelöst und Owner/Beschreibbarkeit erneut geprüft, geschlossener Owner2450.
Vorhandene GUI/Layout gemeinsam2/2 (17,22s), App gebaut. Tatsächliche Qt-Eingabe,
negative Abbruchvalidierung ohne Datenverlust, Undo, präzise SQL-Änderungen,
Text-Großdarstellung ohne Wertänderung und vollständige Uhrzeit hinter Datumsmaske
geprüft. Zunächst fehlende explizite Textgröße im technischen DAO-Aufbau gezielt
korrigiert; keine zusätzliche CTest-Registrierung. Regionale Eingabemasken,
beliebiges reentrantes Fokus-/IME-/Windows-Verhalten und vollständiger Access-
Formulardruck bleiben offen. CONTROL_FORMAT_RUNTIME.md;456/384/72 unverändert.
Ohne Zwischenfinal weitere allgemeine offene Funktionen umsetzen.

Formatierte native Suche ist in Tabellen-, Abfrage- und Formularansichten sowie
DoCmd.FindRecord integriert. Wörtliche String-Suche verwendet tatsächliche
Delegate-/TextBox-Darstellung, getrennten Wiederholungszustand und bei doppelten
Bindungen die gewählte Steuerelementmaske. Formatierte Suche über alle sichtbaren
gebundenen Textfelder ist möglich; Passwortdarstellung liefert keine Klartext-
Treffer. Rohwertsuche bleibt getrennt; formatierte Ersetzung verweigert3251.
Zell-/Zeitgrenzen, Null-/Binärausschluss und vorhandene Validierung gelten.
Vorhandene GUI/Layout gemeinsam2/2 (16,95s), App gebaut; tatsächliche Navigation,
Datumsmasken, doppelte Bindungen, Allfeldsuche und Passwortanzeige geprüft.
Ein veralteter geschlossener Dialogzeiger im Prüffall wurde gezielt korrigiert.
Keine neue CTest-ID oder Windows-Beobachtung; volle Jet-/Windows-Suchgleichheit
bleibt offen. FORMATTED_FIND_RUNTIME.md;456/384/72 unverändert.
Ohne Zwischenabschluss nächste Umsetzung: Bereichseinfügen mit neuen Datensätzen.

Native Tabellen-Bereichseinfügung kann jetzt über das Ende vorhandener Datensätze
hinaus neue Entwurfszeilen anlegen. Vorher werden noch nicht geladene vorhandene
Zeilen abgeholt und alle Zellen/Typen vorgeprüft; AutoWert-/Binärziele bleiben
geschützt. Neue Zeilen nutzen vorhandene Defaults, Datensatz-Undo, Verwerfen und
transaktionales Speichern mit echter Schlüsselerfassung. SQL-Constraintfehler
rollen alle SQL-Änderungen zurück; Entwürfe bleiben korrigierbar. Bestehender
Currency-/Feld-/Clipboard-Prüflauf1/1 (0,44s), App gebaut; Typfehler ohne Teiländerung,
zusätzliche Zeilen, Undo/Verwerfen, SQL-Defaults, erzeugte Schlüssel und Unique-
Rollback geprüft. Technische Schutzprüfung auf explizites AUTOINCREMENT korrigiert;
manuelle INTEGER-Primärschlüssel sind keine behaupteten AutoWerte. Kein neuer
CTest. Noch kein Einfügen in leere Tabellen/gesonderter Tabellen-Anfügebefehl;
weitere F340-Anforderungen bleiben offen. CLIPBOARD_APPEND_RUNTIME.md;
456/384/72 unverändert. Direkt weiterimplementieren.

Leere native Tabellen akzeptieren jetzt Texteingabe aus der Zwischenablage ab der
ersten sichtbaren beschreibbaren Spalte. Ein ausdrücklicher Kontextmenübefehl
und RunCommand acCmdPasteAppend(38) legen neue Datensatzentwürfe nach dem vollständig
geladenen Tabellenende an. Normales Einfügen behält die aktuelle Zelle als Ziel;
AutoWerte werden nicht überschrieben. Typisierte Formular-MIME wird ausdrücklich
verweigert, statt ihren Textkopf als Datensatz einzufügen. Vorhandene GUI/Feld-
Prüfungen gemeinsam2/2 (6,35s), App gebaut: leere Tabelle, Anfügen ohne Überschreiben,
SQL-Persistenz und tatsächlicher Hauptfenster-Dispatcher geprüft. F340 bleibt wegen
weiterer MIME-/Regional-/Undo-/Entwurfsszenarien offen. CLIPBOARD_APPEND_RUNTIME.md;
456/384/72 unverändert. Direkt weitere Clipboard-Teilanforderungen umsetzen.

Native Tabellen-Zwischenablage überträgt jetzt zusätzlich zum TSV eine typisierte,
begrenzte Rechteck-MIME. NULL, leerer Text, leere Binärwerte, Binärbytes, volle
SQLite-64-Bit-Ganzzahlen und vorhandene Currency-/Datumsdarstellung bleiben getrennt.
Binärdaten werden nur über validierte native Nutzdaten in Binärspalten geschrieben;
Text-/AutoWert-/Readonly-/verdeckte Zielgrenzen bleiben geschützt. Fehlerhafte MIME
wird nicht still als TSV interpretiert. Tabellen ohne SQL-Defaults benötigen keine
unnötige DAO-Rekonstruktion ihrer BIGINT-/BLOB-Felder. Für einen unveränderten
ungespeicherten Einfügebereich bestehen jetzt gemeinsame Strg+Z/Y-, Bearbeiten-Menü-
und Undo/Redo(292/199)-Wege; gespeicherte Date-Defaults werden beim Redo übernommen.
Spätere Fremdänderung, Requery oder erfolgreicher Save verwirft den Bereichsverlauf.
Vorhandene GUI/Feld/DAO gemeinsam3/3 (8,59s), App gebaut. Echte SQL-/Qt-Nachweise für
Bytes inklusive leerem BLOB, NULL/leer, große 64-Bit-Ganzzahlen, datumhaltige
Bereiche, Anfügen, leere Tabelle, Undo/Redo und Verlaufverfall; ungültige MIME,
AutoWert und verdeckte Ziele verweigert. Zwei veraltete Binär-Copy-Negativannahmen
angepasst; Qt-interner Schreibschutz während geprüfter Binärübernahme begrenzt
freigegeben, außerhalb weiterhin wirksam. Keine neue CTest-ID oder Windows-
Beobachtung. Kein VBA-Binary-Variant, keine allgemeine SQL-Rücknahme; Einfügen über
fremde Entwürfe, volle Regional-/Access-/MIME-Gleichheit bleiben offen. F340/F168/
F016 unverändert; TYPED_GRID_CLIPBOARD.md;456/384/72 unverändert. Direkt weiter.

Kreuztabellen mit TRANSFORM Count(Feld)/Count(*) und festen Textüberschriften in
PIVOT ... IN(...) laufen jetzt über denselben nativen Compiler, direkte DAO-
Recordsets, parametrisierte QueryDefs und gespeicherte Abfragen. Zählungen lassen
NULL-Feldwerte aus, Count(*) zählt Zeilen; fehlende Kategorien ergeben0. Reihenfolge,
Aliases und Fallkollisionsschutz der Ausgabespalten werden erhalten, Count-Spalten
sind provenient INTEGER4/DAO Long, QueryDef.Type ist dbQCrosstab16. Simple gruppierte
Zeilenfelder, einzelner Quellname mit optionalem Alias sowie WHERE/HAVING/ORDER
bleiben begrenzt; numerische Aggregate, dynamische/numerische Überschriften und
allgemeine Join-/Ausdrucks-Kreuztabellen verweigern3251. Vorhandene Query-Resultat/
Schema-/DAO-Prüfungen gemeinsam3/3 (2,44s), App gebaut. Zusätzlich vorhandene SQL-
Dokument-/DAO-Prüfungen2/2 (2,37s); keine neue CTest-ID. Klammer-/Nichtklammernamen
werden semantisch gleich verglichen. Kein Windows-Ergebnis-/Kollationsorakel und
keine volle Kreuztabellenparität behauptet. Siehe COUNT_CROSSTAB_RUNTIME.md.
F068/F067 und456/384/72 bleiben unverändert. Vier Repository-Originale und drei
Upload-Originale aktuell per SHA unverändert. Ohne Zwischenabschluss fortsetzen.

Feste Count-Kreuztabellen unterstützen jetzt zusätzlich vorzeichenbehaftete
Long-Überschriften und reine Year(Feld)/Month(Feld)/Day(Feld)-Pivot-Ausdrücke.
Integer-Spaltennamen werden kanonisch gebildet;01/1 und Text-/Zahl-Namenskonflikte
werden vor Ausführung verweigert. Nichtintegrale, außerhalb Long liegende und
beliebige andere Pivot-Ausdrücke bleiben3251. Vorhandene Query-Resultat-/Schema-
Prüfungen gemeinsam2/2 (0,12s), App gebaut: echte Datumszeilen mit Uhrzeit,
Schalttag und NULL, vorgegebene Monats-/Jahresreihenfolge, fehlende Kategorie,
Signum und Negativfälle geprüft. Keine neue CTest-ID/Windows-Beobachtung;
COUNT_CROSSTAB_RUNTIME.md; F068/F067 und456/384/72 unverändert. Direkt weiter.

Vollständige native Tabellenzeilen können jetzt über ausdrückliches Anfügen mit
neuen AutoWerten übernommen werden. Die Grid-MIME enthält physische Feldnamen;
Schemaform, Namen und Reihenfolge müssen vollständig passen. Nur provenient
generierte Spalten werden ausgelassen, manuelle Primärschlüssel bleiben echte
Werte mit ihren SQL-Constraints. Die neue Schlüsselvergabe erfolgt ausschließlich
beim normalen transaktionalen Save. Die Regel bleibt bei Bereichs-Undo/Redo erhalten.
Vorhandene GUI/Feld-Prüfungen gemeinsam2/2 (7,05s), App gebaut: vollständiger echter
Datensatz mit Binärbytes und großer Ganzzahl angefügt, Undo/Redo, tatsächlicher
neuer SQL-Schlüssel und verweigerte fremde Feldform geprüft. Positionsbereiche und
ältere Grid-MIME bleiben getrennt; anders geordnete/partielle Namenszuordnung und
Formular-MIME-Interop bleiben offen. TYPED_GRID_CLIPBOARD.md; F340/F168 und456/384/72
unverändert. Direkt weitere Bearbeitungsanforderungen umsetzen.

Saubere native Tabellen-Datenblätter unterstützen jetzt Ausschneiden sichtbarer
zusammenhängender Zellbereiche per Strg+X/Umschalt+Entf, Bearbeiten-/Kontextmenü und
RunCommand Cut189. Die typisierte Zwischenablage erhält die geprüften Quellwerte;
danach werden nur Cache-Zellen auf SQL NULL gesetzt. SQL bleibt bis zum normalen
transaktionalen Save unverändert. Dieselbe begrenzte Bereichs-Undo/Redo-Strecke
stellt die Werte wieder her. AutoWert-/Readonly-/verdeckte Ziele und fremde
Entwürfe bleiben geschützt. Vorhandene GUI/Feld-Prüfungen gemeinsam2/2 (6,91s),
App gebaut: echtes Menü/Dispatcher/Qt-Tastatur, Binär-Cut ohne frühe SQL-Änderung,
Undo/Redo, persistentes NULL statt leerem Text und verweigerter AutoWert-Cut bei
unveränderter Zwischenablage geprüft. Keine neue CTest-ID oder Windows-Beobachtung.
Dies ist Zellbereichsleeren mit Undo, keine Datensatzlöschung oder volle Access-
Cut-/Ereignisparität. TYPED_GRID_CLIPBOARD.md; F340/F168 und456/384/72 unverändert.
Ohne Zwischenabschluss weiterimplementieren.

Bereichseinfügen kann jetzt vorhandene ungespeicherte Zelländerungen physischer
Tabellenzeilen erhalten. Vorprüfung erfasst rohe Werte und geänderte Feldpositionen;
Undo stellt diese vorherigen Entwürfe wieder her, Redo wendet nur den neuen Bereich
an. Fehlgeschlagene Bereichsänderungen nutzen dieselbe Cache-Rücknahme, ohne SQL-
Schreibvorgang oder automatisches Speichern. Vorhandene GUI/Feld-Prüfungen gemeinsam
2/2 (6,28s), App gebaut: aufeinanderfolgende echte VBA-/UI-Pastes, ungültige Eingabe
bei erhaltenem vorherigem Entwurf/Undo, Currencywerte, separater Textentwurf und
Binäränderung mit leerem BLOB vor/nach Undo/Redo geprüft. Veraltete globale Dirty-
Negativannahmen der bestehenden Prüfungen angepasst. Neue/gelöschte Datensatzentwürfe
und mehr als1M geladene Zellen bleiben3251; Cut benötigt weiter einen sauberen Cache.
Keine neue CTest-ID/Windows-Beobachtung. TYPED_GRID_CLIPBOARD.md; F340/F168 und
456/384/72 unverändert. Ohne Zwischenabschluss weiterimplementieren.

Bereichseinfügen erhält jetzt auch bestehende neue Datensatzentwürfe und vorgemerkte
Löschungen. Cache-Rücknahme rekonstruiert eingefügte Zeilen an ihren Positionen,
übernimmt ihre tatsächlich gesetzten Felder, lässt SQL-Defaults ungesetzt und
stellt Löschmarker wieder her. Direkte Ziele in gelöschten Zeilen bleiben2448.
Bereichs-Undo/Redo stellt nach dieser Rekonstruktion eine gültige aktuelle Zelle
wieder her. Vorhandene GUI/Feld-Prüfungen gemeinsam2/2 (6,19s), App gebaut: früherer
Insert plus nachfolgende Bereichseingabe, Qt-Ctrl+Z/Y bei erhaltenem Zellziel,
tatsächliche spätere SQL-Neuanlage einschließlich Default, vorgemerkte Löschung,
verweigerter gelöschter Zielbereich und gemeinsamer SQL-Save geprüft. Erste Strecke
2/2 (6,36s), danach Zellziel-/Defaultnachweis gezielt ergänzt; keine neue CTest-ID.
Kein vorzeitiger SQL-Write, keine neue Windows-Beobachtung; beliebige externe
Index-/Reentranz-/Konkurrenz-/Access-Ereignisgleichheit bleibt offen.
TYPED_GRID_CLIPBOARD.md; F340/F168 und456/384/72 unverändert. Direkt weiter.

Native einfache numerische For-Zähler unterstützen jetzt Byte und Single zusätzlich
zu Integer/Long/Double. Grenzen/Schrittweite werden einmal ausgewertet und nicht
auf den Zählertyp eingeengt; Start/Next werden regulär Let-zugewiesen, danach wird
der tatsächlich zugewiesene (ggf. Single-gerundete) Zähler verglichen. Integer-/Long-
Bruchgrenzen, Byte-Rückwärtsschritte und Step0 mit Exit For funktionieren; Overflow6
ersetzt den letzten angenommenen Byte-/Single-Wert nicht. Deklariertes Date ist
laut MS-VBAL kein numerischer Zählertyp und bleibt vor SQL-Präfix verweigert3251.
Quelle: veröffentlichte MS-VBAL v20260818 §5.4.2.3/5.5.1.2.1, PDF-SHA dokumentiert;
keine neue Windows-/Access2003-Beobachtung. Vorhandene Core/DAO/Layout gemeinsam3/3
(17,41s), App gebaut: Rundung vor Anfangsvergleich, Endtypen Single/Double,
Bruchgrenzen, eingefrorene Grenzen, Body-Zähleränderung, Zero-Step, Overflow und
tatsächliche SQL-Schleifenaufrufe. Bestehende globale Bruchgrenzen-/Single-/Byte-
Negativannahmen gezielt in reale positive Fälle geändert; keine neue CTest-ID.
Volle gemischte Variant-/Fixedpoint-/Fehler-/Compilerparität bleibt offen.
TYPED_FOR_RUNTIME.md; F045/F041/F122 und456/384/72 unverändert. Direkt weiter.

Tabellennavigation kann jetzt von einer sauberen aktuellen Zeile wechseln, ohne
unbeteiligte Cache-Entwürfe zu speichern oder zu verwerfen. Native Pfeiltasten und
GoToRecord/RunCommand nutzen denselben Pfad. Ein einzelner befüllter neuer Datensatz
kann vor dem Verlassen über den normalen transaktionalen Save mit echter
Schlüsselerfassung gespeichert werden; vollständig leere neue Entwürfe bleiben
begrenzt, neue Zeilen werden bereits für Record-Undo als pending erkannt.
Der Shared-Find-Guard erlaubt ebenfalls eine saubere aktuelle Zeile bei anderen
Entwürfen. Andere Guard-Aufrufer behalten ihre bisherigen Insert-Grenzen.
Vorhandene GUI/Feld gemeinsam2/2 (6,20s), App gebaut: tatsächliche Qt-/VBA-Navigation,
erhaltene mehrere Entwürfe, unverändertes SQL für unbeteiligte Änderungen,
verweigerter leerer Insert, echter einzelner Insert beim Wechsel mit neuem
SQL-Schlüssel; Mehrfach-Save-/Sort-/Requery-Sperren weiterhin geprüft. Keine neue
CTest-ID oder Windows-Beobachtung; TABLE_NAVIGATION_DRAFTS.md. Aktueller gelöschter
Datensatz, nötige Teil-Saves mehrerer Entwürfe, allgemeine Addition-Row-/Windows-
Ereignisse bleiben offen. F368/F168/F080 und456/384/72 unverändert. Direkt weiter.

Native Find-Dialoge von Tabellen, Abfragen und Formularen unterstützen jetzt einen
expliziten Platzhaltermodus, getrennt von wörtlicher Suche. DoCmd.FindRecord
String-Muster aktivieren ihn bei * ? # oder [. Syntax/Matching teilen den bestehenden
VBA-Like-Parser und UTF16-DP-Matcher, einschließlich Zeichenklassen/Negation,
Whole/Prefix/Substring-Grenzen, Case-Modus und gespeichertem FindNext-Zustand.
Formatierte Suche behält reale Delegate-/Kontrollmaske und Passwortdarstellung;
typisierte Gleichheit bleibt getrennt, Musterersetzung verweigert3251. Musterfehler93
und absolute Patternlänge werden vor Navigation/Save geprüft; pro Kandidat gelten
1M Input×Pattern-Units,16384 Pattern-Units und Scan-Zell-/Zeitgrenzen. Bestehende
Navigation kann beim Matchscan den angenommenen aktuellen Datensatz speichern.
Vorhandene Core/GUI/Layout gemeinsam3/3 (19,13s), App gebaut: tatsächliche VBA-/Qt-
Muster, Ziffern/Klassen/Negation, Text-/Datumsmasken, Literal-Umschaltung, Query,
Fehler-/Größengrenzen und unveränderte SQL-Daten geprüft. Zunächst kollidierenden
lokalen Testvariablennamen korrigiert; veraltete globale Wildcard-Negativannahme
gezielt ersetzt. Keine neue CTest-ID/Windows-Beobachtung. Nicht-ASCII-Textranges/
Surrogatvergleiche und volle Jet-/Locale-/Allfeld-/Error/Event-Gleichheit bleiben
offen. WILDCARD_FIND_RUNTIME.md; F373/F071/F130 und456/384/72 unverändert. Direkt weiter.

Unformatierte DoCmd.FindRecord-Suche unterstützt jetzt OnlyCurrentField:=acAll0
über die jeweils passenden unterstützten Feldtypen. Text sucht Textfelder,
native Number/Currency/Date/Boolean verwenden ihre vorhandene geprüfte Gleichheit;
typisierte Werte benötigen weiter Whole-Field-Match. In direkten Tabellenformularen
ist der gerade fokussierte Steuerelementtyp im Allfeldmodus nicht mehr maßgeblich.
Tabellen/Readonly-Query-Resultate verwenden den bestehenden Allspalten-Scanner,
Formulare ihren begrenzten Datensatzbestand; formatierte Allfeld-Strings bleiben
getrennt. Vorhandene GUI-Prüfung1/1 (5,90s), App gebaut: tatsächliche VBA-/Qt-
Textsuche vom fokussierten Zahlenfeld, numerische Suche aus Textfokus, vollständiges
Datum mit Uhrzeit sowie Tabelle mit gefundenem Zelltyp/Spaltenwechsel geprüft.
Keine neuen CTest-/Windows-Beobachtungen; keine unbekannten Typen aus Zelltext
erraten. Volle Cross-Type-Coercion, Allfeld-/Spalten-/Fokus-/Ereignis-/Locale-
Gleichheit bleibt offen. ALL_FIELD_FIND_RUNTIME.md; F373/F071/F130/F050 und456/384/72
unverändert. Ohne Zwischenabschluss weitere allgemeine Funktionen umsetzen.

Wörtliches Tabellen-Ersetzen unterstützt jetzt Whole-/Teiltext-/Präfixmodus.
Current ersetzt das erste Vorkommen im Entwurf, Alle ersetzt alle nichtüberlappenden
Vorkommen jedes passenden Feldes bzw. einen Präfix/Whole-Wert. UI übergibt den
gewählten Vergleichsmodus tatsächlich; formatierte/wildcard/typisierte Rückersetzung
bleibt verweigert. Leerer Ersetzungstext ist echte Leerzeichenfolge, kein SQL NULL.
Batch-Ausgaben werden vor Cache-Änderung vollständig vorbereitet,16Mi UTF16-Units/
1M Zeilen/5s-Scan begrenzt und gemeinsam transaktional geschrieben; SQL-Fehler rollen
alle Änderungen zurück und hinterlassen keine Batch-Entwürfe. Vorhandene GUI/Feld
2/2 (6,26s), App gebaut: echte Dialoge, Case-insensitive Mehrfachvorkommen,
Prefix/current-erstes-Vorkommen, getrennte SQL-Save-Grenze, leere Textwerte,
pro Zeile verschiedene Ausgaben, vollständiger Check-Constraint-Rollback und NUL-
Verweigerung geprüft. Keine neue CTest-ID oder Windows-Beobachtung.
LITERAL_REPLACE_RUNTIME.md; volle Match-Offset-/Highlight-/Wildcard-/Locale-/Event-
Gleichheit bleibt offen; F373/F340/F168 und456/384/72 unverändert. Direkt weiter.

Explizites natives Datensatz-Anfügen ordnet Grid-MIME-Felder jetzt nach eindeutigen
Namen zu, auch bei anderer Reihenfolge und Teilmengen. Ausgelassene Felder bleiben
ungesetzt und erhalten SQL-Defaults; ausdrücklich geliefertes NULL überschreibt
Defaults. Nur proveniente AutoWerte werden ausgelassen. Name-/Formprüfung und
Expandierungsgrenze erfolgen vor Cacheänderung; die Supplied-/Default-Regel bleibt
bei Undo/Redo erhalten. Normales Paste/alte namenlose MIME bleiben positionsbezogen.
Vorhandene GUI/Feld gemeinsam2/2 (6,21s), App gebaut; zusätzlich Default/NULL-Nachweis
gezielt1/1 (0,48s): anders geordnete Zieltabelle mit zusätzlichem Defaultfeld,
AutoWert an anderer Position und echter Folge-ID11/12, Binärbytes, voller SQLite-
Integer, vollständiges Datum, Teilrecord mit ausgelassenen NULL-Feldern, explizites
NULL bei existierendem Textdefault, Undo/Redo und verweigerte doppelte Namen.
Keine neue CTest-ID oder Windows-Beobachtung. TYPED_GRID_CLIPBOARD.md; externe/
Formular-MIME-/Access-Ereignisgleichheit bleibt offen. F340/F168/F173 und456/384/72
unverändert. Direkt weiterimplementieren.


Readonly-Abfragedatenblätter veröffentlichen jetzt neben TSV das vorhandene typisierte
Grid-MIME mit Feldnamen. Tatsächliche rohe Querywerte erhalten Binärbytes, NULL,
Leertext/leeres BLOB und volle SQLite-Ganzzahlen; TSV verwendet Hex-Binärtext und
exakte Ganzzahldarstellung. Proveniente CURRENCY_TICKS-Ergebnisse erhalten ein
getrenntes Currency-Tag mit kanonischen 64-Bit-Ticks, keine Umdeutung als Ganzzahl.
Paste wandelt Currency vor Cacheänderung in exakten Text oder verlustfrei begrenzte
bestehende Currency/Double-Speicherung um; andere Zieltypen werden verweigert.
Keine physische CURRENCY_TICKS-Unterstützung oder externe Access-MIME-Parität.
Vorhandene GUI/Feld-Prüfungen gemeinsam2/2 (6,42s), App gebaut: tatsächliches
Query→anders geordnetes Tabellenziel, SQL-Default, volle Integer/Binär-/NULL-
Unterscheidung, beide Currency-64-Bit-Grenzen nach Text, kleine Currency nach REAL,
Undo/Redo/Save und vollständige Verweigerung großer Currency→REAL ohne Entwurf/
SQL-Änderung. Fehlendes QMimeData-Include beim Build gezielt korrigiert. Keine neue
CTest-ID/Windows-Beobachtung. TYPED_GRID_CLIPBOARD.md; F340/F168/F173/F014 und
456/384/72 unverändert. Direkt weiterimplementieren.


Bestehende aktuelle Tabellenzeilen können bei Navigation/Find jetzt einzeln
transaktional gespeichert werden, während andere Update-/Insert-/Delete-Entwürfe
im Qt-Cache unverändert bleiben. Bei mehreren Entwürfen sind geänderte/NULL-
Primärschlüssel und neue/gelöschte aktuelle Zeilen weiter verweigert; alleinige
populierte Inserts behalten den bisherigen Save-/Key-Pfad. Nach Commit aktualisiert
Qt selectRow nur die gespeicherte Zeile; Cache-Undo wird an der SQL-Grenze ungültig.
Vor Commit verschwundene/trigger-umgeschlüsselte Schlüssel führen zum Rollback.
GUI/Feld gemeinsam2/2 (6,21s), App gebaut: echte VBA-Navigation/Find, andere Entwürfe
und unveränderte Fremdzeilen, gemischte Update/Insert/Delete-Caches, Constraint-
Rollback und Trigger-Rekey-Rollback geprüft. Veraltete Find-Negativannahme und
Testversuch, ein geschütztes AutoWertfeld zu bearbeiten, gezielt korrigiert;
manuelles Text-PK separat geprüft. Uploads3/3/Fixtures4/4 hashgleich. Keine neue
CTest-ID/Windows-Beobachtung. TABLE_NAVIGATION_DRAFTS.md; vollständige aktuelle
Insert/Delete-/Fokus-/Ereignis-/Addition-Row-Parität bleibt offen. F368/F168/F080 und
456/384/72 unverändert. Direkt weiter.


„Neuer Datensatz“/acNewRec spart jetzt das pauschale Speichern fremder Entwürfe:
sauberer aktueller Datensatz darf bei weiteren Entwürfen eine neue Zeile öffnen;
bestehender geänderter aktueller Datensatz nutzt den gezielten Save-Pfad. Ein
alleiniger gefüllter Insert wird regulär gespeichert, bevor die nächste neue Zeile
entsteht. Leere aktuelle Inserts oder selektive Inserts mit weiteren Entwürfen
bleiben verweigert. Bestehende GUI-Prüfung1/1 (6,42s), App gebaut: tatsächliche
RunCommand/VBA-Aufrufe, SQL-Abgrenzung, anderer Entwurf, neuer Zeilenfokus und Escape
ohne fremden Datenverlust. Keine neue CTest-ID/Windows-Beobachtung. F368/F173/F168
und456/384/72 unverändert. Direkt weiter.


F088 hat jetzt einen ersten ausführbaren nativen Makroumfang: echter +Makro-
Aktionseditor (anlegen/entfernen/umsortieren/speichern/wiederöffnen), Explorergruppe,
DoCmd.RunMacro mit benannten Parametern und begrenzter Wiederholung. Acht Aktionen
verwenden vorhandene DoCmd-Pfade; StopMacro beendet einschließlich Wiederholungen.
Gesamte Struktur/Arity/Typen/Attribute werden vor erster Aktion geprüft, unbekannte
Bedingungen verweigert. Aktionen behalten ihre jeweiligen Transaktionsgrenzen;
späterer Laufzeitfehler rollt frühere erfolgreiche Aktionen nicht zurück. Diagnostik
bewahrt Fehlernummer/-metadaten mit Makroname/Aktionsnummer. AllMacros bietet den
bewachten vorhandenen Projektkatalog; SysCmd unterstützt nativen acMacro-Zustand.
GUI/Layout gemeinsam2/2 (17,16s), App gebaut: echter Editor/Reopen, gespeicherte
Definition, tatsächlicher benannter VBA-Aufruf mit dreifacher SQL-Wirkung,
StopMacro, präventiv verweigerte spätere unbekannte/bedingte/falsch typisierte
Aktionen, fehlendes Makro2485, bestehende SQL-Fehlernummer3078 mit Aktion2 und
bereits erfolgreich geschriebenen früheren Zeilen, Wiederholungsgrenze und
Katalog Count/Item/Enumeration/Name/IsLoaded geprüft. Anfänglich zu generische
SQL-Fehlernummer3146 im Test auf den tatsächlichen bestehenden3078-Pfad korrigiert.
Keine neue CTest-ID/Windows-Beobachtung. NATIVE_MACROS_RUNTIME.md beschreibt eigene
Format-/Budgetgrenzen; kein Access-Makroimport. Gruppen, Bedingungen/Ausdrücke,
RepeatExpression, Eventbindung, volle Aktionssprache und Oracle bleiben offen.
F088/F050/F057/F148 und456/384/72 unverändert. Direkt weiter programmieren.


Gespeicherte native Makronamen funktionieren jetzt auch als Ereignisbindungen:
vorhandene generische Steuerelementpfade (u.a. Click/DblClick/Change), Kontroll-
BeforeUpdate/AfterUpdate sowie Form-/Berichtsereignisse und Timer verwenden einen
geteilten Dispatcher für [Event Procedure], =Funktion() und benannte Makros.
OnActivate verschiebt Current bei jeder vorhandenen Bindung wie bisher bei VBA.
Makrofehler werden als echte Ereignisergebnisse/Statusdiagnose gespeichert;
fehlendes Kontroll-BeforeUpdate-Makro verhindert Save mit unveränderten SQL-Werten.
GUI/Layout gemeinsam2/2 (17,33s), App gebaut: tatsächliches Load/Current, Qt-Klick,
Doppelklick, Texttastatur/Change, Before/AfterUpdate mit SQL-Persistenz, tatsächliches
Timer-Timeoutsignal, Close sowie fehlender Click/BeforeUpdate-Name2485 nachgewiesen.
Timerprüfung prüft Signaldelivery, keinen neuen Wandzeit-/Access-Timing-Oracle.
Makronamen mit führendem =/[ sind zur Abgrenzung von Ausdrücken/Embedded-Makros
verweigert. Keine neue CTest-ID/Windows-Beobachtung. CancelEvent, Tastatur-
Makroargumente, Gruppen/Bedingungen/volle Ereignisgleichheit bleiben offen.
NATIVE_MACROS_RUNTIME.md; F088/F050/F057/F148 und456/384/72 unverändert. Direkt weiter.


Native Makros/DoCmd unterstützen jetzt CancelEvent ohne Argumente in unterstützten
abbrechbaren Bindungen (Open/Unload/NoData/Dirty/BeforeInsert/BeforeUpdate/DblClick).
Der verschachtelt restaurierte Ereigniszustand wird in VbaExecutionResult.cancelled
übernommen; Folgeaktionen laufen normal weiter. Außerhalb unterstützter Ereignisse
bleibt3251, zusätzliche Argumente450. GUI/Layout2/2 (17,08s) bestätigen echte
Kontroll-BeforeUpdate-Verweigerung ohne SQL-Save, Open-/Unload-Abbruch und keinen
über den Ereignisrahmen hinaus verbleibenden Cancel-Zustand. App gebaut; die neue
Designerzuordnung liegt separat noch als WIP vor.
Zusätzlich bestehender Original-Kassa-Lauf gezielt1/1 (114,46s) bestanden. Dabei
alte Rohtextannahmen für tatsächlich formatierte Original-Form-/Berichtsfelder und
Summen korrigiert: gemeinsame Prüfhilfe vergleicht reale skalare Werte mit SQL und
Anzeige separat über den bestehenden Formatter. SQLite-Ausdrucksmetadaten sind
kein voller Access-Varianttyp-Oracle. Eigenen anfänglich falschen Objektzugriff in
der Prüfhilfe behoben; keine Produktionsformatierung zur Erzwingung alter Erwartungen
abgeschaltet. Ungültige explizite WhereCondition meldet den tatsächlichen
Datenquellenfehler vor nachfolgenden Initialisierungsfolgefehlern; die bekannten
beiden fehlende-Verkaufs-ID-Fälle liefern wieder3075/Zeile1023 statt überlagertem3146.
Das ist Fehlerpriorität nach Konstruktion, keine Zusicherung ausbleibender früherer
Handlerwirkungen. Originalverkaufsworkflow bleibt offen. Uploads3/3 und Fixtures4/4
erneut hashgleich. Keine neue CTest-ID oder Windows-Beobachtung; kein frischer
Gesamtsuite-/Plattform-/CI-Nachweis. NATIVE_MACROS_RUNTIME.md; F088/F057/F114/F129/
F148 und456/384/72 unverändert. Direkt weiter.


Native Formulare können Makros jetzt über „Ereignisse zuordnen...“ an unterstützte
Form-/Feldereignisse binden. Automatisch/Makro/Funktion/[Event Procedure]/Keine
Bindung sind echte gespeicherte Alternativen; „Keine“ unterdrückt automatische
VBA-Erkennung. Unterdialog arbeitet auf Kopien, Cancel verwirft, OK validiert,
Eltern-Save persistiert. NativeStoredForm prüft/emittiert events/field_events/
timer_interval, unbekannte/doppelte Ereignisse/Felder und ungültige Werte werden
verweigert; keine Ausführung beim Save. Makronamenauflösung und Save-Konfliktprüfung
verwenden begrenzten Qt-Case-insensitive-Katalogscan (1M Namen/5s), danach exakten
Definitionslookup; Ä/ä und mehrdeutige Namen sind tatsächlich geprüft, keine
Jet-/Locale-Gleichheit behauptet. Eigenständiges VbaValue-Include korrigiert.
GUI/Layout zusammen2/2 (18,42s), App gebaut: echte verschachtelte Dialoge,
Cancel/Reopen/OK, gespeicherte Definition/Layout, Load-SQL-Wirkung, per Designer
gebundener BeforeUpdate-Abbruch mit unverändertem SQL, explizit deaktivierte
vorhandene Form_Load-Prozedur, unbekanntes Ereignis/Feld und negativer Timer.
Keine neue CTest-ID/Windows-Beobachtung. Originallauf1/1 (114,46s) ist der zuvor
geprüfte Stand, kein neuer Gesamtsuite-/Plattformnachweis. NATIVE_FORM_EVENT_DESIGNER.md;
F080/F088/F050/F057/F130/F148 und456/384/72 unverändert. Direkt weiter.


Native Makroaktionen unterstützen jetzt optionale Bedingungsausdrücke und
verschachtelte RunMacro-Aktionen. Der echte Aktionseditor speichert/verschiebt die
Bedingung mit der Aktion. Syntax/Typ/UTF16-/4096-Unit-Grenzen werden für die ganze
Definition vor ihren Aktionen geprüft; anschließend genau eine Auswertung pro
besuchter Aktion/Iteration im vorhandenen Eval-Kontext. False/Null überspringen.
Öffentliche Standardmodulfunktionen bleiben vom CodeBehind getrennt; anfangs im
falschen Kontext angelegte technische Prüffunktion in echten Standardmodul-Artefakt
verschoben, keine privaten/öffentlichen Zugriffsgrenzen aufgeweicht. Nested RunMacro
teilt16 aktive Aufrufe/10000 besuchte Aktionen, restauriert Budgets bei Fehlern;
StopMacro beendet nur den aktuellen Aufruf samt Wiederholungen, Eltern können
weiterlaufen. Vorherige SQL-Wirkungen werden bei späterem Laufzeitfehler nicht
pauschal zurückgerollt. RepeatExpression/Gruppen/volle Macro-Expr-/Event-Parität offen.
App gebaut; Layout1/1 (11,21s) im initialen Bündel; nach korrigiertem Modulkontext
nur betroffene GUI-Prüfung wiederholt1/1 (6,43s): echte Editorbedingung,
True/False/Null, exakt dreifache öffentliche Function-SQL-Wirkung, ungültige spätere
Bedingung ohne Prefix-SQL, Nested-Wiederholung, Child-Stop/Parent-Fortsetzung,
Rekursionsschutz/anschließende Recovery sowie echte bedingte Formularvalidierung
(blocked verhindert Save, allowed wird gespeichert). Keine neue CTest-ID/Windows-
Beobachtung, kein neuer Vollsuite-/Plattform-/CI-Nachweis.
Aktuelles .deb unter /workspace/work/native-conditional-macros/package/linaccess-
0.7.0-Linux.deb gebaut; vorhandene Paketprüfung bestätigt App-/MDB-Helfer-Build-IDs,
bytegleichen SQLite3.54.1-Backend, tatsächlichen Backend-Selbstcheck und erlaubten
Payload ohne Referenzdaten/Quellen. CPack-Preinstall-Build einmal mit Parallelität4
abgeschlossen, keine zusätzliche Gesamttestsuite ausgeführt. NATIVE_MACROS_RUNTIME.md;
F088/F041/F050/F057/F148 und456/384/72 unverändert. Direkt weiter programmieren.


Eigene Standardmodule sind jetzt echte native Projektobjekte: +Modul/Texteditor,
Ctrl+S ohne Schließen, Wiederöffnen/Bearbeiten, Syntaxinventar ohne Ausführungsclaim,
AllModules-/SysCmd-Metadaten und bestätigte Löschung nur eigener Module. Namen
werden nach Save gesperrt; Umbenennen/Classes/VBE/volle Sprache bleiben offen.
Version1/vba-Definitionen liegen getrennt in _la_objects, importierte VBA-Zeilen
werden nicht geändert. Gemeinsamer Sources-Leser versorgt native Aufrufe/Eval,
öffentliche Variablen und bestehende SQL-Registry, erhält Archiv-Latest-/Shadow-
Prüfungen und Quellenrevision/ROLLBACK-Kohärenz und verweigert Kollisionen statt
Auswahl. Bounds256 Module/16MiB Quellen/1M Katalogzeilen/5s Scan, eigene Quelle1MiB.
Löschen prüft Projektdentität/-generation nach Modalinteraktion und gespeicherte
Revision. Fehlende direkte Includes und gemischte auto-Typdeklaration korrigiert.
DAO1/1 (2,55s), Layout1/1 (11,36s) im Bündel bestanden. Nach tatsächlicher
Qt-Fokussetzung für Ctrl+S und normalem LinAccess-Abfragepfad zur initialen UDF-
Registrierung nur betroffene GUI-Prüfung erneut1/1 (6,23s), App gebaut: echte
Editor-/VBA-/SQL-Aufrufe, gemeinsamer Counter/Private-Helper, SQL-Private-Sperre,
vorbereitete Quellenänderung/Savepoint-Rücknahme, native/importierte Namens- und
TEMP-Schattenfehler, Löschung/alte Prepared-Verweigerung sowie unveränderte
vorhandene Archivbytes. Keine neue CTest-ID/Windows-Beobachtung, kein neuer
Gesamtsuite-/Plattformnachweis. Paketprüfung gehört zum vorherigen64abd69-Stand,
kein frisch paketierter Modulstand behauptet. NATIVE_STANDARD_MODULES_RUNTIME.md;
F042/F050/F068/F118/F130/F155 und456/384/72 unverändert. Direkt weiter.


Application.Run für unqualifizierte öffentliche Standardmodul-Functions/Subs
implementiert: Procedure/Arg1..Arg30 mit bestehender benannter Bindung,
typisierte Rückgabe bzw. Empty, optionale Parameter und geteilter Modulzustand.
Private/fehlende Ziele2517, Projektqualifikation3251, Arity450; Fehlernummer,
Source und Help bleiben erhalten. Eigener Run-Grenzwert32 liefert28 vor der
allgemeinen Interpreter-Stapelgrenze. Ausführungstiefe schützt laufenden Kontext.
Kein DoCmd.Run, keine neue COM-Automation oder ByRef-Aufrufer-Rückschreibung.
App gebaut; bestehendes mainwindow_documents1/1 (6,48s) bestanden. Nach falscher
neuer Arity-Testannahme450 statt449 und konkret gefundenem Zusammenspiel der
Rekursionsgrenzen nur betroffene GUI-Prüfung wiederholt. Keine neue CTest-ID,
Gesamtsuite/Windows-Beobachtung/Paketprüfung behauptet. NATIVE_APPLICATION_RUN.md;
F042/F050/F118/F148 und456/384/72 unverändert. Direkt weiter mit Makro-RunCode.


Native Makroaktion RunCode implementiert, keine DoCmd.RunCode-Methode:
ein einzelner unqualifizierter Function-Aufruf mit Klammern, Rückgabewert ignoriert;
Quellsyntax/äußere Aufrufform vor erster Aktion geprüft. Geschachtelte Argumente,
doppelte Stringquotes, Datums-/Bracket-Literale im Scanner; ASCII-Prozedurnamen,
64Ki UTF16 Ausdruckslimit. Im gebundenen Formular-/Berichtsereignis CodeBehind
vor Standardmodulen, danach restaurierter Ereigniskontext; Standard-Private/
fehlende Funktionen2517, Subs/Namenskonflikte3251, originale Fehlerdaten plus
Makroaktionskontext. GUI/Layout zusammen2/2 (18,17s) ohne Korrekturlauf, App gebaut:
wirkliche Funktionen/Seiteneffekte, verschachtelte Makros, Private-CodeBehind,
Kontextrestauration, verschachtelte/quotierte Argumente, Sub-/Private-Sperren,
struktureller Abbruch ohne vorangestellten SQL-Effekt und Fehler7001/Source.
Aktuelles DEB gebaut; erste Paketprüfung hatte im aktivierten SDK dessen älteres
SQLite im LD_LIBRARY_PATH und scheiterte3251. Nur Paketprüfung mit wie bei den
gezielten Tests vorangestelltem gebündeltem SQLite wiederholt:5-Datei-Payload,
App-/Helper-BuildIDs, MDBTools1.0.1, bytegleiches SQLite3.54.1 samt tatsächlichem
JOIN/NULL-/Counter-Backendcheck bestanden. Keine Wiederholung der GUI-Prüfungen.
Originalupload-Hashes3/3 und Fixtures4/4 unverändert. Keine neue CTest-ID,
Gesamtsuite/Windows-Beobachtung. NATIVE_MACROS_RUNTIME.md; F088/F042/F050/F057/F148
und456/384/72 unverändert. Weiterprogrammieren, kein Gesamtabschluss.


Eigene gespeicherte Standardmodule können im realen Editor mit Prozedurauswahl,
Ausführen/F5 direkt laufen. Neuer NativeStandardModuleHost nutzt tatsächliches
CurrentDb/DBEngine, gemeinsamen Modul-/Staticzustand und öffentliche Modulaufrufe,
bestehende Qt-Vergleichs-/Format-/MsgBox-/InputBox-Pfade, keinen Formularattrappen-
kontext. Me/current-record und fehlende Application-Objekte bleiben abgegrenzt.
Projekt-/Generationsprüfung bei Modulentry/DB/Dialog; gespeicherte Definition
gegen gemeinsamen Katalog geprüft, ungespeicherter Text verweigert. Auswahl ist
transient und markiert Quelle nicht dirty; Edit deaktiviert Run bis Analyze/Save.
Neue gemeinsame MsgBox-Brücke korrigiert zunächst gefundenen Linkagefehler der
anonymen Hilfsfunktion; fehlenden QTextBrowser-Testinclude ergänzt. App gebaut,
bestehendes GUI1/1 (7,20s) bestanden: tatsächliches F5, DAO-Effekt und zweites Modul,
mit Formular geteilter Counter44, Rückgabewert, MsgBox-Ja6, Arity450, Me438,
GoSub-Preflight3251 ohne Präfix-SQL und Unsaved-Sperre. Keine neue CTest-ID oder
Gesamtsuite/Windows-Beobachtung; aktuelles DEB gehört vorherigem d36ca9f-Stand.
NATIVE_STANDARD_MODULES_RUNTIME.md; F042/F043/F050/F060/F118/F148 und456/384/72
unverändert. Direkt weiter mit formularfreiem Application.Run/Eval-Kontext.


Formularfreier Modul-Anwendungskontext erweitert: echte projekt-/generations-
gebundene Application für Run/Eval/DBEngine; Domänenfunktionen lesen tatsächliche
DB. Gemeinsamer Qt-freier Run-Helfer erhält in beiden Hosts Arity, öffentliche
Auflösung, typisierte Rückgaben/Fehlerdaten und32-Level-Grenze. Variant-Argumente
bewahren Objektverweise (DAO.Field); skalare Zielparameter behalten native
Coercion-Grenzen. Schwacher Projektcache hält keine Editor-/Host-Referenz und
prüft Projektgeneration/GUI-Thread; gehaltene Application bleibt nach Editorclose
verwendbar, Generationswechsel3420. Volles Application-Modell/identische Objekt-
identität zu Formularhosts/implizite Objekt-Skalarkoerzierung bleiben offen.
App gebaut; vorhandene GUI/Layout2/2 (18,74s) bestanden: dynamischer zweiter
Modulaufruf mit DAO-Schreiben, Eval/echtes DCount, DAO.Field-Objektargument,
rekursives Run28, gehaltene Application nach Editorclose, Identität über neuen
Host, Fehler7001/Source und Generationssperre. Testzugriff auf private Eval-Methode
vor dem Lauf auf öffentliche Run-Rückgabe eines Access.Application-Funktionstyps
korrigiert. Keine neue CTest-ID/Gesamtsuite/Windows-Beobachtung/Paketprüfung.
NATIVE_APPLICATION_RUN.md; F042/F043/F050/F118/F148 und456/384/72 unverändert.
Direkt weiter mit formularfreier nativer Makroausführung.


Gespeicherte native Makros können über Ansicht/Ausgewähltes Makro ausführen direkt
vom Datenbankfenster laufen. Reeller Modul-/DAO-Anwendungshost mit DoCmd-Proxy,
kein verstecktes Formular. Makro-Budget/Repeat/Bedingungen/Stop/Fehlerkontext in
gemeinsamen Runner extrahiert;16 aktive/10000 besuchte Aktionen gelten auch über
Hostwechsel, VbaExecutionScope hält Stop über den gesamten Aufruf. RunSQL hat
geprüfte Savepoint-/UseTransaction-Pfade; RunCode öffentliche Functions. Tatsächliche
Tabellen/Abfragen über bestehende Dokument-APIs; eigene gespeicherte Normalformulare
und Berichtsvorschauen über bestehende unabhängige showDocument-Lebensdauer.
Aktives reales Formular erhält RunCommand/Requery; sonst tatsächliches Datenblatt.
Projektgeneration/Verbindung/GUI-Thread und Modalgrenzen bleiben geprüft; Menülauf
schützt laufende Projektoperationen. Formular-/Berichtsfilter/weitere Modi und
Archiv-only-Dokumente sind in diesem Host noch nicht implementiert. ReportView0
wird3251 abgewiesen; Vorschau benötigt2. CancelEvent ohne Ereignis3251.
App gebaut, GUI/Layout zusammen2/2 (19,15s) bestanden: tatsächliche Menüstarter-
Kette mit SQL/Function/verschachteltem StopMacro/Bedingungen, Table/Query/Form/
Report/Requery/RunCommand, OnLoad-SQL und Form-RecordCount3, Fehler in Aktion2 mit
persistiertem Vorläufereffekt, Rekursion3251 mit folgendem unabhängigen Erfolg,
SELECT-RunSQL2342, nichtgebundenes CancelEvent und echtes abgebrochenes Form_Open.
Fehlenden QStatusBar-Testinclude vor Lauf korrigiert; keine neue CTest-ID.
Aktuelles DEB mit4 parallelen Preinstall-Jobs gebaut; Payload/Berechtigungen,
App-/Helper-BuildIDs, MDBTools1.0.1 und bytegleiches SQLite3.54.1 samt tatsächlichem
JOIN/NULL-/Counter-Backendcheck bestanden. Keine Gesamtsuite/Windows-Beobachtung.
NATIVE_MACROS_RUNTIME.md; F088/F042/F050/F057/F114/F148 und456/384/72 unverändert.
Direkt weiter, keine Gesamtabschlussmeldung.


Native AutoExec beim tatsächlichen Projektöffnen implementiert: nach Schema/
QueryViews/Explorer-Initialisierung einmal pro neuem Öffnen, CI-Katalogauflösung;
nur gespeicherte native _la_objects-Makros, keine Archiv-only-Access-Makros.
Gehaltener Shift-Zustand überspringt Startmakro. Menü/Startup nutzen gemeinsamen
Projektstarter mit Generations-/Verbindungs-/Modal-/Ausführungsschutz und Refresh
auch nach Teilfehlern; neue Generation leert schwachen Modul-Application-Cache.
AutoExec-Fehler bleiben sichtbar, Projekt offen, bereits abgeschlossene Effekte
werden nicht erfunden zurückgerollt; Shift erlaubt nächsten Start ohne Wiederholung.
AllowBypassKey=False/Startup-Formular/Trust-/Access-Makroimport bleiben offen.
App gebaut; nur bestehende GUI1/1 (7,30s) bestanden. Erste Shift-Prüfung hatte
synthetische Taste ohne Modifier-Zustand; korrigiert auf echte Qt-Shift-KeyPress
mit ShiftModifier im geöffneten Dateidialog, nur betroffene GUI erneut. Reale
Dateiprojekte: SQL+RunCode+Table beim ersten/erneuten Öffnen je einmal, Shift ohne
Effekte, unveränderte Archiv-only-Makrobytes ohne Ausführung, Fehler2517 mit
persistiertem Vorläufer und erfolgreichem Shift-Neustart, gehaltene Application
nach tatsächlichem Projektwechsel3420. Keine neue CTest-ID/Windows-Beobachtung/
Gesamtsuite; Paket gehört vorherigem db892cc-Stand. NATIVE_AUTOEXEC_RUNTIME.md;
F088/F042/F050/F121/F148 und456/384/72 unverändert. Direkt weiter.


RunMacro RepeatExpression implementiert, in API und nativem Aktionseditor:
Bedingung vor jedem Versuch; Anzahl allein begrenzt, Bedingung allein bisFalse,
beide bis erster Grenze, beide ausgelassen einmal. Bool-/Zahlkonstanten über
native String-/Expression-Konvertierung, Textsyntax bis4096 UTF16 vor Effekten.
Ausgelassene Count-Position im nativen JSON ausschließlich für RunMacro als null
(Missing); Editor-Save/Reopen erhält leere Anzahl und Argument3. Ohne Count max
1000 abgeschlossene Wiederholungen; weiterer wahrer Versuch3251, abschließendes
False normal. Gemeinsame16-/10000-Aktionsbudgets und StopMacro unverändert.
Layout1/1 (11,44s) bestand im Bündel, GUI fand verlorene Fehlerquelle im alten
VbaHost::eval-Pfad. Erklärung/Source/HelpFile/HelpContext dort korrekt weitergereicht,
gezielt bestehende Kern-/GUI2/2 (9,32s) bestanden, App gebaut: reale DAO-Predicate-
Effekte3/Prüfungen4, Count+Expr, False0/TrueCount, Fehler7002 samt Wiederholungs-
kontext und separater Kern-Source/Help42-Nachweis, ungültige spätere Expr ohne
Präfixeffekt, nichtendender verschachtelter Noop an1000er-Grenze, danach UI-
Erstellung/CtrlS/Wiederöffnen und erneuter erfolgreicher Ablauf. Alte explizite
False-RepeatExpression-Verweigerung auf tatsächlichen erfolgreichen Null-Durchlauf
umgestellt. Keine neue CTest-ID/Gesamtsuite/Windows-Beobachtung/Paketprüfung;
aktuelles DEB gehört db892cc. NATIVE_MACROS_RUNTIME.md; F088/F041/F051/F118/F148
und456/384/72 unverändert. Direkt weiter mit nativen Makrogruppen.


Native Makrogruppen implementiert: eigenes Formatversion2 mit geordneten groups,
1..64 eindeutig CI benannten Teilmakros, insgesamt max256 Aktionen/1MiB; Version1
bleibt kompatibel. Gruppe startet erstes Teilmakro, Gruppe.Teilmakro gezielt;
Namens-/Callvalidierung und Formular-/Feldereignisdesigner akzeptieren diese Form.
Alle Teilmakros strukturell geprüft, bevor ausgewählter Teil Effekte ausführt;
fehlerhafte inaktive Gruppe blockiert, doppelte Unicode-CI-Namen3251, fehlender
Teil2485. Bestehender Runner hält Wiederholung/Stop/übergreifende Budgets.
Realer Editor ergänzt Gruppenmodus, Name/Auswahl/Hinzufügen/Entfernen/Umordnen,
behält rohe auch ungültige Zell-Drafts beim Wechsel; Auswahl/Render markieren
Quelle nicht dirty. Mehrteilige Gruppe wird nicht still in Einzelmakro verwandelt.
Projektgeneration vor Save geprüft, aktive Zellbearbeitung vor CtrlS-Serialisierung
committed; neue Draft-Serialisierung bewahrt RepeatCount-Missing-Hole. Keine
automatische Verweisrefaktorierung oder Access-SaveAsText-/Binär-Makroimport.
App gebaut; bestehende GUI/Layout2/2 (20,10s) ohne Korrekturlauf bestanden: echte
Gruppenerstellung/CtrlS/Wechsel/Wiederöffnen/Umordnung mit geänderter Default-
Ausführung, Unicode-Teilaufruf, verschachtelter Parent/StopMacro/Repeat, reales
Form_Load, ungültige inaktive Gruppe und doppelte Labels ohne Effekte. Aktuelles
DEB gebaut/geprüft:5-Datei-Payload, Rechte, App-/Helper-BuildIDs, MDBTools1.0.1,
bytegleiches SQLite3.54.1 mit tatsächlichem JOIN/NULL-/Counter-Backendcheck.
Keine neue CTest-ID/Gesamtsuite/Windows-Beobachtung. NATIVE_MACROS_RUNTIME.md;
F088/F050/F057/F118/F148/F155 und456/384/72 unverändert. Direkt weiter mit
Speicher-/Löschschutz für eigene Makroobjekte.


Eigene Makroobjekte erhalten transaktionalen Speicher-/Löschschutz: neue Namen
kollidieren CI mit vorhandenen nativen/importierten Makronamen3010 statt Replace;
gespeicherte Objektnamen sind im Editor readonly. Save und Delete vergleichen
exakte alte Definition/Name unter Savepoint, staleRevision3251 statt Lost Update.
Main-Katalog physisch/ohne TEMP-Schatten geprüft, zusätzliche main/temp Trigger
verweigert; Größenlimit1MiB vor Materialisierung. Lesen/Ändern importierter
Quellarchive bleibt ausgeschlossen. Eigene Makros löschbar mit Bestätigung im
Editor/Explorer; auch unlesbare eigene Definitionen über Explorer entfernbar.
Projektgeneration nach Modalfrage und Definition nach Antwort erneut geprüft.
DoCmd.DeleteObject acMacro4 für eigene gespeicherte Objekte über bestehende
Projekt-API ergänzt. Gemeinsame laufende MacroFrames sperren Ersetzen/Löschen
auch für Gruppen und über Formular-/Modulhostgrenzen; kein globaler Jet-Lockclaim.
Layout1/1 (11,74s) im Bündel bestanden. GUI fand alte vorgelagerte Case-only-
Namensprüfung; entfernt zugunsten zentralem3010-Pfad. Nur betroffene GUI erneut
1/1 (7,41s), App gebaut: echte zweite Verbindung und Revision während bestätigter
Löschfrage, Native-/Archiv-Kollision, No/Yes, TEMP-/Trigger-Verweigerung, Löschen
ungültiger eigener Definition, aktive3251-Sperre in beiden Hosts mit folgenden
SQL-Effekten, anschließendes DeleteObject und unveränderte Archivbytes. Original-
Uploads3/3 sowie Fixtures4/4 erneut hashgleich. Keine neue CTest-ID/Gesamtsuite/
Windows-Beobachtung/Paketprüfung; DEB gehört9c80f00. NATIVE_MACRO_STORAGE.md;
F088/F050/F121/F148/F155 und456/384/72 unverändert. Direkt weiter mit allgemeinen
Makro-Projektbefehlen und getrenntem Projekt-Ausführungsschutz.


Native Makroaktionen Close/DeleteObject/Rename ergänzt über reale bestehende
Projekt-APIs: eigene geschlossene Tabellen/Abfragen umbenennen/löschen, eigene
Makroobjekte löschen, offene Datenblätter sowie reale Form-/Berichtsdokumente
schließen. Keine Save=acSaveNo-Verwechslung mit Datensatz-Undo; bestehende
validierte Close-Pfade committen Recorddrafts und geben SQL-Modelle frei.
Macro-Projektstarter erhält eigene Ausführungstiefe statt SQL-Tiefe; legitime
Objektbefehle bleiben nutzbar. Projekt-/Anwendungsschließen und neue UI-Makro-
Reentry bleiben bei aktiven Frames blockiert, auch im direkten Formularhost.
Gemeinsamer Preview-Dispatch hält NativeExecutionDepth bis Befehlsende, damit
Close/Requery/RunCommand keine laufenden Dokumentaufrufe vorzeitig entsorgen.
Wiederverwendete gespeicherte Fenster wiederholen keine historischen Initfehler;
neue fehlgeschlagene Initialisierungen bleiben echte Fehler mit neuerDoc-Bereinigung.
App gebaut; bestehende GUI/Layout2/2 (18,92s) ohne Korrekturlauf bestanden:
Menü-Makrokette mit realem Table/Query-Rename, Open/Close/Drop und Makrolöschung,
aktuellem Query-DefaultClose, Pending-Draft43 trotz acSaveNo, echtem Form_Unload/
Report_Close und anschließendem SQL. Während MsgBox ist tatsächliches Window.close
abgewiesen bei gleicher Projektgeneration/offener DB; danach Weiterlauf. Bestehende
Form mit altem Load7003 nur aktiviert, neue Init7003 abgewiesen/bereinigt. Grenzen
bleiben explizite Objekttypen/Modi, native Ownership, aktive Model-/Makro-Leases,
keine automatische Verweisrefaktorierung oder vollständige Access-Close-Parität.
Keine neue CTest-ID/Gesamtsuite/Windows-Beobachtung/Paketprüfung; DEB gehört9c80f00.
NATIVE_MACRO_PROJECT_COMMANDS.md; F088/F050/F057/F114/F121/F135/F140/F148/F155
und456/384/72 unverändert. Kontinuierlich weiter.


SelectObject als native Makroaktion und formularfreier DoCmd-Member ergänzt:
wirklich offene Tabellen/Abfragen über bestehende Aktivierung, echte Form-/Report-
Instanzen über lebensdauergeschützten Preview-Dispatch, eigene Makro-/Modulobjekte
und Ordner im Datenbankfenster. Gleiche4/5-InDatabase-Erweiterung im Formularhost;
kein erfundenes Modulfenster, kein Öffnen eines geschlossenen Datenobjekts.
Explorer behält tatsächliche logische Auswahl/Archiv-ID bei Refresh derselben
Projektgeneration; neues Projekt übernimmt keine alte Auswahl. Damit bleibt die
von einem Makro gesetzte Auswahl nach dessen abschließendem Refresh erhalten.
Gewachsene NativeStandardModuleHost/DoCmd/Application-Implementierung aus Header
in eigene Forms-C++-Einheit verschoben (öffentliche schlanke Deklaration); zwei
mechanische Qualifizierungs-/Duplikatfehler vor Testlauf korrigiert. Keine Änderung
an Kern-Qt-Freiheit oder Python/Java-App-Abhängigkeiten. App gebaut, bestehende
GUI/Layout2/2 (19,03s) bestanden: reale Table/Query/Form/Report-Aktivierung,
Macro/Module-/Ordnerauswahl, Refresh und tatsächlicher Menü-Makro-Endzustand,
Formularhost4/5, fehlende/geschlossene/unsupported Ziele. Vorige Modul-/Makro-/
Speicher-/Startprüfungen im gleichen Ziel weiter enthalten. Aktuelles DEB gebaut
und Payload/Rechte/App-/Helper-BuildIDs/MDBTools1.0.1/bytegleiches SQLite3.54.1
samt realem JOIN/NULL-/Countercheck verifiziert. Keine neue CTest-ID/Gesamtsuite/
Windows-Beobachtung. NATIVE_MACRO_PROJECT_COMMANDS.md; F088/F050/F135/F140/F148
und456/384/72 unverändert. Direkt weiter mit StopAllMacros/Echo-Wiederherstellung.


StopAllMacros als echte parameterlose Makroaktion implementiert, kein erfundener
DoCmd-Member. Gemeinsamer Stop-Zustand beendet den aktuellen verschachtelten
Aufrufbaum, Wiederholungen und bedingungsgesteuerte Fortsetzungen; nach äußerstem
Unwind sind Zustand/Budgets/Frames wieder frei. Bereits laufender nativer VBA-
Functioncode wird nicht global abgebrochen; folgende Makroaktionen bleiben aber
beendet. Neue unabhängige Rootaufrufe funktionieren. Echo-Aktion und Echo-/
Application.Echo im formularfreien Host nutzen bestehende reale UI-Suppression;
StopAllMacros stellt implementiertes Echo wieder her. Vollständiger SetWarnings-
Zustand/Access-Stop-/Reentry-Scheduling bleiben offen, keine Rücksetzung eines
nicht implementierten Warnungszustands behauptet. App gebaut, vorhandene GUI/
Layout2/2 (19,55s) ohne Korrekturlauf bestanden: tatsächliche MsgBox bei ausgeschal-
teten Updates, verschachtelter SQL/RunCode/Stop beider Hosts, native VBA-Fortsetzung,
fehlende nachfolgende Aktionen/Wiederholungen, unabhängiger Folgestart, Stop aus
RepeatExpression, False-Stop ohne vorzeitige Echo-Restauration, benanntes
Application.Echo, ungültiger Status ohne UI-Mutation, kein DoCmd.StopAllMacros.
Keine neue CTest-ID/Gesamtsuite/Windows-Beobachtung/Paketprüfung; DEB gehört cb8931b.
NATIVE_MACROS_RUNTIME.md; F088/F050/F051/F059/F135/F148 und456/384/72 unverändert.
Direkt weiter mit vollständigem Aktionsargument-Editor/weiterer Aktionsausführung.


## Fortschreibung 2026-11-09: vollständige native Aktionsargumente / MessageBox
Argumente… bearbeitet sämtliche deklarierten Aktionsargumente; zusätzliche Werte
bleiben bei Teilmakrowechseln und Zeilenbewegungen erhalten, Abbrechen verändert
keinen Entwurf. MessageBox ist eine echte vierargumentige Makroaktion in beiden
Hosts: kontextgebundene =Auswertung, Signalton, benannter Typ, Titel, escaped
@-Abschnitte; kein erfundenes DoCmd.MessageBox. Native JSON-Typnamen sind kein
Access-Binärimport. Quellen-/Ergebnisgrenzen und bekannte Typen werden geprüft.
App gebaut; vorhandene GUI/Layout2/2 bestanden (19,30s), einschließlich echtem
Argumentdialog, vierter Titel, Abbruch, Teilmakrowechsel/Speichern, realen Dialogen
in beiden Hosts und unbekanntem Typ. Build-Include für vollständigen VbaValue
korrigiert, kein Testfehler/Korrekturlauf. Originaluploads3/3 und Fixturehashes4/4
unverändert. Keine neue CTest-ID/Gesamtsuite/Windows-Beobachtung/Paketprüfung.
NATIVE_MACROS_RUNTIME.md; F088/F050/F051/F059/F135/F148,456/384/72 unverändert.
Weiter: GoToRecord aus tatsächlichen Standardmodulen/Makros mit vier Argumenten
an vorhandene Daten-/Formularnavigation anbinden; mehrzeilige Argumente erhalten.


## Fortschreibung 2026-11-09: GoToRecord aus Modul/Makro, mehrzeilige Argumente
Standardmodule Application.DoCmd.GoToRecord und native Makroaktion unterstützen
vier optionale/namentliche Argumente, tatsächliche offene Tabellen/Abfragen und
sichtbare Projektformulare. Bestehende Save-/Ereignis-/Reihenfolge-/Grenzlogik wird
wiederverwendet; keine erzeugten Ersatzdokumente. Native JSON-null ist hier eine
Auslassung, direktes VBA-Null bleibt ein Laufzeitwert. Detailargumente bearbeiten
Text mehrzeilig ohne SQL-Zeilenverlust; Bedingung bleibt separat.
App gebaut; Layout1/1 (11,40s), nach Korrektur nur des neuen Prüffall-Abfragemodus
auf Nur-Lesen GUI1/1 (7,61s) bestanden. Geprüft: tatsächliche Cursorpositionen,
benannter VBA-Modulaufruf, vierargumentiges Makro beider Hosts, optionale Löcher,
Tabellen-/Abfrage-/Formularwechsel, Offsetfehler ohne Cursoränderung, geschlossene
Form verweigert, mehrzeiliger Argumentdialog. Erste GUI verweigerte korrekt die
nicht unterstützte bearbeitbare sortierte Projektionsabfrage; keine Lockerung.
Keine neue CTest-ID/Gesamtsuite/Windows-Beobachtung/Paketprüfung. Modale Designer
und allgemeine versteckte Formularsuche bleiben explizite Grenzen.
NATIVE_MACROS_RUNTIME.md; F178/F088/F052 und456/384/72 unverändert.
Direkt weiter: vollständige FindRecord-Anbindung aus Standardmodulen/Makros.


## Fortschreibung 2026-11-09: FindRecord aus Standardmodul/Makro
Application.DoCmd.FindRecord unterstützt sieben/namentliche Argumente und echte
aktive Projektformulare, Tabellen und Abfragen. Gemeinsame Argumentprüfung mit
bestehendem Formularhost; typisierte Werte/Platzhalter/Formatgrenzen bleiben
unverändert. Native Makroaktion hat literaltextiges FindWhat und optionale Löcher;
keine Behauptung beliebiger Access-Makroausdrücke/Binärdekodierung.
Tatsächliche Tabellenöffnung wählt nun eine erste Zelle, damit FindRecord ohne
vorherigen manuellen Klick ein Suchfeld hat. Datensatzmarkierung selektiert die
aktuelle ganze Zeile ausdrücklich unabhängig vom initialen Qt-Auswahlanker und
bewahrt die aktuelle Spalte. Vorhandene GUI deckte diesen Folgefehler auf.
App gebaut; Layout1/1 (11,23s), nach gezielten Korrekturen GUI1/1 (7,91s) bestanden.
Der neue Suchprüffall setzt seine Sortierung ausdrücklich: die vorherige Annahme
aufsteigender Standardsortierung war falsch, der Zahlen-Treffer bereits korrekt.
Geprüft: sieben tatsächliche Editorargumente/Speicherung, Zahlen/Currency,
Platzhalter/Weitersuche, benanntes Modul, Nur-Lesen-Abfrage, Formular/Makro beider
Hosts, Null-/Mustervalidierung ohne Cursoränderung und vorhandene Zeilenauswahl.
Keine neue CTest-ID/Gesamtsuite/Windows-Beobachtung/Paketprüfung; vollständige
No-Match-/Zeit-/Untertyp-Parität bleibt offen. NATIVE_MACROS_RUNTIME.md;
F088/F052/F178/F363 und456/384/72 unverändert. Direkt weiter mit tatsächlichen
Fokus-/Cursorbefehlen und gemeinsamer UI-Ausführungslebensdauer im Modulhost.


## Fortschreibung 2026-11-09: Fokus/Cursor und gemeinsame GUI-Ausführungsrahmen
GoToControl und Hourglass stehen nativen Makros und dem formularfreien DoCmd-Host
mit namentlichen Argumenten zur Verfügung. Fokus geht an echte Controls eines
aktiven Projektformulars; ohne Formular2046, kein erfundener Me-Kontext.
Modul-Eintritt/-Austritt und beide Makrohosts verwenden denselben nativen Qt-
Ausführungszähler. Wartezeiger bleibt über verschachtelte Funktionen/RunCode bis
zu späteren Aktionen erhalten und wird bei äußerstem normalem/fehlerhaftem Ende
zurückgesetzt; fremde Cursor bleiben geschützt. Echo bleibt separat persistent.
Vorhandene GUI/Layout2/2 (20,12s) bestanden, kein Korrekturlauf: benannter Fokus mit
Textauswahl, tatsächlicher Cursor in verschachtelten und nachfolgenden Dialogen
beider Makrohosts, StopMacro-Ende, Fehler7004, direkte Steuerung/fremder CrossCursor,
fehlender aktiver Formularkontext. App nach finalem Formularhost-Nachlauf über
CPack neu gelinkt. Frisches DEB native-module-ui/package/linaccess-0.7.0-Linux.deb:
Payload/Rechte/Build-ID/bytegleicheSQLite und tatsächlicher Backendcheck mit
MDBTools1.0.1/SQLite3.54.1 JOIN/NULL/Zähler bestanden. Paket ohne MDB/VBA/Fixtures.
Originaluploads3/3 und Fixturehashes4/4 unverändert. Keine neue CTest-ID/Gesamtsuite/
Windows-Beobachtung. NATIVE_MACROS_RUNTIME.md/NATIVE_HOURGLASS.md;
F088/F052/F068/F135/F148 und456/384/72 unverändert. Direkt weiter mit gemeinsamem
CurrentProject/CurrentData aus echten Standardmodulen und Projektlebensdauer.


## Fortschreibung 2026-11-09: gemeinsame CurrentProject-/CurrentData-Objekte
Standardmodule können Application.CurrentProject/CurrentData und die unqualifi-
zierten VBA-Bezeichner verwenden. Formular- und Modulhost nutzen dieselbe echte
Projektkatalogimplementierung und Objektidentität am tatsächlichen Projektfenster;
kein kopierter Katalog/kein erfundenes Formular. Cache zusätzlich generations-
gebunden. Vorhandene Datei-/Readonly-/Deskriptor-/IsLoaded-Grenzen bleiben bestehen.
App gebaut; vorhandene GUI/Layout2/2 (19,49s) ohne Korrekturlauf bestanden:
wirklicher VBA-Katalogzugriff/Counts, gemeinsame Identität in beiden Hosts,
Fortbestand nach Formularschluss, live Modul-Löschung, Argument/Readonly-Fehler,
3420 für gehaltene Projekt-/Datenkatalogobjekte nach tatsächlichem Projektwechsel
im vorhandenen AutoExec-Ablauf. Keine neue CTest-ID/Gesamtsuite/Windows-Beobachtung/
Paketprüfung; zuletzt geprüftes DEB gehört69ac1b9. NATIVE_STANDARD_MODULES_RUNTIME.md;
F052/F042/F059/F135 und456/384/72 unverändert. Direkt weiter: echte Forms-/Reports-
Sammlungen und Screen aus dem formularfreien Host, ohne Platzhalterformular.


## Fortschreibung 2026-11-09: tatsächliche Forms/Reports/Screen aus Standardmodulen
Formularfreier Host unterstützt Forms/Reports und Application.Forms/Reports mit
realen Dokumentobjekten, Count/Name/Index/ForEach, einschließlich nach Öffnen
versteckter Formulare. Application.Screen/bare Screen liefern tatsächliches
ActiveForm/ActiveControl/MousePointer; vorhandene Cursor-/Fokusgrenzen bleiben.
Gemeinsame Implementierung statt Platzhalterformular. Neue Projektobjekte sind
via echtem CurrentProject-Lebenszyklus geschützt; alte formulargebundene Handles
behalten ihren bestehenden Caller-Lebenszyklus. Screen-Suche zusätzlich an DB-
Identität gebunden und nach skalarer Wertauswertung erneut geprüft.
App gebaut; vorhandene GUI/Layout2/2 (20,16s) ohne Korrekturlauf bestanden:
echte sichtbare/versteckte Formulare und Bericht, index-/namensbasierte identische
Dokumentreferenzen, native VBA-ForEach über beide Sammlungen, reale aktive Form/
Control, VBA-MousePointer mit Rücksetzung, live Schließen/alte Dokumentreferenz,
3420 für gehaltene Sammlung/Screen nach tatsächlichem Projektwechsel.
Keine neue CTest-ID/Gesamtsuite/Windows-Beobachtung/Paketprüfung; DEB69ac1b9.
NATIVE_STANDARD_MODULES_RUNTIME.md; F052/F042/F059/F135 und456/384/72 unverändert.
Direkt weiter: globale projektgebundene Sammlungsidentität und tatsächlich
geöffnete/versteckte Formulare als benannte Navigations-/Schließziele.


## Fortschreibung 2026-11-09: globale Projekt-Sammlungsidentität/versteckte Ziele
Native Forms/Reports/Screen teilen einen schwachen, verbindungs-/datei-/generations-
gebundenen Cache am echten Projektfenster, auch über kurzlebige Application-
Wrapper hinweg. Projektwechsel leert den Cache; gehaltene Altobjekte bleiben
3420-geschützt. Tatsächlich geöffnete lebende Dokumente inkl. versteckter/modaler
Instanzen werden nach nativer Öffnungsreihenfolge erfasst; Schließen/Abbruch wird
berücksichtigt. Benanntes GoToRecord und Close erreichen nun echte versteckte
geöffnete Formulare, ohne sie zur Navigation sichtbar zu machen.
App gebaut; vorhandene Layout1/1 und korrigierte GUI1/1 (7,96s) bestanden. Der neue
Prüffall wurde berichtigt: Konstruktor dispatcht bereits Open/Load vor show();
eine erfolgreich geladene noch nicht angezeigte Instanz zählt deshalb richtig
als geöffnet. Kein Produktionsfehler verdeckt. Geprüft: direkte native VBA-Is-
Identität ohne gehaltenen Application-Wrapper, reale unangezeigte/ausgeblendete
Instanzen, versteckte Navigation ohne Anzeigen, benanntes Close und ungültig
gewordene Dokumentreferenz. Keine neue CTest-ID/Gesamtsuite/Windows-Beobachtung/
Paketprüfung; DEB69ac1b9. NATIVE_STANDARD_MODULES_RUNTIME.md;
F052/F178/F042/F135 und456/384/72 unverändert. Direkt weiter mit echtem FileDialog
für Standardmodule und zugehöriger Projekt-/Dialogzustandslebensdauer.


## Fortschreibung 2026-11-09: echte FileDialog-Picker aus Standardmodulen
Application.FileDialog(DialogType) verwendet den vorhandenen tatsächlichen Datei-/
Ordnerpicker (3/4), Office.FileDialog-Typ, Filter/Show/SelectedItems. Keine
Ersatzform, kein injiziertes Ergebnis, Auswahl öffnet/führt keine Datei aus.
Projektfenster hält den Pickerzustand; tatsächlicher Projektwechsel leert den
Cache. Alte Picker bleiben3420-geschützt, neuer Projektaufruf erzeugt funktionie-
renden neuen Picker. Application besitzt jetzt zusätzlich eine eigene Datei-/
Generationslease, unabhängig von einer optionalen Prüfung des Aufrufers.
App gebaut; bestehende Layout1/1 (11,31s) und abschließende GUI1/1 (8,15s) bestanden:
reale temporäre Datei/Ordner, Abbruch/geleerte Auswahl, Filter, typisierter
namentlicher VBA-Aufruf, Open1 verweigert/Null94 sowie tatsächlicher Projektwechsel
mit frisch erzeugtem Application-/Pickerobjekt bei ausdrücklich leerer Zusatz-
prüfung des Hosts. Finale App/GUI auf demselben Quellstand neu gebaut.
Frisches DEB native-module-file-dialog/package/linaccess-0.7.0-Linux.deb:
Payload/Rechte/Build-ID/bytegleicheSQLite und tatsächliches MDBTools1.0.1/
SQLite3.54.1-Backend (JOIN/NULL/Zähler) bestanden; keine MDB/VBA/Fixtures enthalten.
Keine neue CTest-ID/Gesamtsuite/Windows-Beobachtung. NATIVE_FILE_DIALOG.md;
F052/F042/F059/F135 und456/384/72 unverändert. Direkt weiter mit SysCmd am echten
Projekt-Statusbalken/GetObjectState und weiteren vorhandenen Application-Diensten.


## Fortschreibung 2026-11-09: wirkliche Application-Dienste/gemeinsames SysCmd
Standardmodule unterstützen Application.CurrentDb (vorhandener echter DAO-Adapter)
und Application-Domänenfunktionen mit Expr/Domain/Criteria über die bestehenden
Implementierungen. SysCmd1..5 verwendet realen Projektstatus/Meter,10 den echten
CurrentProject-Objektzustand; bare und namentliche Application-Aufrufe unterstützt.
Beide Hosts teilen Argumentprüfung/einmalige Skalargetter-Auswertung mit erneuter
Kontextprüfung. Native Statusoberfläche erkennt jetzt Hauptfenster selbst und
verwendet dessen echten Statusbalken; korrigiert auch Echo-Statustextposition.
App gebaut; bestehende GUI/Layout2/2 (19,72s) ohne Korrekturlauf bestanden: echte
VBA-DAO-Lese-/Domänenaufrufe, benannte Parameter, Status/Meter/Entfernen, tatsäch-
licher Tabellen-Ladezustand, einmaliger Standardwertgetter beider Hosts und Echo
am Projektbalken. Keine neue CTest-ID/Gesamtsuite/Windows-Beobachtung/Paketprüfung;
DEBf4e9ff8. NATIVE_STANDARD_MODULES_RUNTIME.md; F052/F053/F135 und456/384/72 bleiben
unverändert. Direkt weiter mit vollständiger tatsächlicher OpenForm-/OpenReport-
Anbindung (Where/DataMode/OpenArgs/WindowMode) über gemeinsame Dokumentöffnung.


## Fortschreibung 2026-11-09: gemeinsame tatsächliche Dokumentöffnung
Standardmodul-DoCmd.OpenForm hat sieben, OpenReport sechs Argumente und nutzt nun
dieselbe tatsächliche Öffnung wie der Formularhost: gespeicherte/importierte
Objekte, WhereCondition, frische Datenmodi, OpenArgs vor Load, native Fenster-
modi/acDialog-Warten, unabhängige Projektzuordnung und verzögerte Freigabe. Keine
Platzhalterform. Importquellen main-qualifiziert; explizite Where-Quellfehler
behalten Vorrang und Initialisierungsfehler jetzt Source-/Help-Metadaten.
Normalformular0/Berichtsvorschau2 bleiben unterstützt; andere Ansichten/benannte
Filter und unerlaubte Änderung bestehender Instanzen weiterhin3251. Native Makro-
aktionen haben bis zum nächsten Editorumfang weiterhin nur Name/Ansicht.
App gebaut; vorhandene GUI/Layout/Original-Kassa3/3 (137,86s;9,24/12,39/116,22s)
bestanden. Neuer Prüffall-Cast korrigiert, importiertes Testlayout mit tatsächlicher
OnLoad-Bindung ergänzt und alte Report-Window-Verweigerung durch echtes Aus-/
Einblenden ersetzt. Geprüft: Filter/Readonly/OpenArgs/Load, Wiederverwendung ohne
Mutation bei verweigertem Modus, echtes Add-Insert, modales Warten, versteckter
Bericht, Fehler7006 inkl.Source/Help und tatsächlicher importierter Handler.
Original-F129/3075 bleibt bekannte Grenze. Keine neue CTest-ID/Gesamtsuite/Windows-
Beobachtung/Paketprüfung; DEBf4e9ff8. NATIVE_STANDARD_MODULES_RUNTIME.md;
F052/F178/F042/F135 und456/384/72 unverändert. Direkt weiter: optionale Makro-
argumente einschließlich ausdrücklich leerem OpenArgs verlustfrei bearbeiten und
vollständige OpenForm-/OpenReport-Makroargumente ausführen.


## Fortschreibung 2026-11-09: verlustfreie optionale Makroargumente
Argumente… hat Übergeben-Schalter und bewahrt ausdrücklich leeren Text getrennt
vom Standardwert. Präsenz über Laden/Teilmakrowechsel/Aktionsreihenfolge/Speichern
und Zusatzargumente erhalten; Übersichtsänderungen inferieren neu, Pflichtwerte
bleiben erforderlich. Vollständige native OpenForm7-/OpenReport6-Makroargumente.
Optionales JSON-null bezeichnet einheitlich Missing, kein Datenbank-/VBA-Null;
Pflicht-null bleibt ungültig. Alte Formularbefehle nutzen ihre bewährten Defaults,
MessageBox verarbeitet ausgelassene Werte explizit. Kein erfundenes DoCmd-API.
Vorhandene GUI/Layout2/2 (20,59s) ohne Korrekturlauf bestanden: echte Editor-
Schalter/Move/Gruppe/Save/Reopen, leeres OpenArgs bleibt String und wird in beiden
Hosts tatsächlich von Null unterschieden; reale alte Befehle mit Default-Löchern
und MessageBox-Icon-Default. Anschließend nur irreführende optionale Übersichts-
spaltenbeschriftungen entfernt, keine Logikänderung. Frische App/DEB native-macro-
presence/package/linaccess-0.7.0-Linux.deb mit Payload/Rechten/Build-ID/bytegleicher
SQLite und echtem MDBTools1.0.1/SQLite3.54.1-Backendcheck (JOIN/NULL/Zähler) geprüft.
Uploads3/3 und Fixtures4/4 hashgleich. Keine neue CTest-ID/Gesamtsuite/Windows-
Beobachtung. Originaltest zuletzt f86d8eb:3/3 inkl.Original116,22s; F129/3075 offen.
NATIVE_MACROS_RUNTIME.md; F088/F052/F178/F135 und456/384/72 unverändert.
Direkt weiter mit echter SetValue-Makrozuweisung über vorhandene native Formular-
/Steuerelementobjekte und bestehende Ausdrucksauswertung; dokumentierte Grenzen.


## Fortschreibung 2026-11-09: tatsächliche SetValue-Makrozuweisung
SetValue(Item,Expression) greift auf echte native Controls/ausdrückliche Formular-/
Berichts-/Control-Eigenschaften und vorhandene Eval-/Setterwege zu, keine erzeugte
VBA-Zuweisung/keine Ersatzform/kein erfundenes DoCmd.SetValue. Bare Controls brauchen
echten Aufrufer/aktives Dokument, explizite Forms-Ziele funktionieren ohne Me.
Syntax/4096-Grenzen vorab, dokumentierte führende=-Doppelauswertung ausdrücklich
3251, Ziel-/Kontextprüfung vor/nach Auswertung. Vorhandene Readonly-/Puffer-/Save-
und Validierungsregeln gelten; Feld-ohne-Control/Funktionsziel-/Vollreport-/Ereignis-
parität bleibt offen. App gebaut; Layout1/1 (11,34s), abschließende GUI1/1 (8,49s)
bestanden. Gezielter Prüffall nach Einblenden wählt tatsächliche Form ausdrücklich;
vorher438 war mangels aktivem Dokument korrekt. Reale Zuweisung/Speicherung beider
Hosts, explizites verstecktes Ziel/Visible, Null, Fehler11/2448 ohne Wertänderung,
Schließen während RHS2450, ungültiges Rohmakro vor SQL-Präfix, DoCmd.SetValue438.
Keine neue CTest-ID/Gesamtsuite/Windows-Beobachtung/Paketprüfung; DEB22a9284.
NATIVE_MACROS_RUNTIME.md; F088/F052/F030/F135 und456/384/72 unverändert.
Direkt weiter mit tatsächlichem RepaintObject für geöffnete native Dokumente.


## Fortschreibung 2026-11-09: tatsächliches RepaintObject
RepaintObject(ObjectType,ObjectName) in Standardmodulen/Formularen und nativen
Makros zeichnet echte geöffnete Tabellen/Abfragen/Modul-/Makroentwürfe sowie
Formulare/Berichte neu. Form.Repaint verwendet vorhandene berechnete Controls,
Berichte rendern ihren vorhandenen Datensatzbestand erneut; keine Requery oder
implizite Speicherung. Projekt-/Generation-/Lebensdauer- und Kontextprüfungen
bleiben wirksam. Tabellen-Cache, laufende Zelleneingabe und Entwurfsänderungen
bleiben erhalten. GUI1/1 (9,43s), Layout1/1 (11,34s) bestanden; App gebaut.
Der letzte gezielte Korrekturlauf betraf ausschließlich den Prüffall: Speichern
per Schaltfläche schließt den Editor bestimmungsgemäß; Strg+S hält ihn offen.
GUI-Erweiterung in kleinere Quelldatei desselben bestehenden CTest-Ziels getrennt,
keine neue CTest-ID/Gesamtsuite/Windows-Beobachtung/Paketprüfung; DEB22a9284.
Vollständige Access-Neuzeichnungs-/Ereignis-/Berichtparität bleibt offen;
F088/F052/F050/F135 und456/384/72 unverändert.


## Fortschreibung 2026-11-09: SetValue-Ausdruck mit führendem Gleichheitszeichen
Expression mit führendem= wertet erst den Rest einmal aus und danach dessen
Skalarwert als Ausdruckstext einmal. Reale Referenzen/Funktionen beider Hosts
verwenden vorhandene Eval-/Setterwege; keine bloße Entfernung des=, keine
unbegrenzte Rekursion. Null94/Errornummer, Syntax-/4096-Textgrenzen, erneutes=3251
und Kontext-/Lebensdauerprüfungen vor Zuweisung. Dynamisches Item bleibt offen.
App gebaut, vorhandene GUI/Layout2/2 (20,35s) ohne Korrekturlauf bestanden:
Zahl, referenzierender Text, tatsächlicher öffentlicher Funktionsaufruf exakt
einmal je Host (SQL-Audit), sechs Fehlerfälle ohne Wert-/Dirty-Änderung und reale
Speicherung. Ungültiger Rohmakro-Präfixtest verwendet nun tatsächlich ungültige
Syntax=1+; vollständiger Vorabcheck vor SQL bleibt belegt. Keine neue CTest-ID/
Gesamtsuite/Windows-Beobachtung/Paketprüfung; bisheriges DEB22a9284.
NATIVE_MACROS_RUNTIME.md; F088/F052/F118/F164 und456/384/72 unverändert.
Direkt weiter mit echten Formularfeldern ohne gebundenes Steuerelement.


## Fortschreibung 2026-11-09: echte Formularfelder ohne Steuerelement
Native FormField-Referenzen lesen/ändern tatsächliche RecordSource-Spalten ohne
Ersatz-Controls; Name/Value/Parent und Lebensdauer geprüft, Controls haben Vorrang,
bereits gebundene Feldnamen verwenden tatsächliches Control. Eigener typisierter
Datensatzpuffer ist in Save/Undo/ControlUndo/Neuanlage/Navigation und berechnete
Controls integriert. Parameterisierte gemeinsame Speicherung mit anderen Controls,
CHECK-Fehler behalten Entwurf, AutoWert/Readonly/geschlossene Form verweigern.
VBA und SetValue beider Hosts verwenden denselben Weg; keine DAO-/Berichts-/
Join-/Vollereignisparität behauptet. FORM_RECORD_FIELDS_RUNTIME.md.
App gebaut; abschließende GUI1/1 (8,50s), Layout1/1 (11,32s), Original1/1 (124,97s)
bestanden. Korrekturlauf änderte ausschließlich den ControlUndo-Prüffall auf
echte Tastatureingabe statt programmatischer Zuweisung ohne Eingabepuffer.
Original-F129/3075 bleibt erwartete Grenze. Neues DEB native-form-record-fields/
package/linaccess-0.7.0-Linux.deb mit tatsächlichem MDBTools1.0.1/SQLite3.54.1,
JOIN/NULL/Zähler, Payload/Rechten/Build-ID/fehlenden Fixtures geprüft. Uploads3/3
und Fixtures4/4 hashgleich. Keine neue CTest-ID/Gesamtsuite/Windows-Beobachtung.
F056/F050/F088/F052 und456/384/72 unverändert.
Direkt weiter mit echten ApplyFilter-/ShowAllRecords-Dokumentfiltern.


## Fortschreibung 2026-11-09: tatsächliche Dokumentfilter
Gemeinsamer nativer Filterweg für echte Datenblatt-Dialoge und DoCmd/Makro
ApplyFilter/ShowAllRecords beider Hosts. Jet-Prädikat ohne zusätzliche Klauseln/
Subqueries; VBA32768/Makro256 UTF-16-Grenzen, Null94, benannte Filter3251. Tabellen
behalten Entwürfe (Filterwechsel2501), bearbeitbare Abfragen behalten Basis-WHERE.
Readonly-Ergebnisse verwenden tatsächliche Quelle/Parameter und werden erst nach
Erfolg ausgetauscht; Requery behält Filter. Formulare nutzen vorhandenen atomaren
Filter-/Validierungsweg; Berichte nur tatsächliches gebundenes Open-Ereignis,
Kontext nach Verschachtelung wiederhergestellt. DATASHEET_FILTER_RUNTIME.md.
App gebaut; abschließende GUI1/1 (8,62s), Layout1/1 (11,77s) bestanden. Prüffall
sortiert technische Tabellen ausdrücklich statt Qt-Standardreihenfolge zu raten.
Wiederherstellung berücksichtigt Qt: nach fehlgeschlagener Selektion ist Query
inaktiv und setFilter allein selektiert nicht erneut; explizite Auswahl ohne
doppelte erfolgreiche SQL-Ausführung. Reale UI/Cancel/SQL/Dirty/LIKE/Datum/
Parameter-/Abfrage-/Makro-/Form-/Report-Open-Prüfungen, kein neues CTest-Ziel.
Keine Gesamtsuite/Windows-Beobachtung/neue Paketprüfung; DEB073c20a, Original
zuletzt073c20a1/1 (124,97s), erwartetes F129/3075 unverändert offen.
F080/F056/F052/F088/F087 und456/384/72 unverändert.


## Fortschreibung 2026-11-09: echte Fensterbefehle und kürzere Änderungswege
Standardmodule/Makros unterstützen Maximize/Minimize/Restore ohne Ersatzform/Me.
Tatsächliche aktive Form/Report delegiert bestehende native Befehle, Tabellen/
Abfragen/Explorer adressieren echtes Projektfenster. Spezialisierter VBA-
Maximize-Statementpfad überschreibt tatsächliche Host-API; Minimize/Restore
laufen über DoCmd, dafür existieren keine virtuellen Hostmethoden. Modulfenster
3251, andere modale Entwürfe2046; keine implizite Speicherung/Entwurfsverwerfung.
Makroparser/Katalog/Loader einmalig als native_macro.cpp in bestehendem Daten-
Staticlib statt Headerdefinitionen; kleinere GUI-Erweiterungsdatei erhält festen
Einstiegspunkt. Keine neue DSO/GUI-Abhängigkeit des Datenlayers/CTest-ID.
App gebaut; abschließende GUI1/1 (8,69s), Layout1/1 (11,29s) bestanden. Gezielte
Prüffallkorrektur berücksichtigt Qt-Kombination minimiert+zuvor maximiert statt
fälschlich exklusive Statusflags zu verlangen. Reale Fensterzustandsereignisse,
Statement/Makros, unveränderte SQL-/Cache-/Formwerte, extra450, tatsächlicher
Moduleditor mit erhaltenem Entwurf. NATIVE_WINDOW_STATE_RUNTIME.md.
Volle Access-MDI-/Popup-/Geometrie-/Windows-Ereignisparität bleibt offen. Keine
Gesamtsuite/Windows-Beobachtung/neue Paketprüfung; DEB073c20a, Originalzuletzt
073c20a1/1 (124,97s), erwartetes F129/3075 offen. F052/F088/F135 und456/384/72
unverändert. Direkt weiter mit zusammenhängenden Makro-Bedingungsblöcken.


## Fortschreibung 2026-11-09: zusammenhängende Makro-Bedingungsblöcke
Condition=... übernimmt einmal ausgewertetes Ergebnis des unmittelbar vorigen
Bedingungsblocks. Leere Zeile beendet Block, neue Bedingung startet neuen;
Wiederholungen/Unteraufrufe haben eigenen Zustand. Verwaiste Fortsetzungen und
Gruppenanfänge werden im vollständigen Vorabcheck3251 vor SQL-Präfix abgelehnt.
Bestehende Aktions-/Tiefenbudgets zählen übersprungene Einträge weiter. Tatsächlicher
Makroeditor speichert Fortsetzungstext, Standard-/Formhost führen echte SQL-Audits
aus. Runner/Frame-/Stop-Helfer nach native_macro_runner.cpp im bestehenden
Daten-Staticlib ausgelagert, UI bleibt Callback, keine neue GUI-Abhängigkeit/DSO.
App gebaut; bestehende GUI/Layout2/2 (20,25s; 8,89/11,36s) ohne Korrekturlauf
bestanden: echter Editor, Funktion exakt einmal pro Block, True/False-Repeats,
Child isoliert Parent, defekte andere Gruppe/Präfix ohne Ausführung, Folgelauf
sauber. MACRO_CONDITION_BLOCKS_RUNTIME.md; keine neue CTest-ID/Windows-Referenz/
Gesamtsuite. Frisches DEB native-macro-continuation/package/linaccess-0.7.0-Linux.deb
mit Payload/Rechten/Build-ID/echtem MDBTools1.0.1 und SQLite3.54.1-JOIN/NULL/Zählern
geprüft; umfasst auch Filter/Fensterbefehle. Uploads3/3 und Fixtures4/4 hashgleich.
Originalzuletzt073c20a1/1 (124,97s), erwartetes F129/3075 bleibt offen.
Vollständige Access-Makro-/Ereignisparität bleibt offen. F088/F052/F135 und
456/384/72 unverändert. Direkt weiter mit SQL-Zugriff auf tatsächliche skalare
Formular-/Berichts-/Steuerelementeigenschaften ohne Methodenaufrufe.


## Fortschreibung 2026-11-09: echte skalare Objekteigenschaften in SQL
Jet-Tokenadapter unterstützt Forms!Dokument.Eigenschaft und Control.Eigenschaft,
auch Reports/lokalisierte Namen/Bracket-/Punktnamen. Bestehende Werte/.Value
bleiben auf alter3er-API, neue4er-Property-API ist driver-/verbindungsgebunden.
Explizite Getter-Allowlist verwendet tatsächliche Form-/Report-/Control-/
Datensatzfeldobjekte; Undo/Requery/Objektgetter/Methoden werden nicht ausgeführt
(438), tiefere/parametrisierte Pfade bleiben offen. Bisherige skalare SQL-
Konvertierung wiederverwendet, keine neue Typ-/Currency-/Decimal-Vollbehauptung.
Gemeinsamer Lookup prüft GUI-Thread, Connection/Pfad/Kind/echte Lebensdauer,
mehrdeutige Namen5 und nach Getter erneut tatsächliche Dokumentgültigkeit.
SQL_OBJECT_PROPERTIES_RUNTIME.md. App und Ziele gebaut; gebündelt4/4 (115,92s)
ohne Korrekturlauf bestanden: Layout11,68s, GUI9,01s, SQL-Objekte0,02s, Original
115,91s. Reale Caption/Tag/Visible/Dirty/OldValue/ListCount, native versteckte
Feldentwürfe, Live-View nach Änderung, fehlender Kontext, API-Reopen, echte
Verbindungsisolation/Mehrdeutigkeit, Getterverweigerung ohne Undo, Form-/
Reportfehler2450/2451. Original-F129/3075 bleibt erwartete Grenze.
Keine neue CTest-ID/Gesamtsuite/Windows-Beobachtung/Paketprüfung; DEB72ed466.
F067/F068/F118/F164/F050/F052 und456/384/72 unverändert.


## 2026-11-09: nativer VBA-Debugger, F060

Implementiert: tatsächliche gespeicherte Standardmodul-Ausführung mit F9/F5/F8,
StepInto/Over/Out, lebenden Aufrufrahmen/ByRef-Locals, paginierter Array- und
vollständiger Stringinspektion ohne Objektgetter, unbehandelten Fehlern auch
bei normalem F5, Resume-Next-Überspringen und explizitem Abbruch. Keine erneute
Ausführung bereits erfolgter SQL-Aktionen zum Fehlerhalt. Quelle/Projektwechsel
sind während des Halts geschützt; beendete Rahmen werden nicht weiter gelesen.
Details und Grenzen: docs/NATIVE_VBA_DEBUGGER_RUNTIME.md.

Nachweise in bestehenden Zielen: Kern1/1 (2,18s), Layout1/1 (11,80s), originales
Kassa-Ziel1/1 (118,52s), abschließender GUI-Ablauf1/1 (9,06s) sowie separat frisch
GUI-freier Kern1/1 (1,89s). Abschließendes DEB/Build-ID/MDBTools1.0.1/SQLite3.54.1
und vier Fixture-/drei Uploadhashes bestanden. Kein neuer Vollsuite-, Windows-,
Distributionsmatrix- oder GitHub-CI-Nachweis. Original-F129/3075 bleibt offen.
F060 bestätigt diesen beschriebenen nativen Teilumfang; vollständige IDE-,
Makrodebugger-, Sprach-, COM- und Ereignisparität bleibt in den übrigen offenen
Anforderungen. Fester Katalog456:385 bestätigte Teilumfänge/71 offene
Vollanforderungen, ausschließlich theoretisch und ungewichtet. Kein Gesamtabschluss.


## 2026-11-09: projektbezogene Einstufung, F090

Extras-Bericht mit deutschem Statusbaum und explizitem JSON-Export für den
wirklichen Import-/Objekt-/Tabellenkatalog. Getrennte statische Syntaxannahme,
Kontextgrenzen und tatsächlich beobachtete VBA-/Makro-/SELECT-Aufrufe samt
Fehlern und Abbruch. Verschachtelte Aufrufe, Quellrevisionen, nicht gestartete
Prozeduren, ungültige Definitionen und Rollback-Beobachtungen nachgewiesen.
Keine automatische Geschäftsprozedurausführung, keine VBA-/Argument-/
Datensatzwerte im Export. Projektwechsel leert die maximal1024 Sitzungsbelege;
gleichlautende Quellen belegen ausdrücklich keine konkrete Objektinstanz.
Andere SQL-/Tabellenmodellpfade und komplettes Dokumentrendering gelten als
nicht beobachtet, niemals als erfolgreich. Scope: docs/PROJECT_CAPABILITY_REPORT.md.

Gebündelter Lauf: Kern1/1 (2,04s), Layout1/1 (12,09s), Original1/1 (120,17s);
GUI-Testreihenfolge nach absichtlich defektem Modul korrigiert. Abschließende
gezielte GUI/Layout-Nachprüfung2/2 (13,93s; GUI13,93s/Layout12,54s), frisch
GUI-freier Kern1/1 (2,02s). Tatsächliche Makrorevision und gültiger Projektkontext
für Beobachtungen geschützt. Abschließendes DEB/Build-ID/MDBTools1.0.1/
SQLite3.54.1 und vier Fixture-/drei Uploadhashes bestanden. Keine neue CTest-ID,
keine aktuelle Vollsuite/Distributionsmatrix/GitHub-CI behaupten. Original-F129/
3075 bleibt offene Ausführungsgrenze. F090 bestätigt den beschriebenen
Klassifizierungsumfang, keine vollständige Ausführungs- oder Access-Parität.
Fester Katalog456:386 bestätigte Teilumfänge/70 offene Vollanforderungen,
ausschließlich theoretisch und ungewichtet. Kein Gesamtabschluss.


## 2026-11-09: Boolean-Bedingungen und If-Fehlerfortsetzung

Gemeinsame native Let-Boolean-Konvertierung für If/ElseIf, Do-/While-Köpfe und
Fußbedingungen, IIf sowie Makrobedingungen/Wiederholung: skalare Standardwerte
einmal, Null als False, True/False-Strings und case-sensitive #TRUE#/#FALSE#.
CBool(Null) bleibt94; Fehlerwerte/Arrays/Nothing werden nicht als True versteckt.
Belegte Tokenformen mit angrenzenden Klammern, EndIf/End If/End With mit normalen
Tokenabständen und leere einzeilige Then-/Else-Teile. EndIf vor Doppelpunkt ist
kein Label; leeres Case wird vor Effekten verweigert. Fehlgeschlagene If/ElseIf-
Köpfe behalten ausstehende Rahmen für Resume, Resume Next und Debugger-Skip;
keine irrtümlichen Then-/Else-Effekte und keine beschädigten äußeren If-Rahmen.
Scope/Sprachspezifikationsbeleg/Grenzen: docs/VBA_BOOLEAN_CONDITIONS.md.

Gezielte Nachweise nach vollständiger Implementierung/Korrektur der technischen
Prüfszenarien: Kern1/1 (2,22s), GUI1/1 (13,00s), abschließend aktuelles originales
Kassa-Ziel1/1 (113,95s), frisch GUI-freier Kern1/1 (2,00s). Tatsächliche Makro-
SQL-Effekte, Null-Standardwerte/Getterzahl, alle Bedingungspositionen, Resume und
Debugger-Skip geprüft. Finales DEB/Build-ID/MDBTools1.0.1/SQLite3.54.1 sowie
vier Fixture-/drei Uploadhashes bestanden. Kein neuer CTest, kein aktueller
Gesamtsuite-/Windows-/Distributionsmatrix-Nachweis. F129/3075 bleibt erwartet.
F045/F041/F051/F088 und456/386/70 bleiben offen/unverändert; kein Gesamtabschluss.


## 2026-11-09: fehlerhafte Schleifen-/With-Einstiege

Native Resume-Next-/Debugger-Skip-Fortsetzung überspringt fehlgeschlagene
While/Do-/For-/For-Each-Einstiege bis zur tatsächlichen Blockgrenze. Kein
Rumpflauf ohne gültigen Schleifenrahmen; äußere Schleifen/If bleiben erhalten.
With hält die äußere Aktivierung getrennt vom ausstehenden Objektkopf, damit
sein fehlgeschlagener Rumpf inaktiv bleibt und End With den äußeren Zustand
wiederherstellt. Wiederholendes Resume versucht den wirklichen Kopf erneut.
Scope und Grenzen in docs/VBA_BOOLEAN_CONDITIONS.md.

Ein gebündelter bestehender Kern-/GUI-Lauf2/2 (13,20s; Kern2,13s/GUI13,20s)
ohne Korrekturlauf bestanden. Aktuelle App und DEB/Build-ID/MDBTools1.0.1/
SQLite3.54.1 geprüft. Letzter Original- und GUI-freier Kernnachweis bleibt der
unveränderliche vorherige dae6681 (Original113,95s/Kern2,00s), kein neuer
Original-/Qt-frei-/Gesamtsuite-/Windows-/Distributionslauf behaupten.
Keine neue CTest-ID oder Katalog-ID; F045/F051 und456/386/70 bleiben offen/
unverändert. Kein Gesamtabschluss; weitere Arbeit läuft fort.


## 2026-11-09: Datensatzcursor und übergeordnete Requery-Berechnung

Schreibgeschützte Abfrageformulare wählen nach leer→gefüllt den ersten echten
Datensatz. Nur zulässige direkte Tabellenformulare behalten eine neue Leerzeile.
Geänderte tatsächliche Kind-Snapshots aktualisieren berechnete Elternwerte, ohne
Elterneingaben zu speichern; unveränderte Snapshots bleiben zwischengespeichert.
Details und Grenzen: docs/READONLY_FORM_REQUERY_RUNTIME.md.

Gebündelt Layout/GUI2/2 (14,90s), originales Kassa-Ziel1/1 (118,10s), finales
DEB/Build-ID/MDBTools1.0.1/SQLite3.54.1 und vier Fixture-/drei Uploadhashes
bestanden. Originale Keyboard-Grenze jetzt3146/1023; echter getaggter Klick
reicht bis2485 (fehlende exportierte Makroquelle). Beide MDB-Kataloge enthalten
vier Makros, bisheriges Brückenarchiv null. Kein neuer Gesamtsuite-/Windowslauf,
keine vollständige F129-Erfüllung. Katalog456/386/70 bleibt unverändert.


## 2026-11-09: natives Form.RecordsetClone

Eigene DAO-Snapshotcursor über tatsächliche Tabellen-/Abfrageformularergebnisse,
mit unabhängigem Suchen/Navigieren, Bookmark-Rücknavigation, leerem Ergebnis,
Requery-Familienwechsel und weiterlebenden gehaltenen Clones. Eigenschaft ohne
Argumente/schreibgeschützt; keine editierbaren Abfragen behauptet. Bestehende
Layout-/GUI-Prüfung gebündelt2/2 (12,40/14,29s) bestanden, kein neuer Testtarget
oder Windows-/Gesamtsuite-/Installerlauf. Details FORM_RECORDSET_CLONE_RUNTIME.md;
F054/F160 und Katalog456/386/70 bleiben offen/unverändert.


## 2026-11-09: tatsächliche Dateisammlungen und Unterordner

Native Scripting.Files/Folders unterstützen Item(Key), Standardzugriff und
indizierte Folder.Files/SubFolders-Eigenschaften über echte direkte Linux-Kinder.
SubFolders.Add(Name) erzeugt einen wirklichen Unterordner; Duplikate/Null/fehlende
Einträge/ungeeignete Pfade bleiben explizite Fehler. Native Namens-/Snapshotgrenzen
siehe NATIVE_FILE_SYSTEM_OBJECT.md; keine Windows-Gesamtparität behauptet.

Bestehender Kernlauf nach Implementierung gebündelt: Weiterleitung indizierter
Elterneigenschaften korrigiert, neue technische Unterordner vor dem vorhandenen
Symlink-Schutzfall entfernt; derselbe abschließende Kern1/1 (2,17s) bestanden.
App frisch gebaut. Kein neuer CTest, kein neuer vollständiger Desktop-/Windows-/
Qt-frei-/Installernachweis. F091 und Katalog456/386/70 bleiben offen/unverändert.


## 2026-11-09: direkter Formular-Recordset-Feldzugriff

Me.Recordset.Fields unterstützt tatsächliche Felder per Name/Ordinal, Count,
Item, Standard-/Bang-Zugriff und typisiertes For Each. Gehaltene Felder folgen
der aktuellen gespeicherten Zeile; Requery ersetzt ihre Identität (3420),
Formularschluss sperrt weitere Zugriffe (2450). Tatsächliche BLOB-Werte, FieldSize
und GetChunk; keine impliziten Speicherungen oder Feldschreibrechte. Nicht
beweisbare eingefrorene Metadaten bleiben Fehler; DataUpdatable=False beschreibt
den tatsächlich schreibgeschützten Adapter. Details FORM_RECORDSET_FIELDS_RUNTIME.md.

Gebündelt GUI1/1 (14,05s) bestanden. Layoutlauf fand fehlende DataUpdatable-
Fähigkeitsauskunft; korrigiert und nur dieses fehlgeschlagene Ziel erneut:
Layout1/1 (12,76s) bestanden. App gebaut, keine neue CTest-ID, kein neuer
Windows-/Gesamtsuite-/Qt-frei-/Installerlauf. F054/F160 und456/386/70 bleiben
offen/unverändert. Letzter Kern d2488e0:2,17s; letzte DEB-Prüfung9eae43f.


## 2026-11-09: echter Rückgängig-Verlauf im Formulardesigner

Natives Undo/Redo stellt Feldreihenfolge/-auswahl, Beschriftungen, Twip-Geometrie,
Datenquelle, VBA samt importierten Originalbytes, Form-/Feld-Ereignisse und Timer
im geöffneten Entwurf wieder her. Erfolgreiches Speichern setzt den tatsächlichen
Vergleichsstand; Rückkehr dorthin löscht den Änderungsstatus. Kein Undo-Speichern
von Datenbankobjekten; Verlauf bleibt begrenzt im Speicher, kein VBA-Historyexport.
Textfelder behalten ihr eigenes Text-Undo. Grenzen FORM_DESIGNER_HISTORY.md.

Bestehender GUI-Lauf1/1 zunächst13,67s; unverändert bestätigte Ereignisdialoge
als Nichtänderung bereinigt, abschließend dasselbe Ziel1/1 (13,52s) bestanden.
Neue aktuelle DEB-/Build-ID-/MDBTools1.0.1-/SQLite3.54.1-Prüfung bestanden, alle
vier Fixture- und drei Uploadhashes unverändert. Keine neue CTest-ID, kein neuer
Windows-/Gesamtsuite-/Qt-frei-Nachweis. Letzter Kern d2488e0:2,17s; letzter
Layout2a3a781:12,76s; letzter Original9eae43f:118,10s. F080/F148 und
Katalog456/386/70 bleiben offen/unverändert. Allgemeine Formularerstellung folgt.


## 2026-11-09: freie Formularerstellung und eigene Steuerelemente

Native leere/ungebundene Formulare und Formulare ohne ausgewählte gebundene Felder
sind speicherbar. Steuerelementdialog erzeugt bis128 tatsächliche TextBox/Label/
CommandButton/CheckBox/ComboBox/ListBox mit geprüften Namen, Twip-Geometrie,
Bindungen, Listenherkunft und Klickereignis. Vorhandene VBA-Ereignisse binden
automatisch; eigener Entwurf einschließlich Controls nimmt am Undo/Redo teil.
Ungebundene Ansicht zeigt keinen erfundenen Datensatz oder Quellenfehler an.
Umfang/Grenzen FREE_FORM_CONTROLS.md, bestehende F080/F081/F083/F148/F155.

Gebündelter bestehender Lauf: Layout1/1 (12,02s) bestanden. GUI fand einen
Prüfklick außerhalb der tatsächlichen Checkbox-Markierung; auf echten Qt-
Indikator korrigiert und nur dieses Ziel erneut: GUI1/1 (13,08s) bestanden.
Tatsächlicher Designer, Duplikatfehler/Korrektur, Save/Reload, Eingabe/Berechnung,
Label/Checkbox/Combo/List und unveränderte technische VBA mit unabhängig perSQL
bestätigtem Klick-Effekt geprüft. App gebaut; kein neuer CTest oder Katalogpunkt,
kein neuer Original-/Qt-frei-/Windows-/Gesamtsuite-/DEB-Lauf. Letzte geprüfte
DEB a837563; Katalog456/386/70 bleibt unverändert, keine Gesamtparitätsbehauptung.


## 2026-11-09: literale Formular-Feldbindungen

ControlSource dekodiert ]] in geklammerten Feldnamen konsistent für Laden,
Speichern, Feldreferenz auf vorhandene Alias-Controls und OldValue. Native
Generatoren klammern/escapen Feldnamen mit führendem =, äußeren Klammern oder
äußeren Leerzeichen. Keine Umbenennung von Spalten oder Originalquellen.
Native tatsächliche Lese-/Schreib-/Alias-/OldValue-/SQL-Nachweise im bestehenden
Layout-/GUI-Ziel abschließend2/2 (12,39/14,16s), aktuelle DEB-/Build-ID-/
MDBTools1.0.1-/SQLite3.54.1-Prüfung bestanden. Details
LITERAL_CONTROL_SOURCE_BINDINGS.md. Kein neuer CTest oder Katalogpunkt; keine
aktuelle Gesamtsuite/Windows/Original-Workflow-Ausführung behauptet. F155/F056/
F160 und456/386/70 bleiben offen/unverändert. Mehrspaltige native Listen folgen.


## 2026-11-09: mehrspaltige eigene Kombinations-/Listenfelder

Steuerelementdialog speichert/geprüft ColumnCount, BoundColumn und ColumnWidths,
inklusive getrennter Anzeige-/Wertespalte, gebundenem Zeilenindex und versteckten
Spalten. Gemeinsamer tatsächlicher Twip/cm/mm/in/pt-Parser für Designer und
Runtime; keine abweichend erfundene Breiteninterpretation. Bestehende Definitionen
behalten ihre Standardwerte. Native Save/Reload-/ValueList-/SQL-Modelle und echte
Mausauswahl im bestehenden GUI/Layoutlauf2/2 (13,83/12,51s) bestanden.
Details FREE_FORM_CONTROLS.md. Kein neuer CTest/Windows-/Original-/Gesamtsuite-/
DEB-Lauf; letzte DEB15131f0. F080/F083/F162 und456/386/70 bleiben offen/unverändert.


## 2026-11-09: eigene Standardwerte, TextBox-Format und Passwortmaske

Steuerelementdialog speichert DefaultValue/Format/InputMask mit bestehenden
Typ-/Längen-/Kontextgrenzen. Designer und tatsächliche Initialisierung verwenden
denselben nativen Literalparser; keine berechneten Standardwerte im Designer
ausführen. Format bleibt Darstellung bei unverändert typisiertem Value, Password
ist native Echo-Policy. Echte Create/Save/Reload-/Default-/Wert-vs-Anzeige-/SQL-
Nachweise gebündelt im bestehenden GUI/Layout2/2 (14,98/12,78s) bestanden.
Details FREE_FORM_CONTROLS.md. Kein neuer CTest, kein Kern-/Windows-/Original-/
Gesamtsuite-/DEB-Lauf; letzte DEB15131f0. F015/F071/F080/F081/F115 und456/386/70
bleiben offen/unverändert. Allgemeine Steuerelementeigenschaften folgen.


## 2026-11-09: eigene Zustände und Tabreihenfolge

Auswahlpanel für Visible/Enabled/Locked/TabStop/TabIndex im echten Controls-
Designer, mit Typ-/Bereichsprüfung und bestehenden nativen Laufzeitregeln.
Labels ohne Tab-Eigenschaften, Locked für TextBox/ListBox. Tatsächlich
speichern/neuladen, ausgeblendetes deaktiviertes Label, gesperrtes maskiertes
Textfeld, Tab-Ausschluss und echte Tastatur-Zielwahl im bestehenden Layout/GUI
gebündelt2/2 (12,28/14,29s) bestanden. Details FREE_FORM_CONTROLS.md.
Keine neue CTest-ID oder aktuelle Gesamtsuite/Windows/Original/DEB-Prüfung;
letzte DEB15131f0. F080/F081/F104/F148 und456/386/70 bleiben offen/unverändert.


## 2026-11-09: echte Schrift-/Farbauswahl eigener Controls

Qt-Auswahldialoge speichern geprüfte FontName/Size/Weight/Italic/Underline und
RGB-ForeColor; TextBox/Label zusätzlich BackColor/BackStyle transparent/opaque.
Standarddarstellung löscht Overrides. Persistenter Zieleintrag plus Auswahl-/
Fokuswiederherstellung verhindert nach Modalrückkehr die Änderung eines anderen
Controls. Keine Windows-Font-/Pixelparität behaupten. FREE_FORM_CONTROLS.md.

Tatsächliche Picker/Mausklicks, Save/Undo/Redo/Reload, native Paletten/Transparenz
und aufgelöster QFontInfo-Schnitt geprüft. GUI-Testsetup auf wirklich installierten
Fett/Kursiv-Schnitt korrigiert; Modal-Auswahlwechsel behoben. Abschließend nur
fehlgeschlagenes GUI-Ziel1/1 (14,70s) wiederholt/bestanden; Layout1/1 (13,09s)
bereits bestanden. Kein neuer CTest, keine Gesamtsuite/Original/Windows/Kern-
Ausführung; aktuelle DEB-Prüfung folgt.456/386/70 unverändert; F080/F081/F148 offen.


## 2026-11-09: Ausrichtung, Tag und lange Eigenschaftstexte

Native Controls speichern 2048-Zeichen-Tag sowie TextBox-/Label-Ausrichtung
Allgemein/Links/Zentriert/Rechts. Quellen-/Typ-/Bereichsgrenzen bleiben explizit.
Echter Tabellen-Delegate übernimmt den bestehenden65536-Eigenschaftsrahmen statt
32.767 stiller Eingabekürzung; kurze Teilpropertygrenzen gelten weiter.
GUI-Editor bearbeitet33.000 Zeichen+Anhang, Save/Reload prüft33.001, native
Ausrichtung/Tag ebenfalls. Offscreen-Dialog für echte Eingabe explizit aktiviert;
abschließend nur GUI1/1 (14,11s) wiederholt/bestanden; Layout1/1 (12,81s) bereits
bestanden. FREE_FORM_CONTROLS.md samt tatsächlicher Microsoft-Tag-Dokumentation.
Kein neuer CTest/Windows/Original/Kern/Gesamtsuite/DEB-Nachweis; letzteDEB d8ecf41.
F050/F080/F081/F148/F155 und456/386/70 unverändert. SQL-Kompatibilität folgt.


## 2026-11-09: Typisierte Kreuztabellen-Minima und -Maxima

Feste TRANSFORM-Min/Max-Überschriften verwenden echte physische Tabellen und
vorhandene native Currency-/Single-/Ganzzahl-/Boolean-/Datumsprüfungen. Pivottypen
bleiben erhalten, auch bei leeren Ergebnissen; fehlende Kategorien liefern Null.
Parameter, Aliasse, gespeicherte QueryDefs und dbQCrosstab16 sind eingebunden.
Text/Locale, unbestimmte Integerbreiten, Decimal, gespeicherte Quellansichten und
temporäre Überschattung werden nicht als unterstützt ausgegeben. Schemaänderungen
invalidieren typisierte Abfrageansichten. Details: COUNT_CROSSTAB_RUNTIME.md.
Abschließend bestehendes SQL-Bündel3/3 bestanden (0,12s); Anwendung neu gebaut.
Kein neuer CTest/Windows-/Original-/Gesamtsuiten-/Paketnachweis. F067/F068 bleiben
offen; Katalog456/386 bestätigte Teilumfänge/70 offene Vollanforderungen unverändert.


## 2026-11-09: Statistische SQL-Aggregate

Var und StDev sind in echte SQL-Ausführung eingebunden, mit bestehendem nativen
Berichts-/Domänen-Akkumulator, Double-Metadaten, n-1, Null-Auslassung und Null bei
weniger als zwei Werten. Gruppierung/DISTINCT/Parameter und große gemeinsame
Offsets sind geprüft; Text/BLOB13 und Überlauf6 bleiben explizit.
SQL_STATISTICAL_AGGREGATES.md beschreibt Umfang und Microsoft-Dokumentationsbasis.
Abschließend bestehendes SQL-Bündel4/4 bestanden (0,13s), Anwendung neu gebaut;
keine neuen CTests oder Windows-/Original-/Gesamtsuiten-Nachweise. Noch kein neues
Paket; folgt aus dem aktuellen Stand. F067/F068/F071 und456/386/70 unverändert.


## 2026-11-09: Berechnete Domänenausdrücke

DLookup/DFirst/DLast/DMin/DMax/DAvg/DVar/DStDev unterstützen begrenzte skalare
SQL-Ausdrücke, statt ausschließlich Feldnamen. Echte Projektion/Typen/Null, Currency-
und Datumsidentität, Zahlen-/Textausdrücke und statistische Ergebnisse sind im
bestehenden nativen VBA-Host geprüft. Tokenprüfung sperrt Zusatzklauseln, Subqueries,
verschachtelte Aggregate/Window-Aufrufe, mehrere Projektionen und fremde Funktionen.
NATIVE_DLOOKUP.md hält Grenzen fest; keine Text-/Locale-Extrema oder beliebige
Projektfunktionsbindung. Abfrageansichten invalidieren nun auch bei temporären
Tabellen, einschließlich tatsächlichem Entfernen/Wiederherstellen der zuvor
validierten Min/Max-Ansicht bei Überschattung. Bestehendes Bündel5/5 (12,21s)
bestanden, Anwendung gebaut; kein neuer CTest/Windows-/Original-/Gesamtsuitenlauf.
Letzte geprüfte DEB ist eaaeb73, inklusive Build-ID/MDBTools1.0.1/SQLite3.54.1;
alle vier Fixture-/drei Uploadhashes unverändert. Diese neue Domänenetappe ist
noch nicht paketiert. F050/F067/F068/F118/F164 und456/386/70 unverändert.


## 2026-11-09: Native bedingte VBA-Kompilierung

Qt-freier nativer Preprocessor für #Const/#If/#ElseIf/#Else/#EndIf/#End If samt
begrenzter Fortsetzungen. Auch ausgeschlossene Konstanten/Bedingungen werden
entsprechend MS-VBAL3.4/5.6.16.2 ausgewertet; aktive Modulfelder/Prozeduren,
OptionBase/Compare, lokale/öffentliche/SQL-Aufrufe und native Ereignisbindungen
verwenden eingeschlossenen Code. Quellenbytes/Hashes und Debuggeridentitäten
bleiben ursprünglich; echte Haltepunkte prüfen Originalzeilen. Linux-Hostflags
und genaue Grenzen in VBA_CONDITIONAL_COMPILATION.md; keine Windows-API-Fiktion.
Abhängigkeits-/Nesting-/Textbudgets sperren Zyklen und exponentielle Expansion.
Ein fehlender Array-Header/eine Syntaxklammer wurden beim Build korrigiert. Im
gebündelten Prüflauf bestanden7/8; Kernelprüfung fand falsches Zahlenfeld beim
Boolean-Lesen. Exakt auf .boolean korrigiert; danach gezielt Kernel/DAO2/2
(2,67s) bestanden, echter SQL-UDF-True-Zweig/Typ ebenfalls. GUI12,88s und15,34s
waren bereits bestanden und wurden für diesen Einzelfix nicht wiederholt.
Separater Qt-freier Build/Kernlauf ebenfalls bestanden; ldd ohne Qt. Keine neuen
CTest-IDs/Original-/Windows-/frische Gesamtsuitenbehauptung. Anwendung neu gebaut;
letzte geprüfte DEB weiterhin eaaeb73. F040/F042/F043/F045/F051/F068 bleiben offen;
Katalog456/386 bestätigte Teilumfänge/70 offene Vollanforderungen unverändert.


## 2026-11-09: Stop und Debug.Assert im echten Debugger

Native Interpreter-Anweisungen halten mit eigenen Stop-/Assertion-Gründen an
Originalzeilen und lebenden Variablen. Fortsetzen/Schritt/Beenden, eager vollständige
Boolean-Auswertung, einmaliger Standardgetter und Err-Erhalt unter ResumeNext
sind im bestehenden Umfang eingebunden. Ohne Debuggersitzung wird benötigter Halt
3251 gemeldet; kein stiller Erfolg. Echter Modul-Editor/Debugger prüft Status,
Variable17 und tatsächliche SQL-Fortsetzung18 mit anschließender Test-Rücknahme.
Bestehendes Kernel-/GUI-Bündel2/2 (14,49s) bestanden; Anwendung gebaut. Keine neuen
CTest-IDs oder weiteren Original-/Windows-/Gesamtsuiten-/Qt-frei-Nachweise für diese
Stop-Erweiterung; Details in NATIVE_VBA_DEBUGGER_RUNTIME.md.

GitHub-CLI-Token wurde nach931ea64 ungültig. Vorhandener GitHub-Connector besitzt
weiterhin bestätigten Schreibzugriff auf das private Repository. Die vorige
bedingte-Kompilierung-Etappe wurde darüber als1a599f6 veröffentlicht und per Ref
verifiziert: gleicher vollständiger Git-Dateibaum37583763b86188457803ca22944c4911e2c9c9af
wie lokaler6f95ca0, lediglich Veröffentlichungsmetadaten/neue Commit-ID. Ungepushtes
Originalobjekt bleibt als lokaler Sicherungsref erhalten; Arbeitsdateien unverändert.
Kein Token wurde gelesen/ersetzt/gespeichert. Wiederverwendbarer kontrollierter
Veröffentlichungshelfer außerhalb Git: /workspace/work/github-connector/publish.mjs
(über functions.exec/tools ausführen; Blob-/Tree-SHA, private Sichtbarkeit und
Fast-forward-Lease werden geprüft, lokale identische Remote-Commitobjekte übernommen).
F040/F045/F051/F060-Status und Katalog456/386/70 unverändert. Letzte geprüfteDEB eaaeb73;
frisches Paket folgt aus dem aktuellen Quellstand.


## 2026-11-09: Feste skalare VBA-Strings

Lokale/statische und modulweite String*n-Variablen mit1..65526 UTF-16-Einheiten,
bekannten Const-Längen und späteren Modulkonstanten sind implementiert. Initial
NUL-Zeichen, bei Zuweisung Abschneiden/Leerzeichenauffüllung; echte String8-Werte,
Metadatenbreite und statischer Zustand bleiben erhalten. LSet/RSet und Mid sind
eingebunden. Die Kernprüfung fand fehlende Mid$-Zuweisungssyntax; dieser konkrete
Parserpfad wurde ergänzt, inklusive unverwechselbarem intrinsischem Mid$-Modus.
Null94 verändert den vorigen Wert nicht; isolierte Surrogatabschneidung bleibt3251.
Details/Grenzen einschließlich noch fehlender Arrays/UDTs/kompletter ByRef-Parität:
VBA_FIXED_STRINGS.md. GUI14,77s bereits bestanden; nach Mid$-Fix ausschließlich
Kern1/1 (2,31s) wiederholt/bestanden, Anwendung neu gebaut. Keine neuen CTests oder
neuen Windows-/Original-/Gesamtsuiten-/Qt-frei-Belege für diese String-Erweiterung.
Vorige Stop-Etappe ist als72b936d über Connector veröffentlicht, gleicher geprüfter
Baumccd6e1f54fcbed0d1b72a5d4ff4cb6416d07c2be wie lokaler633fb98. Frische72b936d-DEB
unter /workspace/work/native-debug-halts/package/linaccess-0.7.0-Linux.deb besteht
Build-ID/Inhalt/Berechtigungen/MDBTools1.0.1/SQLite3.54.1-Prüfung. Diese neuen
String-Quellen sind noch nicht paketiert. F040/F041/F042/F043/F122 und456/386/70
unverändert. Keine Gesamtparität/Abschlussbehauptung.


## 2026-11-09: Feste Stringelemente in VBA-Arrays

Feste/dynamische/modulweite/statische String*n-Arrays unterstützen Initial-NUL,
Element-Konvertierung, Trunkierung/Leerzeichen, ReDim(Preserve), Erase und begrenzte
Variant-Kopien. Canonical-Elementmetadaten behalten die Breite unabhängig von
öffentlichem String()/8200. Fehlende Null-Zuweisung und Breitenänderung lassen
vorige Werte/Dimensionen unverändert. Neue Elemente/Erase-Festarrays werden NUL,
dynamisches Erase behält den deklarierten Typ.16MiB-Payloadgrenze prüft vor
Reservierung count*n*3 (UTF-8-Obergrenze pro UTF-16-Einheit); bestehende Bounds/
Iterator-/ByRef- und Millionen-Element-Grenzen bleiben aktiv. VBA_FIXED_STRINGS.md.
Skalare String-Prüfung aus großer core_tests.cpp in kleine vba_fixed_string_tests.cpp
innerhalb desselben CTest-Ziels verschoben; keine neuen Testsuite-IDs.
Abschließend Kernel/GUI2/2 (15,24s), Qt-freier Kern1/1 (2,58s) bestanden,
Anwendung gebaut. Kein neuer Original-/Windows-/Gesamtsuiten-/Paketnachweis.
Vorige skalare Etappe als e9ce132 veröffentlicht/verifiziert (gleicher getesteter
Baumce34e6522e2363f234c2ca20c23cca9da0f57395 wie lokaler d053a80). Letzte geprüfte
DEB weiterhin72b936d, frisches Paket folgt. F040/F041/F042/F043/F122 bleiben offen;
456/386 bestätigte Teilumfänge/70 offene Vollanforderungen unverändert.


## 2026-11-09: Visueller nativer Abfrageentwurf

SQL-Editor hat einen echten Qt-Entwurf für reale Tabellen/Felder: Ausgabeauswahl,
Alias/Feldreihenfolge, Sortierung, Kriterien und DISTINCT erzeugen tatsächliches
SELECT. Vorschau bereitet Metadaten ohne Daten-/VBA-Auswertung vor; erst Ausführen
liest die echten Zeilen. Abbruch ändert SQL nicht. Gespeicherter Entwurf ist an
exakten SQL-SHA gebunden; manuelle SQL-Änderung entfernt alte visuelle Metadaten.
Projekt-/Generations-Eingangsprüfungen und Ressourcen-/Namen-/Klauselgrenzen;
fortgeschrittenes SQL bleibt im SQL-Editor. Kein Kassa-spezifischer Designer.
Scopegrenzen in VISUAL_QUERY_DESIGNER.md, F067/F068/F080/F148 bleiben offen.

Im bestehenden SQL-/GUI-Bündel bestand SQL; GUI fand einen realen QJsonValueRef-
Kopierfehler beim Verschieben. Wertkopie korrigiert, danach nur betroffene GUI1/1
(14,73s) wiederholt/bestanden. Echte verschachtelte Dialoge, unabhängiges SQL,
Speichern/Wiederherstellen, Abbruch und Callback-Zähler beweisen tatsächlichen
Pfad ohne Metadaten-Seiteneffekte. Anwendung gebaut; keine neue CTest-ID/Windows-/
Original-/Gesamtsuitenbehauptung. Fester Katalog456/386/70 unverändert.
Vorige Stringarray-Etappe als1dafedd veröffentlicht/verifiziert (gleich geprüfter
Baum023f57fab171cac5324e8853bb2bb7c026c38c01 wie lokaler145786b). Neue1dafedd-DEB
unter /workspace/work/fixed-vba-string-arrays/package/linaccess-0.7.0-Linux.deb besteht
Build-ID/Inhalt/Berechtigungen/MDBTools1.0.1/SQLite3.54.1. Aktueller visueller Entwurf
noch nicht paketiert. Weiterentwicklung läuft; kein Gesamtabschluss.


## 2026-11-09: ODER-Kriterien und typisierte Entwurfsparameter

Visueller Abfrageentwurf unterstützt bis8 UND/ODER-Gruppen und64 deklarierte
Parameter (Text/Long/Integer/Byte/Single/Double/Currency/DateTime/YesNo). Echte
PARAMETERS-Deklaration, bestehende Eingabedialoge und SQL-/DAO-Bindung; Antworten
werden nicht gespeichert. Metadatenvorschau bleibt ohne Zeilenauswertung. Gruppen/
Parameter öffnen wieder korrekt, Feldverschiebung behält alle Kriterien. Alte
Ein-Gruppen-Entwürfe bleiben lesbar; genaue Grenzen in VISUAL_QUERY_DESIGNER.md.
Bestehendes GUI-/SQL-Bündel2/2 (14,51s) bestanden, inklusive tatsächlicher zweier
Parametereingaben, unabhängigem SQL und späterer manueller Metadateninvalidierung.
Keine neuen CTests oder Kern-/Original-/Windows-/Gesamtsuiten-/Paketnachweise.
Voriger visueller Entwurf als33cda14 veröffentlicht/verifiziert; gleicher geprüfter
Baum6e85a5c233a9ab51f14542cb36e703905789ad12 wie lokaler8e5458f. Letzte geprüfteDEB
weiterhin1dafedd. F067/F068/F080/F148 und456/386/70 unverändert; Entwicklung läuft.


Berechnete Felder im visuellen Abfrageentwurf: echte skalare Ausdrücke mit Alias,
Ausgabe, Sortierung und UND/ODER-Kriterien; Hinzufügen/Entfernen betrifft nur den
Entwurf. Gespeicherte Ausdrücke öffnen wieder korrekt. Begrenzte native skalare
Funktionen; Aggregate, Unterabfragen, unbekannte Callbacks und externe/lebende
Werte werden in diesem Entwurf ausdrücklich begrenzt. Keine Zeilenauswertung in
Vorschau/Übernahme. Details: docs/VISUAL_QUERY_DESIGNER.md. Anwendung gebaut;
bestehendes GUI-/SQL-Bündel2/2 bestanden (15,09s), inklusive wirklicher berechneter
Ergebnisse gegen unabhängiges SQL, Sortierung und Wiederöffnung. Keine neuen
CTest-IDs oder Windows-/Original-/Gesamtsuiten-/Paketnachweise. Parameter/Oder-
Etappe als b1b5858 veröffentlicht; letzte geprüfteDEB weiterhin1dafedd.
Katalog456/386/70 unverändert; keine vollständige Abfrageentwurfsparität behauptet.


Gruppierung im visuellen Abfrageentwurf: pro Feld Group By/Where/Count/Sum/Avg/
Min/Max/Var/StDev, echte GROUP BY-/WHERE-/HAVING-Klauseln mit getrennten ODER-
Gruppen; Where-Felder nur ausgeblendet/unsortiert. Entwurfsmetadaten behalten
Funktionen und Modus. Nachträglich hinzugekommene Quellspalten verändern keine
bestehende Gruppierung. Details: docs/VISUAL_QUERY_DESIGNER.md. Bestehendes GUI-/
SQL-Bündel2/2 (14,62s) bestanden; nach Korrektur der Schemaerweiterung nur betroffene
GUI-Prüfung1/1 (14,36s) wiederholt. Anwendung gebaut; keine neuen CTest-IDs oder
Windows-/Original-/Gesamtsuiten-/Paketnachweise. Berechnete-Felder-Etappe als
cee05d8 veröffentlicht/verifiziert; letzte geprüfteDEB weiterhin1dafedd.
F067/F068/F080/F148 sowie fester Katalog456/386/70 bleiben unverändert offen.


Visueller Datenquellenentwurf im nativen Berichtseditor: derselbe tatsächliche
SQL-Entwurf einschließlich berechneter Felder/Gruppierung; Übernahme als ein
Text-Undo, Abbrechen unverändert, Metadaten nur bei passendem SQL-Hash. Speichern/
Wiederöffnen und Invalidierung bei manueller SQL-Änderung nachgewiesen. Keine
Parameterdeklarationen in diesem Berichtsweg, da Berichtsparameterbindung noch
fehlt. Projektidentität/Generation wird bei Entwurf/Definition/Speichern/Ausgabe
geprüft. Bestehendes GUI-/SQL-Bündel2/2 (15,31s), Anwendung gebaut. Tatsächlicher
Callbackzähler: Entwurf und Speichern ohne Auswertung, explizite Ausführung liest
Zeilen. Keine neuen CTests oder Windows-/Original-/Gesamtsuiten-/Paketnachweise.
Details: docs/VISUAL_QUERY_DESIGNER.md. Gruppierungs-Etappe als31c50d2 veröffentlicht;
letzte geprüfteDEB weiterhin1dafedd. F080/F085/F087/F148 und456/386/70 unverändert.


Gespeicherte Auswahlabfragen als visuelle Datenquelle: tatsächliche native Query-
Views werden vor Entwurfsöffnung installiert und zusammen mit Tabellen angeboten;
Metadatenvorschau bleibt ohne Zeilenauswertung. Tatsächliche verschachtelte Abfrage
auf gespeichertem Entwurf inklusive berechnetem Feld/Kriterium/Sortierung und
Wiederöffnung geprüft; Callbackzähler trennt Vorschau von Ausführung. Native
Abhängigkeits-/Typ-/Parametergrenzen bleiben ausdrücklich erhalten. Anwendung
gebaut, bestehendes GUI-/SQL-Bündel2/2 (15,84s) bestanden. Ein Qt-Enum-Baufehler
wurde vor dem Prüflauf korrigiert. Keine neuen CTests oder Windows-/Original-/
Gesamtsuiten-/Paketnachweise. Berichtsentwurf als4e20658 veröffentlicht; letzte
geprüfteDEB weiterhin1dafedd. Details VISUAL_QUERY_DESIGNER.md;456/386/70 unverändert.


Visuelle Verknüpfungen: bis8 geordnete INNER-/LEFT-Equijoins auf tatsächlichen
Tabellen/gespeicherten Auswahlabfragen, pro neuer Quelle vorhandenes Feld = neues
Feld. Vollqualifizierte Feldidentität und persistierte Quellzuordnung; echte
Nullwerte bei fehlendem LEFT-Treffer. Entfernen betrifft nur Entwurfsfelder;
Selbstverknüpfungen/Aliase/zusammengesetzte Vergleichsbedingungen bleiben im
SQL-Editor mit bestehenden Grenzen. Strikte Metadatenprüfung; fehlgeschlagenes
Einlesen hinterlässt keine halbfertigen Feldzeilen. Bestehendes GUI-/SQL-Bündel2/2
(14,92s), nach Metadatenguards nur GUI1/1 (14,90s) wiederholt; Anwendung gebaut.
Tatsächliche zweistufige LEFT-Verknüpfung gegen unabhängiges SQL verglichen,
Wiederöffnung/Entfernen/INNER-Wechsel und kaputte Metadaten geprüft. Keine neuen
CTests oder Windows-/Original-/Gesamtsuiten-/Paketnachweise. Details in
VISUAL_QUERY_DESIGNER.md. Gespeicherte-Abfragequellen-Etappe28e3302 veröffentlicht;
letzte geprüfteDEB weiterhin1dafedd. F067/F068/F080/F148 und456/386/70 unverändert.


Visueller Formulardatenquellenentwurf: derselbe echte SQL-Entwurf erzeugt SELECT-
Quellen mit vorbereiteten Feldmetadaten; keine Parameterbindung und weiterhin
schreibgeschützte Abfrageformulare. Passende SQL-Hash-/Entwurfsmetadaten werden
mitgespeichert und im bestehenden Formular-Undo/Redo erfasst/wiederhergestellt.
Manuelle Quelländerung entfernt veraltete Metadaten. Projektidentitäts-/Generations-
prüfung bei Entwurf/Speichern/Verlauf. Bestehendes GUI-/SQL-Bündel2/2 bestanden
(15,47s), Anwendung gebaut; tatsächliches gespeichertes Formular mit echten
Ergebniswerten/schreibgeschütztem Zustand geöffnet. Callback zählt keine Auswertung
bei Entwurf/Speichern/Wiederöffnung, erst tatsächliche Formularausführung liest.
Keine neuen CTests oder Windows-/Original-/Gesamtsuitennachweise.

Verknüpfungs-Etappe alsfe01e1d veröffentlicht/verifiziert. Geprüftes DEB dieses
Standes unter /workspace/work/visual-query-joins/package/linaccess-0.7.0-Linux.deb:
Build-ID passt, MDBTools1.0.1/SQLite3.54.1, Payload/Berechtigungen/keine Fixtures
bestanden. Paket enthält noch nicht die anschließende Formularintegration. Vier
Fixture- und drei Upload-Originalhashes erneut unverändert geprüft. Details in
VISUAL_QUERY_DESIGNER.md/NATIVE_QUERY_FORM_DESIGN.md. F050/F080/F148 und fester
Katalog456/386/70 unverändert. Nächste unabhängige Untersuchung: ursprüngliche
3146/1023-Grenze beim Öffnen des Originalformulars, keine erneute Makro-Rückfrage.


Korrektur der bisherigen Original-3146/1023-Deutung: private Diagnose zeigte
Document initialization failed: Native macro catalog unavailable bei source-error0.
Die Import-Testkopie besitzt keinen optionalen main._la_objects-Katalog. Native
Makrosuche behandelt dessen Abwesenheit jetzt wie einen leeren Katalog (2485),
ohne Tabellen anzulegen oder temporäre Schatten zu lesen; vorhandene defekte
Kataloge bleiben3146. Beide tatsächlichen Tastaturhandler erreichen nun den
fehlenden Makroquelltext2485 in Zeile1023. Das ist keine offene Datensatzquellen-
SQL-Operation. Getaggter Click erreicht separat ebenfalls fehlenden Makroquelltext.
Bestehendes GUI-/Original-Bündel2/2 bestanden (125,46s; GUI15,85s); Anwendung gebaut.
Isolierte Prüfungen für fehlenden/temporären/defekten/leeren Hauptkatalog sind in
bestehendem Ziel enthalten. Keine neuen CTests, Windows- oder positive Original-
Verkaufs-/Makrolaufnachweise. Die diagnostischen Instrumentierungen/privaten Details
liegen ausschließlich außerhalb Git; sie sind keine zusätzlichen erfolgreichen
Prüfziele. Native Makroquelle fehlt weiterhin im Bridge-Export; vorhandene
Codex-Call-Rückfrage dazu blieb unbeantwortet, kein wiederholter Anruf nötig.
Formularintegration als3b18277 veröffentlicht; letzte geprüfteDEB fe01e1d vor dieser
Integration. Details NATIVE_MACROS_RUNTIME.md/READONLY_FORM_REQUERY_RUNTIME.md.
F129-Text korrigiert, fester Katalog456/386/70 unverändert. Weiter mit echten
Quellenaliasen und Selbstverknüpfungen im visuellen Entwurf.


Quellenaliase/Selbstverknüpfungen im visuellen Entwurf: wiederholte echte Tabellen/
Query-Views mit eindeutigem explizitem Alias, getrennte physische Herkunft und
SQL-Feldidentität, Metadaten immer von der tatsächlichen Quelle. Linker Selektor
bietet vorhandene Aliase und zugehörige Felder, Tooltip zeigt Herkunft. Summen-
dialog qualifiziert doppelte Feldnamen. Reale Eltern-Kind-LEFT-Verknüpfung gegen
unabhängiges SQL verglichen, Nullfälle/Alias mit Leerzeichen/Wiederöffnung und
Ablehnung mehrdeutiger aliasloser Wiederverwendung geprüft; Metadaten führen
keine Zeilenberechnung aus. Anwendung gebaut, bestehendes GUI-/SQL-Bündel2/2
(16,45s) bestanden ohne weitere Wiederholung. Keine neuen CTests oder neue
Original-/Windows-/Gesamtsuiten-/Paketnachweise. Details VISUAL_QUERY_DESIGNER.md;
F080-Beschreibung aktualisiert, Katalog456/386/70 unverändert. Makrokatalog-Fix
als657d6b3 veröffentlicht/verifiziert; letzte geprüfteDEB weiterhinfe01e1d.


Visuelle Verknüpfungen unterstützen nun INNER/LEFT/RIGHT mit1..16 UND-verknüpften
Feldvergleichen (=, <>, <, <=, >, >=) pro neuer Quelle, auch von vorhandenen
Aliasen. Zusammengesetzte Schlüssel werden tatsächlich als ON-Vergleiche erzeugt;
RIGHT erhält die Zeilen der neuen Quelle mit echten Nullwerten links. Vorgemerkte
Vergleiche plus vollständiges aktuelles Paar bilden die Bedingung; unvollständige/
unhinzugefügte Entwürfe müssen ergänzt oder explizit verworfen werden. Visuelle
Metadaten schreibenVersion2, Version1/alte Einfachvergleiche bleiben lesbar.
Anwendung gebaut; bestehendes GUI-/SQL-Bündel2/2 (15,54s) ohne Nachlauf bestanden.
Tatsächlicher nativer executeAccessQueryPlan-Output für zusammengesetzten RIGHT
JOIN gegen umgekehrten unabhängigen LEFT JOIN inklusive Null verglichen; <=-INNER
nach Wiederöffnung ebenfalls unabhängig geprüft. Keine neuen CTests oder neue
Original-/Windows-/Gesamtsuiten-/Paketnachweise. Details VISUAL_QUERY_DESIGNER.md.
Alias-Etappe9953a83 veröffentlicht; letzte geprüfteDEB weiterhinfe01e1d; jetzt
aktuelles Paket nach dieser zusammenhängenden Entwurfserweiterung prüfen.
F067/F068/F080/F148 und fester Katalog456/386/70 unverändert.


Visuelle Aktionsabfragen implementiert: Auswahl/Tabellenerstellung/Anfügen sowie
Aktualisieren/Löschen für einzelne physische Tabellen ohne Joins/Summen/DISTINCT/
Sortierung. Expliziter freier neuer Tabellenname bzw. reale Zielspaltenzuordnung;
Updatefelder mit begrenzten skalaren SQL-Ausdrücken und deklarierten Parametern,
WHERE aus tatsächlichen Kriterien. Vorschau/Übernahme erzeugen nur SQL und
Metadaten; ausschließlich Ausführen verwendet den bestehenden nativen DAO-
QueryDef-/Bindungs-/Aktionspfad. Reale Tabellen/Zeilen/Anfüge-Standardwerte sowie
Parameteraktualisierung und kriteriumsgebundene Löschung nachgewiesen. Speichern/
Wiederöffnung behalten Modus/Ziel/Zuordnung/Ausdrücke; Formular-/Berichtsquellen
bleiben Auswahl-only. Entwurfsformat schreibtVersion3, liestVersion1/2/3.
Initiales bestehendes GUI-/SQL-Bündel2/2 (16,23s) für Make/Append bestanden. Nach
Update/Delete nur betroffenes GUI-Ziel: Qt-Iterator-Baufehler und DELETE-*-Vorschau
korrigiert; endgültig1/1 (15,79s), Anwendung gebaut. Der zunächst fälschlich dem
Update zugeordnete Fehler lag bei der DELETE-Vorbereitung. Keine neuen CTests
oder neue Kern-/Original-/Windows-/Gesamtsuiten-/Paketnachweise für diese Stufe.
Details VISUAL_QUERY_DESIGNER.md; native Laufzeit-/Typ-/Metadatengrenzen bleiben.

Zusammengesetzte Joins als44b1f70 veröffentlicht/verifiziert. Geprüftes DEB dieses
Standes /workspace/work/visual-query-composite/package/linaccess-0.7.0-Linux.deb:
Build-ID, MDBTools1.0.1/SQLite3.54.1, Payload/Berechtigungen/keine Fixtures bestanden;
dieses Paket enthält noch nicht die anschließenden Aktionsentwürfe. F067/F068/
F080/F148 und fester Katalog456/386/70 unverändert. Weiter mit nativer TOP-Auswahl
im visuellen Entwurf; kein Gesamtabschluss, kein Abschlussanruf.


Native TOP-Steuerung im visuellen Auswahlentwurf: positive32-Bit-Zahl oder Alle
Datensätze(0), tatsächlicher nativer RANK-/LIMIT-Pfad einschließlich Grenzwert-
Gleichständen und ohne sichtbare Rangspalte. Gespeicherter Wert öffnet wieder,
Bruch-/negative/überlaufende Metadaten werden abgelehnt. Rohansichten __la_top_
werden aus Quellen-/Join-/Ziellisten ausgeschlossen; öffentliche gespeicherte
Abfrage bleibt verfügbar. TOP in diesem visuellen Pfad nur Auswahl; Aktionsmodus
darf den Wert nicht stillschweigend ignorieren. Native PERCENT-/DISTINCT-/Budget-
Grenzen bleiben. EntwurfsformatVersion4 schreibt, Version1/2/3/4 liest.
Anwendung gebaut; nur bestehendes betroffenes GUI1/1 (16,14s) bestanden. Tatsächliche
TOP2-Ergebnisse mit3 Grenzwert-Zeilen gegen unabhängige nth-Wert-SQL verglichen,
Ausgabe/Metadaten/Begrenzungslöschung/Callback-Nichtauswertung geprüft. Keine neuen
CTests oder zusätzliche Kern-/Original-/Windows-/Gesamtsuiten-/Paketnachweise.
Aktionsstufe als1481810 veröffentlicht/verifiziert; letzte geprüfteDEB44b1f70
vor Aktions-/TOP-Erweiterung. Details VISUAL_QUERY_DESIGNER.md. F067/F068/F080/
F148 und fester Katalog456/386/70 unverändert; Arbeit fortsetzen.


Exakte SQL-Currency-Parameterarithmetik: tabellenfreie, vollständig bewiesene
Currency/Money-Parameter-Ausdrücke mit +,-,*, Vorzeichen und Klammern verwenden
signed64-Ticks bis zu beiden Currency-Grenzen. Berechnete Ausgaben benötigen
eindeutige explizite Aliase; Parser64 Ebenen/4096 Knoten. Native integer-only-
Callbacks verwenden currencyBinary einschließlich Zwischenwert-/Überlaufprüfung
und ties-to-even, Null bleibt Null. DAO-Type5/Size8/Value sowie Anzeige/Clipboard
behalten exakte logische Werte. Bindung regeneriert den bewiesenen SQL-Text und
verweigert auch geänderte Quellausdrücke mit gleicher Ausgabenzahl. Bestehende
REAL-Feldgrenze, Rohzahlen/Mischtypen/FROM/WHERE/Aktionsspeicherung/Filter/Find
bleiben ausdrücklich begrenzt. Keine physische Currency-Speichererweiterung.
Anwendung gebaut; bestehendes Parameter-/Arithmetik-/Ergebnis-/DAO-Bündel4/4
(2,84s) ohne Nachlauf bestanden. Grenzwerte/Rundung/Null/Überlauf/Präzedenz,
DAO-EOF-Metadaten und manipulierter Quellbeweis tatsächlich geprüft. Keine neuen
CTests oder neue Original-/Windows-/Gesamtsuiten-/Paketnachweise dieser Stufe.
Details CURRENCY_PARAMETERS_RUNTIME.md/CURRENCY_ARITHMETIC_SQL_RUNTIME.md.
TOP-Entwurf alsc068ccd veröffentlicht/verifiziert; letzte geprüfteDEB44b1f70
vor Aktions-/TOP-/exakter Parameterarithmetik. Jetzt aktuelles Paket prüfen;
F014/F067/F068 und fester Katalog456/386/70 unverändert, Entwicklung fortsetzen.


Native Debug.Print/Debug.? ist implementiert: echte VM-Auswertung, Komma-/Semikolon-
Listen, Spc/Tab, einmalige Links-nach-rechts-Auswertung und erhaltene Teilausgabe
bei späterem Fehler. Gemeinsamer flüchtiger thread-lokaler Puffer für tatsächliches
Ansichtsmenü-Dock und Debugger-Pane, auch während Stop/Continue; explizites Löschen,
Scrollposition erhalten. Grenzen1MiB/8192 Zeilen, Listen64KiB/256 Elemente/64 Ebenen,
UTF-8-sichere Aufbewahrung, UTF-16-Spalten. Kein Dateidruck/Immediate-Evaluator oder
vollständiger Windows-/Locale-Nachweis. Details NATIVE_VBA_DEBUGGER_RUNTIME.md.
Anwendung gebaut; bestehendes Kern-/GUI-Bündel2/2 (17,43s) bestanden, Kern ohneQt
verlinkt. Keine neuen CTests, kein neuer Original-/Windows-/Gesamtsuitennachweis.
Letzte veröffentlichte Vorstufe0e4f084; dazu DEB unter
/workspace/work/exact-currency-expressions/package/linaccess-0.7.0-Linux.deb mit
scripts/check-deb.sh einschließlich BuildID/SQLite3.54.1/MDBTools1.0.1 geprüft.
Dieses Paket liegt vor Debug.Print. F060 und Katalog456/386/70 unverändert.
Entwicklung ohne Zwischenabschluss fortsetzen.


Native Property Get/Let/Set mit getrennten Zugriffsmethoden ist implementiert:
unveränderte Modulquelle/Zustandsidentität, echte gleichnamige Rückgabebindung,
End/Exit Property, unqualifizierte Modulzugriffe und öffentliche Form-/Report-
Objektzugriffe einschließlich indizierter ByVal-Parameter. Setterwert immerByVal,
Objektidentität/Nothing bleiben erhalten; Erkennung führt keine Getter aus.
Signatur-/Sichtbarkeits-/Eigentümerprüfungen und tatsächliche Fehlerweitergabe;
Modulauswahl/Beobachtungen unterscheiden Get/Let/Set. Grenzen ausdrücklich in
NATIVE_VBA_PROPERTIES.md: keine vollständigen Klassen/ByRef-Indizes/ParamArray/
Default-Properties/Windows-Parität. Anwendung gebaut; bestehendes Kern-/GUI-
Bündel2/2 (16,12s) nach Korrektur der tatsächlich gefundenen Bindungsfehler bestanden.
Keine neuen CTests. Originalinventar32 VBA-Artefakte/0 Property-Deklarationen;
keine zusätzliche Originalausführung behauptet. Vorstufe Debug.Print als469ad64
veröffentlicht und Remote-Baum/privat verifiziert. Letzte geprüfteDEB0e4f084 liegt
vor Debug.Print/Properties. F042/F043/F050 und Katalog456/386/70 unverändert;
anschließend statische Prozedurlebensdauer umsetzen, ohne Zwischenabschluss.


Statische Sub-/Function-/Property-Prozeduren besitzen jetzt tatsächliche lokale
Modullebensdauer: Dim-Skalare/Objekte/feste und dynamische Arrays/feste Strings,
implizite Variant-Bindungen ohne Option Explicit, Alias-Synchronisierung bei
ByRef/Rekursion/Fehlern. Parameter, Rückgabe, Err und Konstanten bleiben proAufruf.
Initial-/Trailing-Static unterstützt, doppelte Angabe vor Effekten abgelehnt.
Def-Typ-Direktiven vor Effekten und implizite Typzeichen ausdrücklich3251 statt
falschem Variant-Erfolg; vollständige Compile-/Klassen-/Reset-Parität bleibt offen.
Details STATIC_LOCAL_RUNTIME.md/NATIVE_VBA_PROPERTIES.md. Anwendung gebaut;
bestehender Kern3,02s und GUI1/1 16,68s nach Korrektur des ungültigen null-Fenster-
besitzers im technischen Host bestanden. Spätere Typgrenzen nur im betroffenen
Kern1/1 (2,95s) geprüft; kein zusätzlicher GUI-/Original-/Windows-/Gesamtsuitenlauf.
Originalinventar32 VBA-Artefakte/0 statische Prozedurkandidaten. Keine neuen CTests.
Vorstufe Properties39e3cb2 auf privatem Branch veröffentlicht/Remote-Baum verifiziert.
F042/F043 und Katalog456/386/70 unverändert; jetzt Paket aktualisieren und danach
parameterlose Prozedur-/Property-Deklarationen ohne Klammern implementieren.
Kein Zwischenabschluss. Letzte geprüfteDEB0e4f084 liegt vor Debug.Print/Properties/
statischen Prozeduren; Original-/Fixture-Hashes7/7 zuvor unverändert nachgewiesen.


Parameterlose Sub-/Function-/Property-Get-Deklarationen ohne Klammern sind jetzt
im echten Interpreter und Modul-Inventar unterstützt, einschließlich typisierter/
Variant-Getter und Initial-/Trailing-Static. Norm MS-VBAL5.3.1; ursprüngliche
Quellidentität bleibt erhalten. Parameter ohne erforderliche Klammern, typisierte
Sub-/Setter-Rückgaben und Setter ohne Wertparameter verweigern vor Effekten.
Anwendung gebaut; erstes bestehendes Kern-/GUI-Bündel2/2 (15,90s) bestanden;
keine neuen CTests, Originalinventar32/0 Headerkandidaten, kein neuer Original-/
Windows-/Gesamtsuitennachweis. Details NATIVE_VBA_PROPERTIES.md.
Vorstufe statische Prozeduren46297bf veröffentlicht, Remote-Baum/privat verifiziert.
Dazu neueDEB /workspace/work/native-language-package/package/linaccess-0.7.0-Linux.deb
mit scripts/check-deb.sh geprüft: BuildID identisch, MDBTools1.0.1/SQLite3.54.1,
Payload/öffentliche Rechte/keine Fixtures. Paket liegt vor klammerlosen Deklarationen.
F040/F042/F043 und fester Katalog456/386/70 unverändert. Anschließend implizite
lokale Variablen auch in gewöhnlichen Prozeduren ergänzen (frisch proAufruf,
OptionExplicit und Typgrenzen beachten), gezielt prüfen und ohne Abschluss weiter.


Implizite lokale Variablen jetzt auch in gewöhnlichen nativen Prozeduren: echte
frische Empty-Variant-Bindungen proAufruf nach bestehender Namensauflösung;
statische Prozeduren behalten Modulspeicher. Unterstützte externe öffentliche
Funktionen werden aufgelöst, Sub-/Prozedurnamen nicht durch erfundene Variablen
ersetzt. Option Explicit verhindert Anlage auch beim Schreiben; native Laufzeit-
prüfung statt vollständiger Access-Compile-Diagnostik. Def-Typ-Direktiven vor
Effekten und implizite Typzeichen verweigert;4096 implizite Namen proRahmen.
Details VBA_IMPLICIT_LOCALS.md. Anwendung gebaut; GUI16,32s bestanden, nur zuvor
betroffene Kern-/Original-Ziele nach technischen Fixturekorrekturen erneut geprüft:
2/2, Kern3,11s und unveränderter real_kassa_dao_event124,12s. Keine neuen CTests,
kein GUI-Nachlauf/Windows-/Gesamtsuitennachweis. Originalinventar32/19OptionExplicit/
13ohne/0Def; unveränderter Ereignislauf bleibt an2485/1023 fehlender Makroquelle,
kein vollständiger Verkauf behauptet. Vorstufe klammerlose Deklarationen b05c240
veröffentlicht/privater Remote-Baum verifiziert; letzte geprüfteDEB46297bf vor
klammerlosen Deklarationen/impliziten normalen Variablen. F037/F042/F043 und
Katalog456/386/70 unverändert. Anschließend objektwertige Getter/default-Member-
Weiterleitung nach MS-VBAL5.4.3.8/5.6.13 ergänzen; ohne Zwischenabschluss weiter.


Objektwertige Getter unterstützen jetzt native Default-Member-Weiterleitung:
parameterlose Property-/samelokale Function-Ergebnisse mit positionalen Indizes,
Get-only Objektproperties für native Dictionary-Let/Set und Default-Wertzuweisung
an tatsächliche Textfelder, auch über indizierten Getter. Echte Setter haben Vorrang;
Getter/Argumente/RHS nicht doppelt ausgeführt, Objektidentität erhalten, native
Fehler13/91/Getterfehler weitergegeben. Eigentümerschließen im Getter verhindert
Default-Zuweisung mit2450. Grenzen/default-Queues/Arrays/COM/ByRef-Fassade bleiben
explizit, Norm MS-VBAL5.4.3.8/5.6.13, Details NATIVE_VBA_PROPERTIES.md.
Anwendung gebaut; erstes Laufzeit-Kern-/GUI-Bündel2/2 (17,19s) bestanden. Keine
neuen CTests. Originalinventar32/0 Properties/0 parameterlose Objektfunktionen;
kein neuer Original-/Windows-/Gesamtsuitennachweis. Implizite normale Variablen
als3c61be9 privat veröffentlicht/Remote-Baum verifiziert; Original-/Fixture-Hashes
7/7 unverändert geprüft. Letzte geprüfteDEB46297bf liegt vor klammerlosen
Deklarationen/normalen impliziten Variablen/Default-Weiterleitung. Katalog456/386/70
und F042/F043/F050 unverändert. Anschließend ByRef-Property-Indizes für direkte
samelokale Get/Let/Set-Aufrufe mit echter Rückbindung ergänzen; externe Objekt-
fassade weiterhin explizit begrenzen. Ohne Zwischenabschluss fortsetzen.


ByRef-Indizes bei direkten samelokalen Property-Get/Let/Set-Aufrufen sind echt
rückgebunden: vorhandener Prozedurargument-Binder statt abgetrennten Wertkopien,
einschließlich Array-Elementen, temporären Ausdrücken und Fehlerseitenwirkungen.
Setterwert bleibtByVal, Objektidentität/Typ-/Array-Allokationsguards bleiben.
Auch Default-Wertzuweisung über indizierten Objektgetter nutzt diese Rückbindung;
realer Textfeldnachweis. Externe Objektfassade bleibt bei ByRef-Indizes explizit3251,
keine erfundene Rückbindung. Details NATIVE_VBA_PROPERTIES.md. Anwendung gebaut;
erstes bestehendes Kern-/GUI-Bündel2/2 (17,35s) bestanden, keine neuen CTests oder
Original-/Windows-/Gesamtsuitennachweise. Vorstufe Default-Weiterleitung e588edc
privat veröffentlicht/Remote-Baum verifiziert. Letzte geprüfteDEB46297bf vor
klammerlosen Deklarationen/normaler impliziter Bindung/Default-/ByRef-Erweiterung.
F042/F043/F050 und Katalog456/386/70 unverändert. Nächste Stufe: optionale und
ParamArray-Indizes in Properties einschließlich besonderem letzten ByVal-Wert-
parameter korrekt binden; danach Paket aktualisieren. Ohne Zwischenabschluss weiter.


Optionale und ParamArray-Property-Indizes sind implementiert: letzter erforderlicher
Setterwert getrennt von Indexliste, fehlende Positionen vor dem Wert ergänzen,
API-Argumentvektorlänge und ByRef-Indexrückbindung erhalten. Unterschiedliche
Get-/Let-Defaults nach MS-VBAL5.3.1.5/5.3.1.7; Cell(,8)/Cell() und benannte optionale
Getter einschließlich echter Formularfassade. ParamArray enthält nur Indizes,
RHS stets actual-last/ByVal, auch bei leerem Array oder zusätzlichen Argumenten;
Objektidentität erhalten. Variadische Signaturen verhindern erfundene benannte
ParamArray-Bindung; Mutationen/Weiterleitung/Holes explizit begrenzt. Ausgelassene
Arrayindizes werden vor Bodyeffekten verweigert, nicht0 geraten. Details
NATIVE_VBA_PROPERTIES.md. Anwendung gebaut; erstes bestehendes Kern-/GUI-Bündel
2/2 (16,60s) bestanden. Keine neuen CTests, kein zusätzlicher Original-/Windows-/
Gesamtsuitennachweis. ByRef-Vorstufe bc52b22 privat veröffentlicht/Baum verifiziert.
Letzte geprüfteDEB46297bf liegt vor mehreren Sprachstufen; jetzt Paket aktualisieren.
Katalog456/386/70 und F042/F043/F050 unverändert; anschließend typisierte Array-
Funktions-/Getterrückgaben und echte Leseindizierung ergänzen. Ohne Abschluss weiter.


Typisierte Function-/Property-Get-Arrayrückgaben sind implementiert: geschlossene
native Elementtypen, vorhandene Array-Typ-/Rang-/Grenz-/Wertkopiesemantik und echte
Leseindizierung mit einmaliger Erzeugerausführung. Auch Variant-Arrayergebnisse,
mehrdimensionale Getter und separate Indexqueue nach parametrisiertem Aufruf.
Bekannte typisierte Array-SQL-Funktionen vor Bodyeffekten3251; keine erfundenen
SQLite-Skalarwerte. Reale Formularfassade und gemeinsamer SQL/Modul-Zustandszähler
beweisen Lesezugriff bzw. fehlende SQL-Seiteneffekte. Details NATIVE_VBA_PROPERTIES.md.
Anwendung gebaut; erstes bestehendes Kern-/GUI-Bündel2/2 (16,91s) bestanden, keine
neuen CTests/Original-/Windows-/Gesamtsuitennachweise. Vorstufe1313196 privat
veröffentlicht und frischeDEB1313196 samt Build-ID/Payload/Permissions und gepinnten
SQLite/MDBTools geprüft. Aktuelle Array-Erweiterung liegt nach diesem Paket.
Katalog456/386/70 unverändert; Array-Ergebniszuweisungen/Klassen-/UDT-/COM-Typen
bleiben begrenzt. Direkt weiter: For Each über tatsächliche Array-Funktions-/Getter-
ergebnisse mit einmaliger Ausdrucksauswertung und erhaltenen Named-Root-Guards.


For Each akzeptiert echte Array-Ausdrücke einschließlich typisierter/Variant-
Funktions- und Getterergebnisse. Quelle nach MS-VBAL5.4.2.4 genau einmal auswerten,
Ergebniswert während Schleife halten; bisherige intrinsische Namensliste entfällt.
Benannte Wurzeln behalten vorhandene Live-Elementänderungen und Austausch-/Resize-
Guards. Erstes bestehendes Kern-/GUI-Bündel2/2 (17,83s) und Anwendung erfolgreich:
Erzeugerzähler, Matrix-Reihenfolge, Objektinhalte, ByRef-Input/Wertkopieunabhängigkeit,
verschachtelte Schleifen/Exit For. Details VARIANT_ARRAY_RUNTIME.md. Keine neuen
CTests/Original-/Windows-/Gesamtsuiten-/Paketnachweise; letzte geprüfteDEB1313196.
Arrayrückgabe-Vorstufe2ce6323 privat veröffentlicht/Remote-Baum verifiziert.
Katalog456/386/70 unverändert. Direkt weiter: native Def-Typdirektiven statt der
bisherigen pauschalen Verweigerung, mit konsistenter Typauflösung und Grenzen.


Native DefBool/Byte/Cur/Date/Dbl/Int/Lng/Obj/Sng/Str/Var-Typdirektiven sind integriert:
aktive Moduldeklarationen, ASCII-Einzelbuchstaben/auf- und absteigende Bereiche,
A-Z als universelle Vorgabe einschließlich Unicode-Identifier. Dieselbe Auflösung
für Modul-/Lokalskalare und Arrays ohne As, Parameter, Function/Get-Rückgaben,
Setterwerte und gewöhnliche/statische implizite Locals; explizites As gewinnt,
Const bleibt aus Ausdruck typisiert. Option Explicit bleibt wirksam. Fehlerhafte/
überlappende/64-Bit-/prozedurinterne Direktiven vor Bodyeffekten abweisen. Implizite
Currency-For-Zähler exakt und neue ReDim-Ziele bei Def-Vorgabe typisiert; keine
Behauptung vollständigen Deklarations-Hoistings oder Suffix-/Klassenparität.
Details VBA_DEF_TYPES.md. Erste GUI-Prüfung erfolgreich, erster Kernblock fand
leere DefInt-Spezifikation; genau diesen Parserfehler behoben und fehlende Schleifen-/
ReDim-Bindung ergänzt. Final bestehendes Kern-/GUI-Bündel2/2 (17,68s; Kern4,33s),
Anwendung gebaut; alle7 Upload-/Fixturehashes unverändert. Keine neuen CTests,
Original-/Windows-/Gesamtsuitennachweise. Vorstufe a6b2429 privat veröffentlicht und
Remote-Baum verifiziert. Letzte geprüfteDEB1313196; Paket jetzt aktualisieren.
Katalog456/386/70 unverändert; ohne Zwischenabschluss weitere offene Umfänge umsetzen.


Echtes Tabellen-Datenblatt: Dialogbereich Alle sichtbaren Textfelder und atomare
Mehrfeld-Alle-ersetzen-Aktion implementiert. Vorbereitete Zellen über mehrere
Spalten/einen SQL-Transaktionsblock; späterer Feld-/Datensatzfehler rollt frühere
Schreibzugriffe zurück und entfernt nur eigene Ersetzungsentwürfe. Primärschlüssel,
generierte/ausgeblendete/numerische/Binärfelder ausschließen, Nullwerte erhalten,
generierte Werte normal ableiten. Aktuell-Ersetzen bleibt Entwurf am aktuellen
Textfeld und behält Allfeld-Suchbereich. Validierungszeitliche Quell-/Feld-/Sicht-
Änderungen vor Schreiben2501; sortierten aktuellen Schlüssel/Spalte wiederherstellen.
Details LITERAL_REPLACE_RUNTIME.md. Anwendung gebaut, erstes bestehendes GUI-/Modell-
Bündel2/2 (16,60s): reale Buttons/SQL, vier Feldwerte, später CHECK-Rollback, fremder
Entwurf, Scopewechsel und doppelte Spalten geprüft. Keine neuen CTests/Original-/
Windows-/Gesamtsuitennachweise. Def-Typ-Vorstufe cd0cd59 privat veröffentlicht/Remote-
Baum verifiziert; frischeDEBcd0cd59 samt Payload/Rechten/Build-ID und gepinnten
SQLite/MDBTools erfolgreich geprüft. Dieses Paket liegt vor Mehrfeld-Ersetzen.
Katalog456/386/70 unverändert; ohne Zwischenabschluss weitere offene Umfänge umsetzen.


Bearbeitbare vollständige Einzeltabellen-Abfragen erweitert: SELECT */alias.*,
explizite/implizite Aliasnamen, qualifizierte WHERE-Felder und bis10 unterschiedliche
physische ORDER-BY-Felder mit ASC/DESC. Nur geprüfte Quelltabelle/-felder; Alias-
Filterqualifikation kanonisch umsetzen, Literale erhalten. Nativer Tabellenmodell-
Sortierpfad verwendet geprüfte Index-/Richtungsmetadaten, escaped SQL und behält
sekundäre Reihenfolge bei Select/Save/Requery. Kopf zeigt erstes Feld, andere
Benutzersortierung ersetzt Quellsortierung; vorhandene Draft-/Editor-Guards bleiben.
Echte gespeicherte Abfrage schreibt Basisdatensätze, erhält Schlüssel bei Sortier-
bewegung und entfernt nicht mehr passende Zeilen aus Filteransicht. Keine pauschale
Join-/Projektions-/Parameter-/Nullschlüssel-/Shadow-Updatebarkeit. Details
OPEN_QUERY_RUNTIME.md. Anfangs bestehendes GUI/Modell2/2 (16,82s); anschließend
Mehrfeldsortierung und echte neue GUI-/Modellfälle komplettiert, fehlendes Modell-
Test-Include korrigiert und nur dessen Ziel neu gebaut. FinalGUI/Modell2/2 (16,59s,
Modell0,09s), Anwendung gebaut. Keine neuen CTests/Original-/Windows-/Gesamtsuiten-/
Paketnachweise. Mehrfeld-Ersetzungs-Vorstufe1c66d81 privat veröffentlicht/Baum
verifiziert; letzte geprüfteDEBcd0cd59 liegt vor beiden Datenblatt-Erweiterungen.
Katalog456/386/70 unverändert; F079/F373-Beschreibungen bestehenden Umfängen
zugeordnet. Ohne Zwischenabschluss weiter, danach Paket aktualisieren.


Native eigene Klassenmodule und echte Instanzen implementiert: Klasseneditor/
Explorer/kind=class, gespeicherte CLASS/VB_Name-Quelle, gemeinsamer kollisionsgeprüfter
Modulkatalog/AllModules mit echter Editor-IsLoaded-Erkennung, keine Registrierung
von Klassenmethoden als freie SQL-Funktionen. New erzeugt getrennte Feld-/Array-/
Static-Zustände und führt privaten parameterlosen Class_Initialize einmal aus.
Öffentliche Felder/Konstanten/Sub/Function/Get/Let/Set, Me-/Objektidentität,
Arrayrückgaben, native Default-Weiterleitung und echte private Aufruferscopes.
Reentrant fremde Prozeduren/Ausdruckshosts erben keine privaten Klassenrechte.
Kontextdienste verwenden echte Projekt-/DAO-Hosts ohne erfundenes Formular;
Instanzen überleben aufrufende Editor-/Formfenster im lebenden Projekt. Beobachtete
Quellrevision-/Lösch-/Projektänderungen invalidieren dauerhaft. Grenzen ausdrücklich:
Object/Variant-Referenzen statt benannter Custom-As-Typen, externe Argumente ByVal,
kein importierter generischer Klassenbau/As New/Implements/WithEvents/Default-
Instanzen/COM; Class_Terminate vor Konstruktion verweigert. Details NATIVE_VBA_CLASSES.md.
Erster Kern/DAO bestanden; Original fand Def-Vorprüfung gewöhnlicher Kontrollnamen,
GUI erreichte20s-Watchdog im parallelen Block. Def-Bindung auf echte typisierte Ziele
und anwendbare Def-Namen begrenzt, Fehler kontrolliert transportieren; New-Prüfung
verwechselt rohe Kontrollnamen nicht mit Konstruktoren. Private Reentry-Scope-Lücke
vor Veröffentlichung geschlossen. Final Kern/GUI/Original3/3 (143,28s:4,08/17,58/
121,62s), DAO vorher3,49s. Anwendung gebaut, KernQt-frei, alle7 Original-/Uploadhashes
unverändert. Keine neuen CTests/Windows-/Gesamtsuitennachweise; Original bleibt am
bekannten fehlenden Makroquelltext statt positiver Verkaufsausführung begrenzt.
Vorstufe3fe7245 privat veröffentlicht/Remote-Baum verifiziert; frischeDEB3fe7245 mit
beiden Datenblatt-Erweiterungen vollständig geprüft, liegt vor Klassen-Erweiterung.
Katalog456/386/70 unverändert; ohne Zwischenabschluss weitere offene Umfänge umsetzen.


Eindeutig gewöhnliche importierte Klassen können jetzt mit New aus der neuesten
wirklichen VBA-Archivrevision konstruiert werden. CLASS/VB_Name und kanonischer
Archivname müssen passen; VB_Base und tatsächlich zugeordnete Form_/Report_-Module
bleiben Dokument-CodeBehind. Unbelegter Präfix allein wird nicht zum Dokumenttyp
geraten. AllModules enthält nachgewiesene gewöhnliche Importklassen. UTF8/BOM,
UTF16LE und vorhandener CP1252-Ausführungsview ohne Änderung der Rohbytes;
verlustbehaftete undefinierte Fallbackbytes3251. Herkunft, Archiv-ID und rohe
Revision dauerhaft gebunden, auch semantisch gleiche anders kodierte Neuversion
invalidiert alte Instanz. Namespacekollisionen, TEMP-/Lösch-/Kontextguards bleiben.
Details NATIVE_VBA_CLASSES.md. Erstes bestehendes GUI/DAO-Bündel2/2 (20,44s:
17,74/2,70s), App gebaut: tatsächliche Konstruktion/Methoden/SQL, Kodierung/
Versionen, Metadaten, Dokumentausschlüsse und Originalarchivhash auf Testkopie.
Keine neuen CTests/Original-/Windows-/Gesamtsuitennachweise. Original-Bridge nur
aggregiert geprüft:28 Klassendateien/28 deklarierte Dokumentmodulnamen/0 gewöhnliche
Kandidaten; technische Quellen sind kein Windows-Exportoracle. Klassenkern bleibt
auf voriger geprüfter Umsetzung. Vorstufe7f40893 privat veröffentlicht/Baum geprüft,
frischeDEB7f40893 samt Payload/Rechten/Build-ID/gepinnten SQLite/MDBTools bestanden.
Katalog456/386/70 unverändert. Ohne Zwischenabschluss mit benannten Klassen-
Typbindungen und weiteren offenen Anforderungen fortfahren.


Benannte Projektklassen-Typbindungen: eigene/gewöhnliche archivierte echte Klassen
in skalaren Dim/Feldern/Parametern/Rückgaben/GetSet/TypeOf; optionalNothing und
owning-ByRef, Typfehler13, unbekannte Bindung3251 vor Body. Selbsttypen ohne
rekursive Instanziierung; direkte Formularfelder und SQL-Klassenresultatgrenze
geprüft. Details NATIVE_VBA_CLASSES.md. Klassenarrays/AsNew/Interfaces/externes
ByRef bleiben offen. Katalog456/386/70 unverändert; keine Paritätsbehauptung.

Nachweise dieses Blocks: native Kernprüfung nach If-Typbindungsfix1/1 (4,15s),
mit OptionExplicit-Fixture erneut1/1 (4,04s); Original-Ereignispfad1/1 (123,33s)
für die native VM, anschließend nur noch Qt-Formularfeld-Set ergänzt. Finales
vorhandenes GUI-Ziel1/1 (17,80s), App gebaut. Erster GUI-Basisblock17,88s bestand;
Erweiterungen korrigierten technische SQL-Registrierung/Deklarationsreihenfolge
und fanden den fehlenden echten Formularfeld-Set-Pfad. Betroffene Ziele gezielt
wiederholt; keine neue CTest-ID oder Vollsuite. Sieben Upload-/Fixturehashes
unverändert. Kein neuer Windows-/Distributions-/DEB-Nachweis; letzter geprüfter
DEB-Stand7f40893 liegt vor archivierten und typisierten Klassenerweiterungen.


## Benannte Projektklassen-Arrays (2026-11-09)

Feste/dynamische lokale und Modul-/Klassenfelder, ReDim/Preserve, geprüfte Set-
Elementzuweisungen, gleiche ByRef-Arrayparameter und Function-/Get-Rückgaben
verwenden jetzt die reale gebundene Projektklasse als Elementtyp. Initialisierte
und durch Erase zurückgesetzte Plätze enthalten Nothing; falsche Objekt- oder
Arraytypen scheitern13. Kopien behalten getrennte Arraywerte und gemeinsame reale
Objektreferenzen. Klassenname/Enum-Grenzen werden einschließlich expliziter ReDim-
Typangaben vor Effekten gebunden. Ein Array wird nicht als skalares Klassenobjekt
behandelt. VarType ist8201; VBA-TypeName löst den kanonischen tatsächlichen Namen
im Projekt auf, auch für undimensionierte Rückgaben. Debugger-/C++-Typanzeigen
behalten die normalisierte Deklarationsschreibweise, ohne interne Typpräfixe.

For Each behält die normativ erforderliche Variant-Schleifenvariable für Arrays;
die Klassentypisierung der Elemente hebt diese Regel nicht auf (MS-VBAL5.4.2.4).
Reale Rückgabearrays werden einmal enumeriert. Bestehende Rang-/Bounds-/Preserve-/
Speicher-/ByRef-Element- und Mutationsgrenzen bleiben erhalten. Keine allgemeinen
Collection-/DAO-Typarrays, AsNew, Interfaces oder vollständige Referenzparität.
Dies ersetzt die vorherige Klassenarray-Grenze für den beschriebenen Umfang.

App gebaut; vorhandenes Kern-/GUI-Bündel nach Fixture-/Binderkorrektur2/2 bestanden
(22,53s:4,01/18,52s). Erste Prüfung fand die falsche ForEach-Fixture-Deklaration;
danach war der skalare Objektbindungszweig für ganze Klassenarrays zu korrigieren.
Nur betroffene bestehende Ziele wiederholt, keine neue CTest-ID/Original-/Windows-
oder Vollsuite. Der GUI-Archivtest umfasst gewöhnliche importierte Klassen mit
solchem Arrayfeld und erhält den ursprünglichen Archivhash. Kein neuer DEB-Nachweis.
F042/F043 bleiben offen; Katalog456/386/70 unverändert. Ohne Zwischenabschluss an
nachgewiesenen benutzerdefinierten Klassen-Standardmitgliedern und weiteren offenen
Anforderungen weiterarbeiten.


## Deklarierte Klassen-Standardmitglieder (2026-11-09)

Ein tatsächlich vorhandenes Memberattribut VB_UserMemId=0 bindet jetzt einen
öffentlichen Property Get oder eine öffentliche Function als Klassenstandard.
Die ursprünglichen Attributzeilen/Zeilennummern bleiben erhalten; der Parser
überspringt nur den passenden geschlossenen0-Attributsatz der gewählten Prozedur.
Unbekannte/private/nicht lesbare Standardmitglieder, widersprüchliche Namen und
nicht implementierte andere Member-IDs/Enumeratoren werden vor Konstruktion
abgewiesen. Ein unannotiertes Item wird nicht zum Klassenstandard erfunden.
Ohne deklarierten Standard schlägt ein direkter Standardzugriff438 fehl.

Implizite skalare Auswertung verwendet nur einen Standardleser, der ohne
Pflichtargument aufrufbar ist. Getterfehler und die vorhandene16-stufige Grenze
rekursiver Standardwerte bleiben erhalten. Explizite Positions-/Namensargumente,
Missing/optionale Indizes und vorhandene Property-Schreibbinder werden verwendet;
Standardargumente bekommen die echte Signatur, unabhängig von einem anderen
öffentlichen Item-Member. Namensfehler448 werden vor Getteraufruf erkannt.
Standard-Schreiben verwendet tatsächliche Let/Set-Accessoren beziehungsweise
vorhandene geprüfte Rückgabeweiterleitung. Benannte Setterindizes, Klassen-Bang-
Syntax, vollständiges externes ByRef, allgemeine COM-Attribute/-IDs und
Enumeratoren bleiben offen; objektwertige Standard-Set-Kombinationen sind kein
separater Windows-Nachweis. Diese Umsetzung ersetzt die frühere pauschale
Standardmitglied-Attributgrenze nur für den beschriebenen0-Pfad.

App gebaut, erstes Laufzeitbündel der vorhandenen Kern-/GUI-Ziele2/2 bestanden
(22,36s:4,07/18,29s). Gemeinsame technische Quellen prüfen Default-Get/Let,
impliziten Skalarwert, optionale/namentliche Indizes, unverwechselbares Item,
private Zustände und widersprüchliche/andereIDs. Der echte Projektpfad sowie eine
gewöhnliche archivierte Defaultklasse und die unveränderte ursprüngliche
Archivsignatur sind geprüft. Keine neue CTest-ID/Original-/Windows-/Vollsuite.
Katalog456/386/70 unverändert, F042/F043 bleiben offen. Nach privater Veröffentlichung
aktuelles Paket prüfen und ohne Zwischenabschluss weitere Anforderungen bearbeiten.


## Skalare automatische Klassenvariablen As New (2026-11-09)

AsNew ist nun für skalare lokale/statische Variablen, Modul-/Klassenfelder und
öffentliche Formularfelder gebunden: gewöhnliche tatsächliche Projektklassen
sowie die schon vorhandenen nativen Collection/Dictionary/FileSystemObject-
Konstruktoren. Prüfung bindet die reale konstruierbare Klasse vor Effekten,
führt aber keine Initialisierung aus. Erst tatsächliches Lesen erzeugt eine
Instanz; Set auf Nothing erzeugt keine Zwischeninstanz, späteres Lesen eine neue.
Normbasis MS-VBAL5.2.3.1.4/5.3.1.6. Eigene Parameter bleiben normale ByRef-Parameter;
eine Rücksetzung durch sie erhält die automatische Eigenschaft des Aufrufers.

Statische/Modulzustände und gewöhnliche Felder behalten tatsächliche Referenzen.
Direkter Me-/öffentlicher Feldzugriff und lokale Aliasansichten werden bei AsNew
synchronisiert, ohne ausstehenden lokalen Set-Wert zu überschreiben. Metadaten-
und Debuggeranzeigen lesen den Zustand ohne Erzeugung. Skalare Standardzuweisung
verwendet den realen Default-Setter. Direkte Formularzugriffe halten den Modulframe
bei Initialisierungs-Reentry fest und prüfen Schließen; Zielobjekte werden während
Default-Aufrufen gehalten. Arrays mit AsNew/automatischen Elementen, WithEvents,
Defaultinstanzen und komplette Terminierung/GC bleiben offen, keine allgemeine
AsNew- oder Vollklassen-Parität behauptet. Unbekannte/nicht konstruierbare Klassen
werden nicht als leere Scheinobjekte angelegt.

App gebaut; erstes bestehendes Kern-/GUI-Bündel2/2 (22,90s:4,40/18,49s). Nach
abschließendem Aliasabgleich ergänzt und geprüft:2/2 (22,46s:4,25/18,20s).
Gemeinsame technische Quellen zählen tatsächliche Initialisierungen, prüfen
verzögerten ersten Zugriff/SetNothing/Identität, ByRef-Reset, Static-/Modul-Erhalt,
Native-Collection/Dictionary-Neuanlage, echte Default-Zuweisung, Me-/unqualifizierte
Feldsicht und echte öffentliche Formularfelder. Kein neues CTest/Original-/Windows-
oder Gesamtsuite-Orakel. Katalog456/386/70 unverändert, vollständige F042/F043 offen.

Vorstufe dced5e6 privat remote/Baum geprüft; aktuelles DEB dieses Stands unter
/workspace/work/class-default-members/package/linaccess-0.7.0-Linux.deb besteht
Payload/Build-ID/Rechte/keineFixtures, MDBTools1.0.1 und SQLite3.54.1 samt
NestedJOIN/NULL-/Counter-Prüfung; bestehendes DAO-Ziel1/1 (2,65s). Dieses Paket
enthält die zuvor geprüften Klassen-/Array-/Default-Stufen, noch nicht AsNew.
Ohne Zwischenabschluss mit belegten weiteren Deklarations-/Laufzeitbindungen
fortfahren; keine pauschale Referenzanforderung oder Etappen-Abschlussanruf.


Enum-Typbindungen (2026-11-09): tatsächliche quelllokale Enum-Namen werden gemäß
MS-VBAL5.2.3.4 in As-Deklarationen als Long gebunden, einschließlich Parametern/
Rückgaben, Feldern/Static, Arrays, GetLet und typisierten Konstanten. Einzeilige
If-Pläne erben den echten Deklarations-/Options-/Konstantenkontext; Cache vermischt
keine fremden Typbindungen. Details VBA_ENUM_TYPE_BINDINGS.md. Bestehendes Kern-/
GUI-Bündel nach Fixturekorrektur2/2 (23,08s:4,60/18,48), App; Original-Ereignispfad
1/1 (121,37s), bekannte Makrogrenze beibehalten. Sieben Hashes unverändert. Kein
neues CTest/Windows-/Vollsuite-/DEB-Orakel. Crossmodule-/qualifizierte Typreferenzen,
UDTs und vollständige Unicode-/Compilersemantik offen;456/386/70 unverändert.
Vorstufe455abfa privat remote/Baum geprüft. Ohne Zwischenabschluss mit echten
editierbaren einzelnen Tabellenfeldprojektionen in Abfragedatenblättern fortfahren.


Editierbare einzelne Abfragefelder: EDITABLE_QUERY_PROJECTIONS.md beschreibt die
geprüfte physische Feldzuordnung mit Alias/Subset/Reihenfolge, verborgenem echtem
Schlüssel, SQL-Speichern unveränderter ausgelassener Felder, Sortierung nach
Ausgabealias sowie Layout-/Header-Erhalt nach Requery. Nur angeforderte Projektion
wird erweitert; andere Stern-Caller bleiben unverändert. App/GUI1/1 (18,42s),
nach Aliasnormalisierung/Unicode-Namensschutz nur betroffenes Modellziel1/1 (0,11s).
Keine neuen CTests/Original-/Windows-/Vollsuite-/Paketbehauptungen. Mehrspaltige
umgeordnete Zwischenablage zunächst geschützt3251; als nächste zusammenhängende
Etappe echte sichtbare Spaltenabbildung für Kopieren/Ausschneiden/Einfügen/Undo
implementieren. F079/F340/F373/456/386/70 bleiben offen/unverändert. Vorstufe03b3de2
privat veröffentlicht, Remote/Baum geprüft; ohne Zwischenabschluss weiterarbeiten.


Sichtbare Zwischenablage-Spaltenabbildung: PROJECTED_GRID_CLIPBOARD.md ersetzt die
vorherige mehrspaltige Projektionsverweigerung. Kopieren/Ausschneiden/Einfügen und
Undo/Redo verwenden dieselben geprüften realen Spalten, verborgene Schlüssel/
ausgelassene Werte bleiben geschützt; mapped Append bewahrt tatsächliche SQL-
Defaults. Vorhandenes GUI/Currency-Feld/Modell-Bündel im ersten Lauf3/3 (18,69s:
18,11/0,48/0,11), App gebaut. Reale SQL-/Undo-/Cut-/MIME-/Negativfälle, keine neuen
CTest-/Original-/Windows-/Vollsuiten-/Paketbehauptungen.456/386/70 unverändert;
vollständige F079/F340 bleiben offen. Vorstufe2d95062 privat remote/Baum geprüft.
Ohne Zwischenabschluss den tatsächlichen Makroreferenz-Blocker über die bestehende
Codex-Call-Schnittstelle prüfen und unabhängig weitere Anforderungen implementieren.


Native Objektarrays (2026-11-09): NATIVE_OBJECT_ARRAYS.md beschreibt die geschlossene
DAO-/Access-/Scripting-/VBA-Typbindung für Deklarationen, ByRef, Rückgaben, Preserve
und Set-Elemente mit echter Schnittstellenprüfung statt Klassenname. Reale lokale
Forms-Arraybindung hat Vorrang vor globaler Sammlung. Nach gezielter Korrektur
bestehendes Kern-/GUI-Bündel2/2 (24,46s:4,85/19,61), DAO zuvor1/1 (2,83s), App.
Keine neuen CTests/Original-/Windows-/Vollsuiten-/Paketbehauptungen.456/386/70
unverändert; vollständige F042/F043 offen. Vorstufe9bc07b2 privat remote/Baum geprüft.
Gespeicherten Makro-Codex-Call einmal ohne neuen Anruf fortgesetzt: failed, keine
Antwort/Entscheidung; Ende vor Agent-Abschlussverarbeitung. Makroreferenz bleibt
offen; unabhängig ohne Zwischenabschluss weiterprogrammieren.


ReDim-Deklarationswirkung: REDIM_DECLARATIONS.md beschreibt echte prozedurweite
Arraydeklaration bei ungebundenem Namen (auch OptionExplicit/übersprungene Zweige),
Def-/Elementtyp, ByRef/Rückgaben/Static und Schutz bestehender Modulbindungen.
App; bestehendes Kern-/GUI-Bündel2/2 (22,96s:4,74/18,22). Gemeinsame Integration
mit nativen Objektarrays: DAO2,71s + Original-Ereignispfad119,91s,2/2 in122,62s.
Bekannte Makrogrenze beibehalten; sieben Hashes unverändert. Keine neuen CTests/
Windows-/Vollsuiten-/Paketbehauptungen;456/386/70 unverändert. Vorstufe58b1723
privat remote/Baum geprüft. Ohne Zwischenabschluss mit belegten Err-Reset-
Grenzen nach MS-VBAL6.1.3.2.1.1 weiterprogrammieren.


Automatisches Err.Clear: VBA_ERROR_CLEAR_BOUNDARIES.md dokumentiert echte aktive
OnError-/ExitSub-/ExitFunction-/ExitProperty-Grenzen nach MS-VBAL6.1.3.2.1.1,
alle fünf Eigenschaften, übersprungene Anweisungen und zurückbehaltene ErrObject-
Referenzen. App; bestehendes Kern-/GUI-Bündel im ersten Lauf2/2 (23,07s:4,76/18,30).
Keine neue CTest-ID/Windows-/Vollsuitenbehauptung. Vollständige umgebungsweite
Err-Identität und Fehlerhistorie bleiben offen;456/386/70 unverändert. Vorstufe
113b829 privat remote/Baum geprüft; Original zuletzt119,91s und sieben Hashes
unverändert vor dieser begrenzten Reset-Korrektur. Ohne Zwischenabschluss ein
aktuelles Debian-Paket erstellen und weitere echte Laufzeitlücken bearbeiten.


Projektweite native Err-Instanz: VBA_ERROR_ENVIRONMENT.md beschreibt echte
Host-/Projektbindung statt pro-Prozedur-Neuanlage, gemeinsame Referenzen über
Modul/Formular/Klasse, lokale Handler, tatsächliche Fehlerweitergabe und getrennte
Datenbanken. Reset invalidiert alte Referenzen; geschlossene Verbindung verweigert
Eigenschaftszugriff3420. App; initial GUI18,42s/DAO2,75s/Original118,10s bestanden.
Eine alte Kernfixture benötigt für ihren Null-Start ausdrücklich Err.Clear; nur
Kern wiederholt, bestanden5,03s. Sieben Hashes unverändert; bekannte Makrogrenze.
Keine neue CTest-ID/Windows-/Vollsuitenbehauptung;456/386/70 unverändert. Vorstufe
55096a7 privat remote/Baum geprüft. Dessen aktuelles DEB unter
/workspace/work/error-clear-boundaries/package/linaccess-0.7.0-Linux.deb geprüft
(Build-ID/Payload/Rechte/keineFixtures, MDBTools1.0.1, SQLite3.54.1 JOIN/NULL/Counter),
noch ohne diese Projektbindung. Ohne Zwischenabschluss mit echter CVErr-
Fehlerwertübernahme/Let-Konvertierung nach MS-VBAL6.1.2.3.1.14 weiterprogrammieren.


CVErr-Fehlerwertübernahme: CVERR_CONVERSION.md beschreibt unveränderte echte
Error-Argumente und sonstige Long-Let-Konvertierung, typisierte Rundung und
Grenzfehler. App; bestehendes Kern-/GUI-Bündel im ersten Lauf2/2 (22,83s:4,65/18,17).
Echte gespeicherte skalare SQL-VBA-Funktionen lesen denselben Projekt-Err.Number907
und liefern CVErr(Error2015) als Text, ohne einen neuen Fehler auszulösen. SQL-
Objekt-/Array-Schutz bleibt. Keine neue CTest-ID/Original-/Windows-/Vollsuiten-/
Paketbehauptung;456/386/70 unverändert. Vorstufe725b4bd privat remote/Baum geprüft.
Aktuell geprüftes DEB55096a7 liegt vor Projekt-Err/CVErr. Ohne Zwischenabschluss
mit nativen Error/Error$-Funktionen und getrennter zuletzt-ausgelöster Fehlernummer
nach MS-VBAL6.1.2.3.1.15 weiterprogrammieren.


Native Error/Error$: VBA_ERROR_TEXT.md beschreibt optionale/benannte/qualifizierte
und nackte Aufrufe, Long-Konvertierung, begrenzte invariante englische Meldungen
und getrennte zuletzt-ausgelöste Fehlernummer. Err.Number-Zuweisung und Clear/
OnError/Exit ersetzen diese Historie nicht. Echte SQL-VBA-Funktionen unterscheiden
Err.Number907 von zuvor ausgelöstem91. Projektfunktion/lokales Array Error haben
unqualifiziert Vorrang; VBA.Error bleibt Intrinsik, Access-Feld Error bleibt Feld.
App; initial bestehendes Kern-/GUI-Bündel2/2 (23,15s:4,81/18,33). Nach Feldnamens-
Schutz nur betroffenes Kernziel wiederholt1/1 (4,87s). Keine neue CTest-ID/Windows-/
Vollsuitenbehauptung;456/386/70 unverändert. Vorstufe2d3e8db privat remote/Baum
geprüft. Letztes Original118,10s/Hashes vor CVErr/Errortext; DEB55096a7 vor Projekt-
Err/CVErr/Errortext. Ohne Zwischenabschluss aktuelles Paket prüfen, danach echte
ByRef-Arraysperren nach MS-VBAL5.4.3.3 weiterbearbeiten; Erase-/ForEach-Grenzen
nicht ungeprüft als vollständige native Semantik deklarieren.


Echte ParamArray-Aliase/Arraysperren: PARAMARRAY_ALIASES.md dokumentiert skalare/
Objektvariablen und unterstützte Arrayzellen, eigene Literal-/Klammerwert-Slots,
doppelte Aliase, echte ByRef-Weitergabe, typisierte Rückschreibfehler und Schutz
unaufgelöster/Fassaden-/Arraywertfälle. Aktive ByRef-Arraywurzel liefert10 statt
3251 und wird nach Aufrufende freigegeben. App; initial GUI18,94s/DAO2,78s/
Original118,62s bestanden. Alte Literalpack-Erwartung angepasst und gewöhnliche
Aliaswerte geliehen statt Byte/Memo-Kopie; nur betroffenes Kern-/DAO-Bündel
wiederholt2/2 (7,97s:5,06/2,90), GUI/Original nicht erneut. Sieben Hashes unverändert;
bekannte Makrogrenze. Keine neuen CTests/Windows-/Vollsuitenbehauptungen;456/386/70
unverändert. Vorstufe9affc5a privat remote/Baum geprüft, dessen frisches DEB unter
/workspace/work/error-text-functions/package/linaccess-0.7.0-Linux.deb besteht
Build-ID/Payload/Rechte/keineFixtures, MDBTools1.0.1 und SQLite3.54.1 JOIN/NULL/Counter;
Paketintegration DAO2,76s/Original120,25s,2/2 in123,02s. Noch ohne diesen Aliasumfang.
Ohne Zwischenabschluss mit echten benannten Property-Setter-Indexargumenten
weiterprogrammieren; keine Getter zur Metadatenentdeckung ausführen.


Benannte Property-Setter-Indizes: NAMED_PROPERTY_SETTERS.md beschreibt echte
Let-/Set-Metadaten (auch ohne Get), optionale/Objektindizes und gemeinsame
Namens-/ByRef-Adresszuordnung einschließlich ParamArrayzellen. Projektklasse,
Default-Property und echte Formularfassade geprüft; externe ByRef-Grenze bleibt.
Getter werden nicht zur Entdeckung ausgeführt. App/Kern5,02s; GUI-Watchdog zeigte
fehlende Bereinigung der eingeführten NamedWriter-Testklasse vor späterer TEMP-
Katalogprüfung. Klasse scoped bereinigt, Katalogschutz unverändert; nur GUI
wiederholt, zuletzt18,54s bestanden. Sprach-Doppelmatrizen bleiben im Kern; GUI
behält echte Qt-/SQL-/Klassen-/Formularprüfungen und präzise20s-Schranke. Keine
Nachrichteninhalte in dauerhaftem Watchdoglog, keine neuen CTests/Windows-/
Vollsuitenbehauptungen;456/386/70 unverändert. Vorstufe859d2ef privat remote/Baum
geprüft. Letztes Original118,62s/DAO2,90s/Hashes liegen vor diesem Umfang; geprüftes
DEB9affc5a unter /workspace/work/error-text-functions/package/linaccess-0.7.0-Linux.deb
liegt vor ParamArray/benannten Settern. Ohne Zwischenabschluss Set-RHS-Auswertung
nach MS-VBAL5.4.3.9 vor dynamischer Zielauflösung weiterimplementieren.


Set-RHS-Reihenfolge: SET_EVALUATION_ORDER.md beschreibt einmalige Objektauswertung
vor dynamischem Empfänger/Index und Alias-Aktualisierung nach MS-VBAL5.4.3.9.
Echter Empfängerwechsel, Default, ReDim durch RHS, eigener/echter Formularsetter,
RHS-Fehler ohne Zielaufruf und danach Nothing→91 geprüft. Fixtureobjekte/Klassen
scoped bereinigt. App; erstes bestehendes Kern-/GUI-/DAO-/Original-Bündel4/4 in
145,89s (5,53/19,19/2,75/118,41), sieben Hashes unverändert. Bekannte Makrogrenze
bleibt; keine neuen CTests/Windows-/Vollsuiten-/Paketbehauptungen;456/386/70
unverändert. Vorstufe d618eaf privat remote/Baum geprüft. Letztes geprüftes
DEB9affc5a unter /workspace/work/error-text-functions/package/linaccess-0.7.0-Linux.deb
liegt vor ParamArray/benannten Settern/Set-Reihenfolge. Ohne Zwischenabschluss
mit echtem Undo/Redo vorgemerkter nativer Mehrfachlöschungen weiterprogrammieren.


Undo/Redo vorgemerkter Mehrfachlöschungen: TABLE_DELETE_HISTORY.md beschreibt die
bestehende Einzelaktionshistorie für saubere gespeicherte Schlüsselzeilen.
Strg+Z/Strg+Y ändern nur den Cache; genaue nicht zusammenhängende Löschmarkierungen
und Rücknahme/Wiederholung nach transaktionalem Fremdschlüsselfehler geprüft.
App gebaut; bestehendes GUI-/Währungsfeld-Bündel2/2 in19,17s (18,68/0,49).
Keine neuen Tests/IDs, keine Aussage zu Windows-Ereignissen oder committed SQL
Undo.456/386/70 unverändert. Vorstufe ba256ac privat remote/Baum geprüft.
Letztes Original118,41s/DAO2,75s/7Hashes liegt vor diesem Modellumfang; geprüftes
DEB9affc5a bleibt älter. Ohne Zwischenabschluss mit sauberer gespeicherter
Einzeldatensatzlöschung und Invalidierung der Aktionshistorie weiterarbeiten.


Saubere gespeicherte Einzeldatensatzlöschungen nutzen ebenfalls die geprüfte
Cache-Aktionshistorie aus TABLE_DELETE_HISTORY.md. Zelländerungen invalidieren
alte Wiederholungen; doppelte/ungültige Zeilen und schreibgeschützte Quellen
bleiben unverändert. Dirty-/schlüssellose Einzelpfade behalten bestehendes
Verhalten. App gebaut; bestehendes GUI-/Währungsfeld-Bündel2/2 in19,88s
(19,39/0,49), keine neuen CTests;456/386/70 unverändert. Vorstufe22a29fa privat
remote/Baum verifiziert. Weiter ohne Zwischenabschluss: native VarP/StDevP
in Berichts-/SQL-Aggregaten und DVarP/DStDevP in echten Domainaufrufen ergänzen.


Native Populationsstatistik: POPULATION_STATISTICS.md beschreibt VarP/StDevP in
SQL und Berichtssteuerelementen samt PDF/Requery, DVarP/DStDevP für echte Domains
inklusive benannter/Application-Aufrufe und die realen visuellen QBE-Funktionsmodi.
Null/leer/Einzelwert, große Offsets, Gruppen/DISTINCT/Parameter und Fehler geprüft.
App gebaut; erstes Bündel Kern5,32s/Bericht12,97s bestanden. Neue SQL-Testannahme
StDevP-Überlauf korrigiert (StDevP kann1e308 liefern, VarP überläuft): nur betroffenes
Abfragemodell erneut0,13s. Designer-Fixture erwartet zusätzlich verborgene Where-
Quellfelder; nur GUI erneut18,41s. Keine neuen CTests/Vollsuiten-/Windowsbehauptung;
456/386/70 unverändert. Sieben Hashes unverändert; Vorstufe c119f5e privat remote
und Baum geprüft. Anschließend aktuelles Debianpaket bauen/gegen echte Binärdatei
prüfen und unabhängige Implementierung fortsetzen; kein Zwischenabschluss.


Native Folder.Name-Zuweisung in NATIVE_FILE_SYSTEM_OBJECT.md: tatsächlicher
Ordner wird mit vorhandenem atomarem No-Replace-Helfer im bisherigen Elternpfad
umbenannt; gehaltene Objektidentität/Alias und Pfad erst nach Erfolg aktualisiert.
Unicode, Dateiinhalt, zweimaliges Umbenennen und unveränderte Pfade bei Ziel58,
Traversal3251 und Null94 geprüft. App gebaut; bestehender Kern im ersten Lauf1/1
in5,20s; keine weiteren Tests/IDs. F091 und456/386/70 bleiben offen/unverändert.
Vorstufe c8bd82b privat remote/Baum geprüft. Dessen aktualisiertes DEB liegt in
/workspace/work/population-statistics/package/linaccess-0.7.0-Linux.deb: tatsächliche
Build-ID, MDBTools1.0.1, SQLite3.54.1/nestedJOIN/NULL/counter, Rechte/Desktop/keine
Fixtures geprüft; danach DAO2,81s+Original120,75s,2/2 in123,57s. Bekannte Makrogrenze
bleibt. Sieben Hashes vor Paketlauf unverändert. Paket liegt vor Folder.Name.
Ohne Zwischenabschluss weitere allgemeine native Objekt-/Entwurfsanforderungen
implementieren; keinen Abschlussanruf für diese Etappe auslösen.


Native Ordnerverlagerung: Folder.Move und FileSystemObject.MoveFolder benutzen
den vorhandenen atomaren No-Replace-Helfer für einzelne tatsächliche Ordner und
exakte Zielpfade; Name-/Alias-Pfad erst nach Erfolg. Benannte Source/Destination,
Unicode, tatsächliche Unterdatei, frische Collections und unveränderte Pfade nach
58/3251/94 geprüft. App gebaut; bestehender Kern erster Lauf1/1 in5,26s. Keine
neuen Tests/IDs; F091/456/386/70 unverändert. Vorstufe1d8f606 privat remote/Baum
verifiziert. Geprüftes DEBc8bd82b unter /workspace/work/population-statistics/package
liegt vor Name/Move; letztes Original120,75s/DAO2,81s ebenfalls davor. Weiter ohne
Zwischenabschluss: feste numerische TRANSFORM-Kreuztabellen um Avg/Var/StDev/
VarP/StDevP mit nachgewiesener Double-Ausgabetypisierung erweitern.


Feste statistische Kreuztabellen: STATISTICAL_CROSSTABS.md dokumentiert tatsächliche
TRANSFORM Avg/Var/StDev/VarP/StDevP über nachgewiesene numerische physische Felder,
inklusive Aliase, Parameter, Null/fehlende Spalten und stabile Double-Metadaten.
Zuerst gemeinsames Schema-/Abfragemodell-Bündel; Modell0,12s bestand. Zwei konkrete
Implementierungslücken geschlossen: Feldkonverter REAL/DOUBLE und tatsächliche
DAO-Zuordnung nachgewiesener DOUBLE→Typ7. Danach nur betroffenes Schema+DAO2/2 in
2,98s (0,13/2,85); App gebaut. Vorhandene Count/Min/Max-Währungs-/Datums-/Ganzzahl-
Fälle bleiben im gleichen Schemafall geprüft. Keine neuen Tests/Katalogpunkte;
F068/F122 und456/386/70 unverändert. Vorstufe8cbf0af privat remote/Baum geprüft.
Geprüftes DEBc8bd82b und Original120,75s liegen vor diesen späteren Änderungen.
Ohne Zwischenabschluss weitere allgemeine SQL-/Objektmodellanforderungen umsetzen.


Feste Kreuztabellensummen: FIXED_CROSSTAB_SUMS.md beschreibt TRANSFORM Sum mit
nachgewiesenen Double/REAL- oder nativen Currency-Quellen. Geld verwendet den
vorhandenen exakten Currency-Akkumulator, keine rohe SQLite-Gleitkommasumme.
Tatsächliche Double-/Currency-Werte, Dezimalkompensation, Null und DAO-Typ7/5 im
bestehenden Schemafall geprüft: erster Lauf1/1 in0,13s; App gebaut. Integer/Single/
Decimal-Promotion und größere physische MONEY-Bereiche bleiben ausdrücklich offen.
Keine neuen Tests/IDs;456/386/70 unverändert. Vorstufe103af93 privat remote/Baum
geprüft. Geprüftes DEBc8bd82b bleibt älter. Ohne Zwischenabschluss weiter an echten
festen Pivot-Auswahlen First/Last arbeiten, ohne nichtpassende Null-Zeilen als
falsche erste/letzte Pivotwerte einzubeziehen.


Feste First/Last-Pivots: FIXED_CROSSTAB_SELECTORS.md beschreibt tatsächliche
Filterung passender Pivotzeilen vor der Auswahl, inklusive echter passender
Nullwerte. INTEGER2/Currency/Datum und reale DAO-Typen/Variantwerte gegen separate
native Zeilenscans geprüft; keine Behauptung über identische Windows-Besuchsfolge.
App gebaut; vorhandener Schemafall erster Lauf1/1 in0,17s. Nach zusammenhängenden
SQL-Erweiterungen originales Abfrageinventar einmal geprüft1/1 in47,83s:100
klassifiziert,78 ausgeführt (73 gewöhnlich/5 mit technischen Parametern ohne
Access-Ergebnisoracle),22 Grenzen (3 Quellspalten/9 Abhängigkeiten/10 Jet-Bindung).
Alle sieben Upload-/Fixturehashes unverändert. Vorstufe49a60c7 privat remote/Baum
geprüft. Keine neuen Tests/IDs; F068/456/386/70 unverändert. Geprüftes DEBc8bd82b
liegt vor späteren Ordner-/Kreuztabellenänderungen. Ohne Zwischenabschluss
weitere allgemeine DAO-/SQL-Metadatenlücken implementieren.


Nachgewiesene Boolean-DAO-Ergebnisse: DAO_BOOLEAN_RESULTS.md dokumentiert die
fehlende Abbildung tatsächlicher BOOL/BOOLEAN-Deklarationen auf dbBoolean1 und
Size1. Echte QueryDef.Fields ohne Modulausführung sowie Recordset-True/False-
Varianten, vorhandene Currency/Date/Long- und Identitätsgrenzen im gleichen DAO-
Ziel geprüft. App gebaut; Test-Compilerfehler (fehlender Namensraum) präzise
korrigiert, erster tatsächlicher DAO-Lauf1/1 in2,78s. Keine neuen Tests/IDs und
keine Boolean-Inferenz aus beliebigen Integern. F121/F122/456/386/70 unverändert.
Vorstufeff933a1 privat remote/Baum geprüft. Letztes Originalinventar100/78/22 und
sieben Hashes liegen vor diesem kleinen Metadatenumfang; geprüftes DEBc8bd82b
bleibt älter. Ohne Zwischenabschluss allgemeine native SQL-Funktionen fortsetzen.


Native SQL-Zahlenkonvertierung: SQL_NUMERIC_CONVERSIONS.md dokumentiert tatsächliche
CBool/CByte/CInt/CLng/CSng/CDbl-Callbacks über vorhandene VBA-Konvertierung sowie
Boolean/Byte/Integer/Long/Single/Double-Abfrage-/DAO-Metadaten. Rundung, Grenzen,
Verschachtelung, Parameter, WHERE und präzise Fehler im gemeinsamen bestehenden
Bündel geprüft. DAO erster Lauf2,92s bestand; SQLite-BLOB-Literal war im neuen
Test keine Access-SQL-Grammatik: durch tatsächliches Binärfeld ersetzt, nur
Abfragemodell erneut1/1 in0,12s. App gebaut. Keine neuen Tests/Katalogpunkte;
Datums-/Objekt-/Decimal-Transport, CStr und vollständige Regionalsemantik bleiben
offen.456/386/70 unverändert. Vorstufebc7cbf9 privat remote/Baum geprüft. Ohne
Zwischenabschluss nachgewiesene Count-Ergebnistypen in DAO/SQL ergänzen und danach
aktuelles Debianpaket bauen; geprüftes DEBc8bd82b liegt noch vor späteren Umfängen.


Count-Ergebnistypen: COUNT_RESULT_TYPES.md dokumentiert nachgewiesene Long/
INTEGER4-Metadaten tatsächlicher Count-Ausdrücke, QueryDef.Fields und echter DAO-
Felder (Typ4/Size4/Long-Variante). Gesamt-/Null-/Feld-/Gruppen-/Parameter-/leere
Counts und metadatenreine Projektzustandserhaltung im ersten bestehenden Bündel
2/2 in3,12s (DAO2,99/Abfragemodell0,13) geprüft; App gebaut. Keine neuen Tests/IDs,
keine Ableitung aus beliebigen Integerwerten; F068/F121/F122/456/386/70 unverändert.
Vorstufe3fa14b9 privat remote/Baum geprüft. Anschließend ohne Zwischenabschluss
aktuelles Debianpaket bauen und prüfen, GUI/Originalpfad nach zusammenhängenden
SQL-/DAO-Änderungen einmal integrieren; geprüftes DEBc8bd82b bleibt bis dahin älter.


Direkte deklarierte Parameterprojektionen: PARAMETER_RESULT_TYPES.md dokumentiert
Byte/Short/Long-Integer/Single/Double/Boolean-YesNo-Metadaten einschließlich
normalisierter Aliase, kompletter Klammern, echter DAO-Typen/Größen/Varianten und
Null. Nur direkte Parameter annotiert; Byte-Negation erhält keinen falschen Byte-
Typ. App gebaut; erstes bestehendes DAO-/Abfragemodell-Bündel2/2 in3,02s
(2,89/0,13), keine neuen Tests/IDs;456/386/70 unverändert. Vorstufecc19000 privat
remote/Baum geprüft. Dessen aktualisiertes DEB /workspace/work/count-result-types/
package/linaccess-0.7.0-Linux.deb besteht echte Build-ID/MDBTools1.0.1/SQLite3.54.1/
nestedJOIN/NULL/counter/Rechte/Desktop/keine-Fixtures-Prüfung; danach GUI19,32s und
Original119,69s,2/2 in139,02s, alle sieben Hashes unverändert. Bekannte Makrogrenze
bleibt. Paket und Integration liegen vor dieser Parameterannotation. Ohne
Zwischenabschluss native SQL-Potenzoperatoren mit echter VBA-Präzedenz fortsetzen.


Nativer SQL-Potenzoperator: SQL_POWER_OPERATOR.md dokumentiert tatsächliche
^→VBA-Auswertung mit strukturellen Atomen, Linksassoziativität/Vorzeichen nach
MS-VBAL5.6.9.1, Klammern/Funktionen/Parametern/Feldern/WHERE, Double-/DAO-Typen und
Null. Quotierte/literale ^ bleiben intakt, Originalüberschriften bleiben sichtbar.
App gebaut; erstes bestehendes DAO-/Abfragemodell-Bündel2/2 in3,17s (3,02/0,15).
Vorhandene Zero-negativ-/Unterlaufgrenze3251 beibehalten; kein Windows-/voller SQL-
AST-/Datum-/Decimal-Anspruch. Keine neuen Tests/Katalogpunkte;456/386/70 unverändert.
Vorstufeca0ebc0 privat remote/Baum geprüft. Geprüftes DEBcc19000 unter
/workspace/work/count-result-types/package/linaccess-0.7.0-Linux.deb und GUI19,32/
Original119,69/7Hashes liegen vor Parameterannotation/Potenz. Ohne Zwischenabschluss
weiter echte typisierte numerische SQL-Arithmetik und Divisionsfehler umsetzen.


Typisierte numerische SQL-Arithmetik: TYPED_SQL_ARITHMETIC.md beschreibt echte
Double-Operationen, gebrochene Ganzzahldivision, nachgewiesene Feld-/Parametertypen,
Null, WHERE/HAVING und reale SQLite-Ursprünge auch bei Sternprojektionen. Read-only
Treiber-Cache prüft Schema/Katalog/Definition/Speicher; keine Werte/VBA-Module.
0/0 und numerischer Überlauf melden normgemäß6, sonstige Division durch0 meldet11;
SQL-Fehler11 jetzt korrekt weitergereicht. Core5,55s/GUI18,91s bestanden vor letzten
SQL-Verfeinerungen; zuletzt DAO3,00s und Abfragemodell0,19s. Import-Cachefixture
verwendet tatsächlichen REAL-Import mit Double-Metadaten samt normaler Logtabelle;
DAO erstellt physisch DOUBLE. Metadatenänderung ohne Schemaversion, TEMP-Schatten,
reale Untertypen und unveränderte Überschriften geprüft. Normale indizierte
ID-Prädikate bleiben unverpackt, tatsächlicher Scan-Zähler0 bestätigt. Keine
neuen Tests/IDs; F041/F067/F068/F071/F121/F122/456/386/70 unverändert. Vorstufe2fdd9dc
privat remote/Baum geprüft. Geprüftes DEBcc19000 liegt vor späteren Umfängen.
Danach ohne Zwischenabschluss aktuellen Paket-/Original-/Inventarlauf ausführen
und weitere native Zahlentyp-/Operatorschritte fortsetzen. Kein Abschlussanruf.


Tatsächliche Report-RunningSum-Ausgabe: REPORT_RUNNING_SUM.md dokumentiert
Over-All-Detailpräfixe aus tatsächlichen geordneten numerischen Feldern und =1,
Currency-/Decimal-Summenpfad, Formatierung, berechnete Abhängigkeiten, Mehrseiten-
PDF und wiederholte Ausgabe ohne Doppelzählung. Steuerelement-/Datensatzwerte
werden wiederhergestellt; Null/Gruppen/andere Ausdrücke/Running-CanGrow bleiben
explizit gesperrt. Vorhandene Druck-/Ereignisgrenzen gelten. App/Layout gebaut,
ein bestehender gebündelter access_2003_layouts-Lauf13,23s bestanden, echte PDF-
Texte verglichen und atomarer Zielschutz geprüft. Keine neuen Test-/Katalog-IDs;
456/386/70 unverändert. Vorstufe2db95826 privat remote/Baum geprüft. Geprüftes
DEB8fd1d50 liegt vor nativen Operatoren und RunningSum. Ohne Zwischenabschluss
native Berichtsentwurfsbedienung für laufende Feldsummen fortsetzen.


Native RunningSum-Berichtsbedienung: optionales running_sum_field im echten
Berichtsentwurf, vorbereitete tatsächliche Ergebnisspalten/-typen ohne Zeilen/
Benutzerfunktionsausführung, fehlend3265/nichtnumerisch3251. Speicherung/Wieder-
öffnung, laufendes Detailfeld ohne CanGrow und unveränderte andere Textfelder,
VBA-/Quell-/Seitenmetadaten-Zusammenführung geprüft. Präfixumfang auf1MioZellen
begrenzt. App/GUI/Layout gebaut; ein bestehendes Bündel2/2 in32,05s (19,09/12,96)
bestanden, PDF-Zeilen mit tatsächlichen Ordinalen1/2/3 und Currency-Präfixen geprüft.
Keine neuen Test-/Katalog-IDs;456/386/70 unverändert. Vorstufe3dd9019c privat remote/
Baum geprüft. Ohne Zwischenabschluss jetzt aktuellen DEB-/Originalpfadnachweis
für beide Operator- und RunningSum-Umfänge bündeln, danach weiter implementieren.


Gebundene native Arithmetik: BOUND_NUMERIC_ARITHMETIC.md dokumentiert tatsächliche
Deklarationskategorien getrennt von VarType, lokale/Argument-/Alias-/Funktions-
und Arrayelementtypen, Literale/unterstützte explizite Konvertierungen und CVar.
Festes Integer/Long/Single meldet Überlauf6, nachgewiesene Variant-Werte erweitern
sich Long/Double; Unary-Minima und Date-Grenzen unterscheiden feste/Variant-Typen.
Erster gezielter Lauf fand echten ParamArray-Bindungsfehler99*1000: Element ist
Variant, darf keinen alten RHS-Static-Tag übernehmen. Tatsächliche Arraylesegrenze
korrigiert. Allgemeine Host-/Member-/Cache-/rohe Lokalwerte ohne Deklarationsresolver
bleiben ungebunden; interne Hilfskonvertierungen erzeugen keine Quell-Deklaration.
CVar erhält tatsächliche Daten-/Objektidentität, liest kein ScalarDefault hinzu.
Variant-Byte-Überlauf, vollständige Intrinsics/Member/Logik/Integerdivision-
Provenienz/Misch-Currency/AST/IEEE/Zusatzsuffixe/Windows bleiben offen.
App/Core/DAO/Abfragemodell/GUI gebaut. DAO2,93s/Modell0,19s bestanden zunächst;
nach Array-/Hostkorrektur Core5,67s/GUI19,70s gemeinsam2/2 in25,37s; letzte rohe
Lokalgrenze nur Core5,47s erneut bestanden. Testhelper-Final-Ableitung kompiliert
zunächst nicht, vor bestandenem Lauf behoben. Keine neuen CTest-/Katalog-IDs;
456/386/70 unverändert. Vorstufe7d78d04c privat remote/Baum geprüft. Letztes DEB/
Originalinventar46,35/Originalpfad120,51/7Hashes liegt auf5df3cba2 vor Literal- und
Deklarationsänderungen. Ohne Zwischenabschluss native numerische Intrinsic-
Rückgabebindungen und weitere tatsächliche Anforderungen fortsetzen.


Numerische Intrinsic-Rückgabebindungen: NUMERIC_INTRINSIC_BINDINGS.md ergänzt
tatsächliche deklarierte Variant-Ergebnisse für Abs/Int/Fix/Round/Sgn und
Choose/Switch/IIf/Nz; ausgewählte feste Integer-Argumente machen diese Funktionen
nicht selbst statisch. Sqr/Exp/Log/Sin/Cos/Tan/Atn sind nachgewiesenes Double.
Unveränderte eigentliche Zahlen-/Fehler-/Präzisions-/Eager-Pfade, generische Host-
Aggregate ohne fälschlich übernommene Cache-Deklaration, Access-Eval unverändert.
App/Core/DAO/Abfragemodell gebaut; erstes bestehendes Bündel3/3 in8,36s
(5,18/2,98/0,20) bestanden. Keine neuen CTest-/Katalog-IDs;456/386/70 unverändert,
keine Voll-Varianten-/AST-/Windows-Parität. Vorstufef38f6970 privat remote/Baum
geprüft. Letztes geprüftes DEB5df3cba2 liegt vor Literal-/Deklarationsumfängen.
Ohne Zwischenabschluss aktuellen Paket-/Originalnachweis für diese Umfänge
bündeln und danach weitere tatsächliche Anforderungen implementieren.


Native typisierte Logik: NATIVE_TYPED_LOGIC.md dokumentiert Nicht-Null-And/Or/
Xor/Eqv/Imp mit Byte-/Integer-/Boolean-/Long-Ergebnissen und tatsächlicher8Bit-
Eqv/Imp-Maske. Not/Binäroperationen tragen nachgewiesene feste/Variant-Bindungen
in Folgearithmetik; vorhandene Nullpfade und Access-Eval-Sonderpfad unverändert.
MS-VBAL-Allgemeintabelle Or/Null ist kein vollständiges Access-Orakel: tatsächliche
Referenz Null Or False ergibtNull. Unbelegte numerische Imp/Null-Fälle bleiben
bewusst offen. Metadatenlesepfad vermeidet Typstringkopie ohne Auswahländerung.
App/Core/DAO/GUI gebaut; erstes bestehendes Bündel3/3 in27,29s (5,34/18,72/3,22)
bestanden. Tatsächliche native Byte-VBA-Funktion über SQL/DAO liefertEqv252 mit
Type2/Size1/Byte-Value. Keine neuen CTest-/Katalog-IDs;456/386/70 unverändert.
Vorstufe4cd0f848 privat remote/Baum geprüft und aktuellstes geprüftes DEB unter
/workspace/work/numeric-intrinsic-bindings/package/linaccess-0.7.0-Linux.deb:
Build-ID/SQLite3.54.1/MDBTools1.0.1/Inhalt/Rechte/Desktop/keine-Fixtures bestanden,
danachGUI19,97/Inventar46,77/Original118,66,3/3 in185,40s;100/78/22 und7Hashes
unverändert. Makroquellengrenze bleibt. Paket liegt vor Logikumfang. Ohne
Zwischenabschluss native Ganzzahldivision/Modulo und weitere Anforderungen
fortlaufend implementieren, kein Abschlussanruf vor tatsächlichem Gesamtabschluss.


Native Ganzzahldivision/Modulo: INTEGRAL_DIVISION_TYPES.md dokumentiert tatsächliche
Byte-/Integer-/Long-Ergebnisse für unterstützte integrale Operandentypen und
Deklarations-/Variant-Provenienz für Folgearithmetik. Feste signed-Minima bleiben
Überlauf6, Variant-Integer/Long erweitern Long/Double. Bisherige geprüfte Long-
Eingabekonvertierung, int64-Zwischenrechnung für32Bit-Werte, Abschneiden Richtung0,
vorzeichenbehaftetesMod, Null und Nullteiler11 bleiben. Gemischte weitere Typen,
Suffixe/64Bit/Windows nicht pauschal behauptet. App/Core/DAO gebaut; erstes
bestehendes Bündel2/2 in8,72s (5,55/3,17) bestanden. Tatsächliche native Byte-
Funktion prüft Quotienten-VarType und liefertMod1 über SQL/DAO Type2/Size1/Byte.
Keine neuen CTest-/Katalog-IDs;456/386/70 unverändert. Vorstufe9226c05e privat remote/
Baum geprüft. Letztes geprüftes DEB4cd0f848 liegt vor Logik-/Divisionsumfang.
Ohne Zwischenabschluss genaue gemischte Currency-/Integral-Arithmetik und danach
weitere offene Anforderungen fortlaufend umsetzen. Kein Abschlussanruf.


Gemischte native Currency-/Zahlenarithmetik: CURRENCY_NUMERIC_ARITHMETIC.md
beschreibt +,- mit ganzzahligen/Single/Double/Boolean/Empty-Operanden und * mit
Integral/Boolean/Empty in beiden Reihenfolgen über vorhandene exakte64Bit-
Skalierung/int128-Kernel. Keine Double-Näherung für vollständige Currencywerte;
Range6 und ScalarDefault-einmal belegt. Microsoft-Plus-/Minus-Dokumentation
bestätigt PräzisionsrangCurrency überDouble/Single für+,-. Double-Produkte bleiben
Double; vollständige Single-Produktpromotion/Strings/Locale bleiben unbewiesen.
Decimal-/Date-/Null-/Access-Eval-Grenzen bleiben; keine erweiterte physische
MONEY-/rawSQL-/MDB-Parität. App/Core/DAO gebaut, finales bestehendes Bündel2/2
in8,64s (5,37/3,27) bestanden; tatsächliche native Funktion liefertCurrency2.0003
überSQL/DAO Type5/Size8/20003Einheiten. Keine neuen CTest-/Katalog-IDs;456/386/70
unverändert. Vorstufeb80c9521 privat remote/Baum geprüft. Letztes DEB4cd0f848 liegt
vor späteren Umfängen. Ohne Zwischenabschluss nach öffentlich dokumentierter
Byte-Variant-Erweiterung weitere Sprach-/Datenbankanforderungen fortsetzen.


Byte-Variant-Erweiterung: BYTE_VARIANT_WIDENING.md stützt fehlende Byte→Integer→
Long-Stufe auf tatsächliche öffentliche Microsoft-Plus-/Minus-/Multiplikations-
Dokumentation. Nur nachgewiesene Variant-Deklaration plus tatsächlicher Byte-
Subtyp, feste Byte-Überläufe bleiben6; ungebundene/Access-Eval-/physische SQL-
Grenzen unverändert. App/Core/DAO gebaut; erstes bestehendes Bündel2/2 in8,68s
(5,50/3,18) bestanden. Echter nativer Integer-VBA-Rückgabewert überSQL/DAO
Type3/Size2/Integer256 geprüft. Keine neuen CTest-/Katalog-IDs;456/386/70 unverändert.
Vorstufe04db5776 privat remote/Baum geprüft. Letztes geprüftes DEB4cd0f848 liegt vor
späteren Umfängen. Ohne Zwischenabschluss typisierte native Zahlenliterale mit
Suffices und weitere offene Anforderungen fortlaufend implementieren.


Native Zahlenliteral-Typzeichen: NATIVE_NUMERIC_SUFFIXES.md dokumentiert%Integer,
&Long,!Single,#Double,@Currency im expliziten VBA-Lexermodus. Currency direkt
Dezimaltext→exaktesFestkomma, Single direktfrom_chars→float ohneDouble-Rundung;
32Bit-Ranges, Exponenten, Präzedenz, Dates und bisherige angrenzende Konkatenation
bleiben. Metadaten-Klassentypescanner verwenden denselben nativen Lexermodus.
Keine Identifier-Typzeichen-/signedMinimum-Faltung-/Vollcompiler-/SQL-/Eval-
Behauptung. App/Core/DAO gebaut; erstes bestehendes Bündel2/2 in8,61s
(5,38/3,23) bestanden. Tatsächliche VBA-Funktion verwendet alle5Zeichen samt
OptionalLong=2& und liefertüberSQL/DAO Type5/Size8/exakt20002Currencyeinheiten.
Keine neuen CTest-/Katalog-IDs;456/386/70 unverändert. Vorstufe56f0e012 privat
remote/Baum geprüft. Letztes geprüftes DEB4cd0f848 liegt vor späteren Umfängen.
Ohne Zwischenabschluss weitere tatsächliche DAO-/Datenbankfunktionen fortsetzen.


DAO Recordset.CopyQueryDef: DAO_COPYQUERYDEF.md beschreibt echte unabhängige,
unbenannte QueryDef-Kopien aus parameterfreien QueryDef-Cursorn mit beim Öffnen
bewahrtem Access-SQL. Spätere Quelländerung/Close und Cursor-Clone sind belegt;
parameterisierte Quellen sowie Cursor ohne belegte QueryDef-Herkunft bleiben3251.
App/DAO gebaut, einmaliges bestehendes dao_execution-Bündel1/1 in3,23s bestanden,
einschließlich tatsächlichem typisiertem VBA-Aufruf und unabhängigen SQL-Zeilen.
Keine neuen Test-/Katalog-IDs,456/386/70; F054/F055 bleiben offen. Vorstufe
df22cf9b privat remote/Baum geprüft. Kein Gesamtabschluss/Abschlussanruf.


DAO Filter/Sort/abgeleitete Snapshots: DAO_DERIVED_SNAPSHOTS.md dokumentiert echte
SQL-Filterung/Sortierung bestehender querybasierter Cursor über explizites
OpenRecordset(dbOpenSnapshot). Ausgangscursor unverändert, neue Ergebnisse und
Klon-Eigenschaften unabhängig; Parameter und Currency-Metadaten bleiben erhalten.
Editierbare/implizite Modi, Tabellen-/Frozenadapter und volle Oracleidentität offen.
App gebaut; bestehendes DAO-Bündel1/1 in3,37s bestanden nach Korrektur eines
Prüfcode-Feldnamens beim Übersetzen. Keine neuen CTest-/Katalog-IDs;456/386/70.
Vorstufe ad42eadace27a518fbbaf1f7a9819facc1e383c1 privat remote/Baum geprüft:
neues DEB geprüft, bestehende Integration3/3 in188,36s (GUI19,85/Inventar47,59/
Original120,91), sieben geschützte Hashes unverändert. Inventar100/78/22;
Original-Makrogrenze bleibt. Paket liegt vor abgeleiteten Snapshots. Kein
Gesamtabschluss/Abschlussanruf; weitere allgemeine Anforderungen fortsetzen.


Native Ordner-Metadaten: NATIVE_FILE_SYSTEM_OBJECT.md beschreibt Folder.Size
über tatsächliche verschachtelte/versteckte Dateien mit exakter Double-Ganzzahl-
Grenze, Besuchs-/Tiefenlimits und ausdrücklichen Link-/Spezialdatei-Grenzen sowie
Folder.DateLastModified über vorhandene lokale Date-Konvertierung. Kein Windows-
Variant-/atomarer Snapshot-/FSO-Vollparitätsanspruch. App/Core gebaut, einmaliges
bestehendes Core-Bündel1/1 in5,56s bestanden, reale kleine und sparse>2GiB-Dateien,
Zeitstempel, Link-/Löschgrenzen und typisierter VBA-Aufruf.456/386/70 unverändert,
keine neuen Test-/Katalog-IDs. Vorstufe45a6f86a privat remote/Baum geprüft.
Aktuelles geprüftes DEB bleibt ad42eada vor abgeleiteten Snapshots/Ordner-Metadaten.
Ohne Zwischenabschluss weitere allgemeine offene Anforderungen fortsetzen.


Native DateCreated/DateLastAccessed für File und Folder: vorhandene
NATIVE_FILE_SYSTEM_OBJECT.md dokumentiert tatsächliche Linux-statx-Zeitstempel,
Verfügbarkeits-/Objektartprüfung und3251 bei fehlender Erstellzeit; keine
Änderungs-/Inode-/Jetzt-Ersatzwerte. Vorhandene lokale Date-/Range-Konvertierung
nimmt zusätzlich geprüfte Unixsekunden an. App/Core gebaut; einmaliges bestehendes
Core-Bündel1/1 in5,45s bestanden: echte getrennte Zugriffs-/Änderungszeiten,
typisiertes VBA, tatsächliche/ausdrücklich fehlende Erstellzeit und Löschgrenzen.
Keine neuen Test-/Katalog-IDs;456/386/70 unverändert. Vorstufe1da6ad45 privat
remote/Baum geprüft. Geprüftes DEB weiterhin ad42eada vor späteren Umfängen.
Ohne Zwischenabschluss weitere allgemeine offene Anforderungen umsetzen.


Quelllokale qualifizierte Enum-As-Typen: VBA_ENUM_TYPE_BINDINGS.md dokumentiert
Modulname.Enumname nur aus tatsächlichem Attribute VB_Name derselben Quelle,
inklusive Const-/Feld-/Parameter-/Rückgabe-/Array-/Inline-Bindungen alsLong.
Keine erfundene Modulidentität oder projektweite Sichtbarkeit; fremde Qualifier
bleiben abgewiesen. Quellreine Parsercache-Identität bleibt unverändert. App/Core/
GUI gebaut; einmaliges bestehendes Bündel2/2 in25,29s (5,51/19,78), unveränderter
präziser20s-GUI-Watchdog. Echte SQL-Funktion mit qualifiziertem Enum-Resultat
registriert Long und liefert1234. Keine neuen Test-/Katalog-IDs;456/386/70.
Vorstufef4e44eb2 privat remote/Baum geprüft; DEB weiterhinad42eada vor späteren
Umfängen. Projektweite Enum-Bindung benötigt sichtbarkeitsabhängiges Typumfeld
und dessen Cache-/Modulzustandsidentität; keine Klassenheuristik einbauen.
Ohne Zwischenabschluss weitere echte offene Anforderungen fortsetzen.


Native MonthName/WeekdayName: CALENDAR_FUNCTIONS_RUNTIME.md dokumentiert echte
lokalisierte String-Namen über aktuellen QLocale-/Standalone-Kalenderadapter,
WeekdayName-Systemstart0 bzw. explizite VBA-Sonntag1..Samstag7, Abkürzungen,
benannte/ausgelassene/qualifizierte Aufrufe und kontrollierte Range-/Null-Grenzen.
Formular-, Klassen-, Standard-/Importeditor-, direkter SQL- und quellbasierter
SQL-Modulhost integriert; kein zusätzlicher SQL-UI-/DAO-Zugriff. SQL-Funktionen
wegen Sprachabhängigkeit nicht als deterministisch registriert. WindowsNLS/
Hijri-/vollständige Metadaten-/Oracle-Parität bleibt offen. App/Core/GUI/Format
gebaut; einmaliges bestehendes Bündel3/3 in24,21s (5,44/18,72/0,05),20s-Watchdog
unverändert. Reale deutsche/englische SQL-Namen und echte Formular-/Modul-/SQL-
Modulfunktionen geprüft. Keine neuen CTest-/Katalog-IDs;456/386/70. Vorstufe
d339d4b2 privat remote/Baum geprüft. Letztes geprüftes DEB weiterhinad42eada vor
den späteren Umfängen; anschließend gebündelt Paket/Originalbelege erneuern und
weitere offene Anforderungen ohne Zwischenabschluss umsetzen.


DAO SourceField/SourceTable: DAO_FIELD_ORIGINS.md dokumentiert echte physische
main-Schema-Herkunft aus vorbereiteten Statements für Recordset/QueryDef/TableDef,
Aliasse/Views/Joins/Snapshots, schemareine QueryDef-Metadaten ohne Parameterausführung
und erweiterte Herkunftsidentität. Keine erfundene Label-/UNION-/TEMP-/berechnete
Herkunft. Tabellen-Recordset.SourceTable ist ausdrücklich unbelegt3251, Index- und
Relationskontexte bleiben geschützt. App/DAO/Query/Currency gebaut; einmaliges
bestehendes Bündel5/5 in3,90s (0,09/0,06/0,16/3,41/0,18) bestanden. Kein neuer
CTest-/Katalogpunkt;456/386/70 unverändert. Vorstufe4ebc3f6c privat remote/Baum
geprüft; neues DEB/Paketprüfung und Originalbündel2/2 in167,38s (47,31/120,07),
sieben geschützte Hashes unverändert. Inventar100/78/22, fehlende Makros weiterhin
Grenze. Kein erneuter bereits bestandener GUI-Lauf; DEB liegt vor Feldherkunft.
Nächster belegter DAO-Fehler: AddNew/Update soll einen zuvor gültigen aktuellen
Datensatz erhalten; neuer Datensatz über LastModified. Bisher springt update()
pauschal zur neuen Identität. Erhaltenen aktuellen Datensatz und LastModified
getrennt behandeln; leere/BOF/EOF-/entfernte Vorposition nicht ungeprüft als
Windows-konform behaupten. Danach weitere offene Anforderungen ohne Abschluss.


DAO AddNew/Update aktueller Datensatz: DAO_ADDNEW_CURRENT.md dokumentiert Erhalt
eines zuvor gültigen weiterhin vorhandenen aktuellen Datensatzes und davon
unabhängiges LastModified des eingefügten Datensatzes. Tabellen/Dynaset, Klone,
Bookmarkidentität, Cancel/Fehler/Retry und echtes VBA belegt. Erstes DAO-Bündel
legte QSQLITE-null-lastInsertId bei gültiger physischer Zeilenkennung0 offen;
gezielt durch tatsächliche SQLite-Verbindungskennung ohne Double/zusätzlichesSQL
behoben. Nur betroffenen DAO-Lauf wiederholt:1/1 in3,27s bestanden, App gebaut.
Leere/BOF/EOF/verschwundene Voridentität, Navigation mit Entwurf, dynamische
Anfügereihenfolge und Windows-Vollorakel bleiben offen. Keine neuen CTest- oder
Katalog-IDs;456/386/70. Vorstufec77bdf6a privat remote/Baum geprüft. Geprüftes DEB
weiterhin4ebc3f6c vor Feldherkunft/AddNew. Ohne Zwischenabschluss weitere echte
DAO-Zustands-/Sprach-/Datenbankanforderungen implementieren; keine AddNew-ADO-
FieldList/Values-Signatur in DAO erfinden (DAO-AddNew ist parameterlos).


DAO Entwürfe bei tatsächlicher Navigation: DAO_PENDING_NAVIGATION.md dokumentiert
Verwerfen von Edit-/AddNew-Wert-/Chunkpuffern erst nach erfolgreicher Zielprüfung
für tatsächliche Positionswechsel in Move/Positions-/Bookmark-/Find-/Seek-Pfaden.
Keine SQL-Speicherung, Klonentwürfe getrennt; Reentranz-/Versionsprüfungen erhalten.
Ungültige Ziele/Suchkriterien erhalten Entwürfe. No-op-Navigation bleibt3219;
Clone/Index/Requery/GetRows und vollständige Windows-Fehlerpriorität weiter offen.
Formular-UI-Speicher-/Cancelereignisse unverändert. App/DAO gebaut; einmaliges
bestehendes DAO-Bündel1/1 in3,67s bestanden, echte Tabellen/Dynaset Edit-/AddNew-
Bewegungsmatrix und unabhängiger SQL-Nichtschreibbeleg. Keine neuen CTest- oder
Katalog-IDs;456/386/70. Vorstufec0dc87a9 privat remote/Baum geprüft. Letztes
geprüftes DEB4ebc3f6c liegt vor Feldherkunft/DAO-Zustandsänderungen. Kein neuer
GUI-/Original-/Paketbeleg für diese Etappe. Ohne Zwischenabschluss weitere
allgemeine offene Anforderungen fortlaufend implementieren.


Native Dynaset-Reihenfolge: DAO_DYNASET_ORDER.md beschreibt linearen Merge nach
exakter physischer Identität bei nichtinvalidierendem Refresh nativer Tabellen-
Dynasets: vorhandene Zeilenordnung behalten, neue Zeilen anhängen, anschließend
vorhandene Bookmark-/Klon-/Cursorwiederherstellung. Nachfolgende Updates setzen
Reihenfolge nicht auf SQLite-Physicalorder zurück. Tabellen-Cursor/Indexpfade und
explizites invalidierendes Requery bleiben getrennt; keine neue SQL-Ausführung.
App/DAO gebaut, einmaliges bestehendes DAO-Bündel1/1 in3,24s bestanden: reale
Ordnung mit neuem Schlüssel0 (Tabelle0/1/2/3 vs.Dynaset1/2/3/0), Klone, Updates,
weitere Inserts und unveränderte Positions-/LastModified-/Entwurfsmatrizen.
Keine neuen CTest-/Katalog-IDs;456/386/70. Vollständige externe Live-Dynaset-/
Delete-/Orakelparität bleibt offen. Vorstufe32d4a49d privat remote/Baum geprüft.
Geprüftes DEB4ebc3f6c liegt vor Feldherkunft/DAO-Zustands-/Ordnungsänderungen;
anschließend gesammelte Änderungen gebündelt paketieren/Original prüfen und
weitere Anforderungen ohne Zwischenabschluss fortsetzen.


Physische DAO-Quellbindung: DAO_PHYSICAL_SOURCE_SCOPE.md dokumentiert tatsächliche
main-Qualifikation der Tabellen-Lese-/INSERT-/UPDATE-/DELETE-Pfade und explizite
Prüfung auf TEMP-Namensüberlagerung vor Laden/Reload und Mutationsbeginn. Verhindert
Mischen echter main-Metadaten mit anderen TEMP-Zeilen/-Zieltabellen. Verweigerung3251
vor SQL erhält Entwurf; nach Entfernen der Kollision wird beabsichtigtes main-
Update möglich. Cached main-Herkunft bleibt tatsächlich zu cached main-Werten.
App/DAO gebaut; einmaliges bestehendes DAO-Bündel1/1 in3,42s (gesamt3,43) bestanden,
unabhängige SQL-Belege für beide unveränderten Tabellen und TEMP-only-Grenze.
Keine neuen CTest-/Katalog-IDs;456/386/70. Vorstufea4bbf4a4 privat remote/Baum geprüft:
neues DEB/Paketprüfung und gebündelter Originaltest1/1 in118,03s, sieben geschützte
Hashes unverändert, Makrogrenze bleibt. Bereits bestandene GUI-/Inventarläufe nicht
wiederholt; tatsächlicher letzter Inventarbeleg4ebc3f6c100/78/22. Dieses DEB liegt
vor physischer Quellbindung. Danach weitere allgemeine offene Anforderungen ohne
Zwischenabschluss fortsetzen; keine umfassende roheSQL-/Provider-/Crossconnection-
Schattierungs-/Generationsparität behaupten.


DAO Restartable/Origin-Properties: DAO_RESTARTABLE_PROPERTIES.md dokumentiert
Boolean-Capability für unterstützte Query/Dynaset/Snapshot-Cursor, dokumentiertes
False für Tabellen-Cursor sowie False für eigenständig nicht-requeryfähige Frozen-
Adapter. Stale Query nach Rollback kannFlag lesen und anschließend echtRequery
aufrufen; native bisherige Tabellen-Requery-Erweiterung bleibt getrennt. Native
SourceField/SourceTable über echte schreibgeschützte Properties-Objekte/dbText10,
Quelle/Getterschutz unverändert. App/DAO gebaut; einmaliges DAO-Bündel1/1 in3,37s
bestanden inklusive echtem VBA und tatsächlicher Rollback-Wiederabfrage. Keine
neuen CTest-/Katalog-IDs;456/386/70. Vorstufe718ca355 privat remote/Baum geprüft;
letztes DEBa4bbf4a4 vor Quellbindung/Metadaten. Moderne veröffentlichte DAO-PIA
15.0.4420.1018 nur als Vertrag inspiziert: GetRows.NumRows istoptional OHNE
Metadaten-Standardwert, AddNew parameterlos. Keine Default1-Behauptung aus
ADO/Optionalität ableiten; keine PIA inGit oder Original-Ausführung.
Ohne Zwischenabschluss weitere echte DAO-/Datenbankanforderungen implementieren.


DAO gelöschte Cursorposition: DAO_DELETED_POSITION.md dokumentiert einen echten
logischen Löschmarker mit Nachbaridentitäten statt Sprung zum ersten Datensatz.
Betroffene Klone, ungültige Bookmarks, Next/Previous/Move/FindNext und BOF-/EOF-
Zustände nach letzter Löschung sowie invalidierendes Requery werden berücksichtigt.
Pending Edit nach Klonlöschung prüft Current vor Identitätsindex/SQL und kann
keine Nachbarzeile überschreiben; Reload-Fehlerrestore umfasst neue Zustände.
Native3021 bleibt; vollständige3167-/Eigenschafts-/Provider-/Concurrency-Identität
unbelegt. App/DAO gebaut; einmaliges bestehendes DAO-Bündel1/1 in3,40s bestanden,
reale Tabellen/Dynaset-/Klon-/Lösch-/Entwurfsfälle und echtes VBA Delete/MoveNext
löscht alle3Zeilen, unabhängigSQLCOUNT0. Keine neuen CTest-/Katalog-IDs;456/386/70.
Vorstufe063402b9 privat remote/Baum geprüft. Letztes geprüftes DEBa4bbf4a4 liegt
vor Quellbindung/Restartable/Löschzustand; anschließend diese gesammelten Änderungen
gebündelt paketieren/Original prüfen und weitere offene Anforderungen fortsetzen.
Spezifische BOF-/EOF-/Move-Dokumentation verwenden; moderne generierte Delete-
Seiten enthalten widersprüchliches unsupported, keine Blindübernahme als Orakel.


DAO schreibbare Sternabfragen: DAO_EDITABLE_STAR_QUERIES.md dokumentiert echte
Mode2-Edit/Update/Delete auf streng geprüften schlüsselgebundenen einzelnen main-
Tabellen mit vollständigem */t.*. WHERE wird beim Laden und Schreiben angewandt;
ORDER BY nutzt physische Felder. Klone, Herkunft und abgeleitete Snapshots bleiben
nutzbar. Ungefiltertes AddNew unterstützt; gefiltertes AddNew explizit3251 vor
Entwurf/Defaults/Reservierung. Projektionen/Parameter/komplexe/keylose Quellen und
Mode4 werden nicht schreibbar erklärt. App/DAO/QueryModel/Currency gebaut; einmaliges
bestehendes Bündel5/5 in3,88s bestanden, DAO3,34s; echte SQL-/VBA-/Klon-/Stale-
Predicate-/Retry-/Lösch-/Insert-Belege. Keine neuen CTest-/Katalog-IDs;456/386/70.
Vorstufec67cbd9e privat remote/Baum geprüft, neues DEB geprüft und Original1/1
in120,41s bestanden, sieben geschützte Hashes unverändert; Makrogrenze bleibt.
Dieses Paket liegt vor Sternabfragen. Inventar unverändert letzter tatsächlicher
Lauf4ebc3f6c100/78/22. Keine Windows-Orakel-/vollständige Dynaset-Paritätsbehauptung.
Weitere allgemeine Anforderungen ohne Zwischenabschluss umsetzen.


DAO Recordset.Name: DAO_RECORDSET_NAME.md dokumentiert getrennte echte Öffnungs-
identität für Tabelle/gespeicherte Abfrage/QueryDef und erste256Zeichen des
ursprünglichen Access-SQL vor Übersetzung/physischem Schreibziel. Klone behalten
Identität; schreibgeschützt383, Argument450, geschlossen3420. Frozen/transformierte
Adapter ohne bewiesene Namenssemantik explizit3251. App/DAO gebaut; einmaliges
bestehendes DAO-Bündel1/1 in3,41s bestanden. Keine neuen CTest-/Katalog-IDs;
456/386/70. Vorstufe8de1ead privat remote/Baum geprüft; letztes geprüftes Paket/
Original weiterhinc67cbd9e vor Abfrageschreiben/Namen. Ohne Zwischenabschluss
weitere allgemeine Anforderungen implementieren; keine Windows-Orakelbehauptung.


DAO physische Abfrageprojektionen: DAO_EDITABLE_QUERY_PROJECTIONS.md dokumentiert
schreibbare eindeutige physische Spaltenlisten/Teilmengen/Umordnung ohne Alias.
Sichtbare Reihenfolge/Herkunft/Typen bleiben echt, unsichtbarer Schlüssel wird
getrennt für Mutation geführt. Nur ausgewählte Defaults gelangen in Entwürfe;
ausgelassene Text-/Date-Defaults bleiben echter DB-Ausführung. Ungefiltertes
AddNew unterstützt; fehlende Generatorprojektion vor Entwurf/Reservierung3251.
App/DAO/QueryModel/Currency gebaut; einmaliges Bündel5/5 in3,98s bestanden.
Review ergänzte Generatorgrenze, nur betroffenes DAO erneut1/1 in3,43s.
Echte SQL-/VBA-/Klon-/Metadaten-/Subset-/Umordnungsbelege. Keine neuen CTest-/
Katalog-IDs;456/386/70. Vorstufe93a8a9f privat remote/Baum geprüft; letztes geprüftes
DEB/Original weiterhinc67cbd9e. Alias-/Join-/Provider-/Orakelparität offen.
Ohne Zwischenabschluss weitere allgemeine Anforderungen implementieren.


Formular-Quell-/Insertidentität: FORM_PHYSICAL_INSERT_IDENTITY.md dokumentiert
tatsächliche main-Qualifikation beim Laden/INSERT/UPDATE/DELETE und TEMP-/Main-
Quellprüfung nach Benutzerereignissen sowie vor Schreibausführung. INSERT nutzt
wie UPDATE echte einzelne RETURNING-rowid einschließlich0; DataEntry/Selektion/
Snapshot/Folgeänderung behalten Identität. Abgewiesene Entwürfe bleiben erhalten.
App/FormCurrent/Layout/Delete/Currency-Insert/Update gebaut. Fixture-Compilezugriff
auf private Methoden durch echte öffentliche VBA-Routen ersetzt, davor keine
Prüfung ausgeführt; einmaliges bestehendes Bündel5/5 in14,42s bestanden.
Echte Tastatur-/VBA-/SQL-/TEMP-Refusal-/Retry-Belege; keine neuen CTest-/Katalog-IDs;
456/386/70. Vorstufe65fa8d1 privat remote/Baum geprüft. Abfrageformulare/Provider/
Quellgenerations-/vollständige Orakelparität offen. Letztes geprüftes DEB/Original
weiterhinc67cbd9e vor späteren Änderungen. Ohne Zwischenabschluss weiterarbeiten.


DAO-Aliasse/Record-Clipboard: DAO_EDITABLE_QUERY_ALIASES.md und
RECORD_CLIPBOARD_INSERT_IDENTITY.md dokumentieren getrennte sichtbare/physische
Feldzuordnung für Schreiben, Generatoren, Defaults/Typen/Herkunft/Properties und
Klone. Gemeinsamer Validator prüft WHERE physisch bei abgeschaltetem DQS-Fallback,
ohne Ausführung; mehrdeutige Qt-Fold-Ausgaben nicht schreibbar. Clipboard nutzt
main-Bindung/Quellprüfung und echte RETURNING-Identität auch0. Gemeinsames Bündel:
5/6 bestanden; DAO-Fixtureannahmen zu unbewiesenen Original-Properties und noch
nicht unterstütztem Currency-CreateField korrigiert, ausschließlich DAO wiederholt,
zuletzt1/1 in3,61s. Modell/Currency/Layout13,78s nicht erneut ausgeführt.
Echte VBA-/SQL-/Counter-/Single-/Date-/BLOB-/Property-/Clipboard-Belege; exakte
Currency-Aliasse an unterstützter physischer Tick-Tabelle. Keine neue Currency-
CreateField-/Orakel-/vollständige Query-Paritätsbehauptung. Keine neuen CTest-/
Katalog-IDs;456/386/70. Vorstufe73328ff privat remote/Baum geprüft. Letztes DEB/
Original weiterhinc67cbd9e; nach Sicherung gesammelten Stand paketieren/Original
prüfen und ohne Zwischenabschluss weitere allgemeine Anforderungen implementieren.


DAO Feld-Schreibfähigkeiten: DAO_FIELD_WRITE_CAPABILITIES.md dokumentiert echtes
Boolean DataUpdatable aus Cursor-/Verbindungsschreibfähigkeit für alle aufgelösten
Felder statt Counter-only-Guard. Readonly-/Snapshot-/berechnete ErgebnisseFalse;
native schreibbare Tabellen/AliasfelderTrue. Native Attributes nutzen validierte
Flags1/2/17 plus32 nur bei Schreibfähigkeit; Original-/Providerflags nicht erfunden.
Property-Objekte, Argument450/Readonly383/Closed3420 und Counter-Edit3164 bleiben.
App/DAO gebaut; einmaliges bestehendes DAO-Bündel1/1 in3,69s bestanden,
echte SQL-/VBA-/native Feld-/Property-/Snapshot-/Readonly-Umschaltungsbelege.
Keine neuen CTest-/Katalog-IDs;456/386/70. Vorstufea23c7009 privat remote/Baum geprüft.
Neues DEBa23c7009/Paketprüfung bestanden, Original1/1 in119,54s bestanden;
alle sieben geschützten Hashes unverändert, bekannte Makrogrenze bleibt.
Dieses Paket liegt vor Feldfähigkeiten; letzter Inventarbeleg4ebc3f6c100/78/22;
keine vollständige Windows-/Provider-/Fehlerreihenfolgenparität behaupten.
Ohne Zwischenabschluss weitere allgemeine Anforderungen implementieren.


Dictionary Container-Objektschlüssel: SCRIPTING_DICTIONARY_RUNTIME.md dokumentiert
echte Collection-/Dictionary-Referenzidentität als Schlüssel, gehaltene Original-
referenzen, Keys-/ForEach-Objekte, Lookup/Item/Remove/Umbenennung ohne Scalar-Default.
Besitzgraph umfasst Schlüssel und Werte; Add/Key-Rename prüfen vor Mutation,
direkte/über Schlüssel und Werte/Collection verlaufende Zyklen verweigern3251.
Nothing/Arrays/beliebige Host-/Projektklassen-/COM-Schlüssel weiter explizit begrenzt.
App/Core gebaut; vor Prüfungen versehentliche gemeinsame Textänderung an Collection
auf Dictionary eingeschränkt und Fixture-API-Aufrufe korrigiert; einmaliges
bestehendes Core-Bündel1/1 in5,70s bestanden. Echte VBA-/Objekt-/Weak-
Referenz-/Snapshot-/Zyklusbelege. Keine neuen CTest-/Katalog-IDs;456/386/70.
Vorstufec59e828 privat remote/Baum geprüft; letztes geprüftes DEBa23c7009, Original
1/1 in119,54s und sieben geschützte Hashes unverändert. Dieses Paket liegt vor
Feldfähigkeiten/Objektschlüsseln. Kein vollständiger COM-/Windows-/Lifetimebeleg;
ohne Zwischenabschluss weitere allgemeine Anforderungen implementieren.


Berichtsgruppen: REPORT_GROUP_METADATA.md dokumentiert echte42 aktive BreakLevel-
Blöcke in neun Originalberichten sowie30/29 BreakHeader/Footer. Ausgabegrenze
übersah bislang diese Access2003-Struktur; aktive Quellen/Break-Abschnitte werden
nun ausdrücklich vor vereinfachtem Druck/PDF abgefangen, Ziel-PDF bleibt erhalten.
Quellparser bewahrt bis10 tatsächliche Gruppen und explizite Eigenschaften;
Report.GroupLevel(Index) liefert lebensdauergebundene gecachte native Objekte mit
ControlSource String, SortOrder/Header/Footer Boolean, GroupOn Integer, Interval
Long und KeepTogether Byte. Qualifizierte As-Bindung/Funktionsrückgabe/named Index
funktionieren; fehlende Defaults3251, Mutation383, Arity450, Null94, Index5.
App/Core/Layout gebaut; frühe Ausgabegrenzenprüfung1/1 in13,67s, anschließend
kombiniertes bestehendes Core/Layout-Bündel2/2 in18,77s bestanden. Echte VBA- /
Objektidentitäts-/Metadaten-/Closed-/PDF-Erhalt-Belege; Originalkorpus bindet seine
neun tatsächlichen Quellen ohne neue private VBA-Ausführung. Keine neuen CTest-/
Katalog-IDs;456/386/70. Gruppendruck/-sortierung/-intervalle/-events/-Mutationen und
Windows-Parität bleiben offen. Vorstufe29ed6fd5 privat remote/Baum geprüft;
letztes DEBa23c7009/Original119,54s/sieben Hashes unverändert liegt vor späteren
Feld-/Container-/Gruppenänderungen. Ohne Zwischenabschluss weiterentwickeln.


Build-Abhängigkeit nativer Typnamen: Die unveränderten drei Lookup-Funktionskörper
stehen jetzt in native_object_types.cpp statt inline im breit eingebundenen Header;
CMake bindet diese Datei in dieselbe Laufzeitbibliothek ein. Künftige Vokabular-
änderungen betreffen eine Übersetzungseinheit. Funktionskörper mechanisch gegen
Vorstufe verglichen, App/Core gebaut, bestehendes Core-Bündel1/1 in5,57s bestanden.
Keine neue Laufzeitfunktion/CTest-/Katalog-ID und keine Laufzeit-Benchmarkbehauptung;
456/386/70. Vorstufee8f54673 privat remote/Baum geprüft. Weiterentwickeln ohne
Zwischenabschluss; letztes geprüftes DEB/Original weiterhina23c7009 vor Folgeänderungen.


Native DAO-Currency-Erzeugung: DAO_NATIVE_CURRENCY_CREATION.md dokumentiert echtes
CreateField(...,dbCurrency) für neue Tabellen und Fields.Append an bestehenden.
CURRENCY_REAL nutzt den bestehenden verlustfreien Vierstellenpfad streng innerhalb
+/-2^39, physische CHECK erzwingt Normalisierung; MetadatenType5/Size8/Required und
Default-Drafts/SQL werden tatsächlich genutzt. Defaults Null oder Dezimalliterale
mit höchstens vier Stellen exakt konvertiert, überschüssige Präzision/Exponent/
Aufrufe/Bereich vor Default-Ersatz3251. DefaultValue zeigt normalisierten Currency-
Text statt binärer SQL-Serialisierungsziffern. Fehlendes Fixture-Include vor Checks
korrigiert; App/DAO/Modell/Currency gebaut. Bestehendes Bündel4/5 bestanden, DAO
fand echten physischen Typkonflikt: Speicherprotokoll verlangt REAL. Deklaration
auf REAL korrigiert, logischer Typ CURRENCY_REAL aus Metadaten; nur DAO erneut1/1 in3,50s
bestanden, echte SQL-/Append-/Constraint-/Metadaten-/VBA-Erzeugungsbelege. Keine
neuen CTest-/Katalog-IDs;456/386/70. Vollständiger Jet-Tickbereich/Decimal/Currency-
Rebuild/Validation/Provider-/Windows-Parität bleibt offen. Vorstufe2af79959 privat
remote/Baum geprüft; letztes DEB/Originala23c7009 vor späteren Änderungen.
Ohne Zwischenabschluss weitere allgemeine Anforderungen implementieren.


Native Currency-Schemaerhaltung: DAO_NATIVE_CURRENCY_SCHEMA.md dokumentiert
kanonische Wiederherstellung tatsächlicher Currency-Defaults für bekannte native
Rekonstruktionen. Default/Required, Feldreihenfolge, Index-Append/Delete und Tabellen-
Umbenennung erhalten vorhandene Werte/rowids; bestehende Äquivalenz-/Digest-/
Savepoint-/Rollback-/Metadatenprüfungen unverändert. Feldumbenennung ebenfalls
belegt. App/DAO gebaut; erste DAO-Fixture erwartete falschen Fehler für angehängtes
Field.Type (korrekt383); tatsächlicher Designerpfad separat3251 geprüft. Nur DAO
wiederholt1/1 in3,71s bestanden, unabhängige SQL-Bitvergleiche/NULL-Rollback/
neue Defaults nach Umbenennungen. Keine neuen CTest-/Katalog-IDs;456/386/70.
Vorstufe90f672b9 privat remote/Baum geprüft. Currency-Typkonversion/Seek/FK-Schlüssel/
voller Tickbereich/Originalschema-/Windows-Parität bleiben offen. Letztes DEB/
Originala23c7009 vor Folgeänderungen; ohne Zwischenabschluss weiterentwickeln.


Currency-Feldeditor/-Typwechsel: CURRENCY_FIELD_DESIGNER.md dokumentiert echte
Currency-Auswahl/Anzeige/Anlage und geprüfte numerische Typwechsel. Vorhandene
Werte und Defaults müssen Ziel-Speicher bitgenau numerisch durchlaufen, Currency-
Defaults erst danach kanonisch normalisiert. Rundung/Bereich/Ziel-Currency mit
Regeln/Meldungen verweigert; DAO Field.Type weiterhin readonly383. GUI zeigt
unterstützten Betragsbereich, Currency-Regelbearbeitung bleibt begrenzt. App/DAO/
GUI gebaut; einmaliges bestehendes Bündel2/2 in22,84s bestanden (GUI19,25s bei
unverändert20s-Watchdog, DAO3,59s). Echte GUI-Anlage/Double-Rückwechsel/Default/
SQL-Werte/Löschen und Backend-Matrix mit NULL/Integralen/Präzisions-/Bereichs- /
Regelrefusal belegt; keine neuen CTest-/Katalog-IDs;456/386/70. Vorstufe6786526d
privat remote/Baum geprüft; letzter Paket-/Originalbelega23c7009 vor Folgeänderungen.
Nach Sicherung gesammelten Stand paketieren/Original/Inventar gebündelt prüfen;
keine vollständige Jet-/Currency-/Windows-/Distributions-Parität behaupten.
Ohne Zwischenabschluss weitere allgemeine Anforderungen umsetzen.


Currency im anfänglichen Tabellenassistenten: CURRENCY_TABLE_WIZARD.md dokumentiert
Auswahl/Typ5-Mapping und numerische Vorprüfung, anschließend exakte bereits
implementierte DAO-Default-Prüfung vor TableDefs.Append. Abgewiesene Präzision
hinterlässt keine Teiltabelle und erhält Dialogeingaben; korrigierter Vierstellen-
Default wird tatsächlich gespeichert. App/GUI gebaut; erster GUI-Lauf überschritt20s; erster Lookup-Retry ebenfalls. Phasenmessung
zeigte nur26ms für den Assistenten. Unveränderliche VBA-Syntaxmuster werden nun
einmalig kompiliert. Bestehendes Core/GUI/DAO-Bündel3/3 in15,18s bestanden, GUI in12,29s bei unverändert20s-Watchdog bestanden; echte Dialog-/Metadata- /
SQL-/Currency-Wert-Belege, keine neuen CTest-/Katalog-IDs;456/386/70.
Vorstufeda72e727 privat remote/Baum geprüft: neues DEB/Paketprüfung bestanden,
Original1/1 in120,33s und Inventar1/1 in46,11s, sieben geschützte Hashes unverändert.
Aktueller Inventarbeleg jetztda72e727100/78/22 (73+5 Ausführungen ohne Access-Orakel,
3 Originalspalten/9 fehlende Abhängigkeiten/10 Jet-/Bindinggrenzen). Makrogrenze
bleibt, kein positiver initialisierter Verkauf/Druck. Dieses Paket liegt vor
Currency-Auswahl im anfänglichen Assistenten. Ohne Zwischenabschluss fortsetzen;
keine vollständige Currency-/Jet-/Windows-/Distributions-Parität behaupten.


VBA-Syntaxmuster-Engpass: VBA_IMMUTABLE_SYNTAX_PATTERNS.md dokumentiert62 feste
Regexmuster, deren unveränderte Bytes/Flags jetzt einmalig initialisiert werden.
Keine Modulquellen/Symbolumgebungen/Argumente/Host-/Variablen-/Ergebniszustände neu
gecacht; Match-Ergebnisse bleiben lokal. Primitive Typen überspringen Objekttabellen,
Scripting-Lookup fremde Präfixe. Optionaler GUI-Phasenlogger enthält nur Namen/Zeit.
20s-Watchdog unverändert, keine vorhandenen Prüfungen entfernt; finales bestehendes
Core/GUI/DAO-Bündel3/3 in15,18s (0,55/12,29/2,35) bestanden. Tatsächliche Etappenzeiten,
keine allgemeine Benchmark-/Paritätsquote; keine neuen CTest-/Katalog-IDs;456/386/70.
Ohne Zwischenabschluss weitere allgemeine Anforderungen fortsetzen.


Native Währungs-Beziehungsschlüssel: DAO_NATIVE_CURRENCY_RELATIONS.md. Typ5 ist
für geprüfte native REAL/CURRENCY_REAL-Paare mit expliziten Felddeklarationen
freigeschaltet; bestehende Required-/Unique-/Schema-/Wertdigest-/FK-Prüfungen
bleiben. Verwaiste Datensätze brechen atomar ab, Retry nach Korrektur funktioniert;
echte Update-/Delete-Kaskaden, Null-Fremdschlüssel, Binärpayload, Reorder und
Relation-Delete im bestehenden DAO-Bündel1/1 in2,44s bestanden. App/DAO gebaut.
Keine volle Jet-Currency-Reichweite/Windows-Parität; F014/F055/F121 bleiben offen,
keine neue CTest-/Katalog-ID;456/386/70 unverändert. Vorstufe33b97df privat und
Remote-Head überprüft. Letztes DEB/Original/Inventar bleibt da72e727; dessen
bekannte Makrogrenze ist kein positiver initialisierter Verkaufs-/Drucknachweis.
Unabhängige schreibgeschützte MDB-Kopiebeobachtung:29 Tabellen/316 Spalten,
81 tatsächliche Defaults vollständig übernommen ohne Literalabweichung;
15 zusätzliche Backenddefaults ausschließlich Boolean. Keine VBA-/Geschäftsdaten
für diese Beobachtung geladen oder ausgegeben. Ohne Zwischenabschluss fortsetzen.


Native Währungs-Gültigkeitsregeln: NATIVE_CURRENCY_VALIDATION.md. Erstellung,
Feldmetadaten, tatsächlicher Feldeditor und explizite verlustfreie Typwechsel
unterstützen Vergleiche/And/Or/Between mit exakt darstellbaren Dezimalgrenzen bis
vier Nachkommastellen im vorhandenen nativen Bereich. Keine stille Rundung von
Regelgrenzen; Exponenten/fünfte Stelle/Funktionen/Bereichsüberschreitung werden
vor Änderung abgewiesen. SQL-Constraints mit ValidationText, atomarer Abbruch
bei Bestandsverletzung und kompatible Currency/Double-Regelerhaltung geprüft.
GUI12,45s bestanden (20s-Watchdog unverändert). Erstes DAO-Fixture benutzte nach
Spaltenaddition eine erwartungsgemäß ungültige alte Feldreferenz; Fixture greift
nun neu auf das tatsächliche Feld zu. Nur DAO neu gebaut/geprüft:1/1 in2,44s.
App/GUI/DAO gebaut; keine neuen CTest-/Katalog-IDs;456/386/70. Frühere pauschale
Currency-Regelsperre ist für diesen begrenzten Umfang ersetzt; keine volle
Jet-/Windows-Ausdrucks-/Currency-Parität. Vorstufeb8fee632 privat/Baum geprüft;
sieben geschützte Hashes unverändert. DEB/Original/Inventar weiterhinda72e727.
Ohne Zwischenabschluss bei weiteren allgemeinen Anforderungen fortsetzen.


Native Currency Index/Seek und Beziehungsoberfläche:
NATIVE_CURRENCY_INDEX_AND_RELATION_UI.md. Gemeinsame Prüfung tatsächlicher
nativer REAL/CURRENCY_REAL-Speicherung für DAO-Beziehungen, Indexnavigation und
GUI-Auswahl bei referenzieller Integrität/1:1. Seek akzeptiert jetzt tatsächliche
Currency-Schlüssel und numerische Schlüssel über den vorhandenen Feldcodec;
Text bleibt abgewiesen, Null-/absteigende Bereichsgrenzen bleiben ausdrücklich.
Echte GUI-Beziehung mit Update-/Delete-Kaskade, DAO Gleichheits-/Bereichssuche,
NoMatch/Bookmark und Zurückweisung bei erhaltenem aktuellen Wert geprüft.
App/GUI/DAO gebaut; ein bestehendes Bündel2/2 in14,37s bestanden:
GUI11,91s (20s unverändert), DAO2,46s. Keine neue CTest-/Katalog-ID;456/386/70.
Vorstufe8a198ce2 privat/Baum geprüft; letztes DEB/Original/Inventar bleibt da72e727.
Keine volle Jet-/Windows-Parität; ohne Zwischenabschluss weiterentwickeln.


Berichtsdruckoptionen: REPORT_PRINT_OPTIONS.md. Tatsächlicher Druckdialog bietet
nun Seitenbereich; Renderer berücksichtigt Originalseitennummern, Vorwärts- /
Rückwärtsfolge,1-100 Exemplare und Sortierung. Aufgezeichnete Quellseiten werden
wiederverwendet, keine erneute Berechnung/Eventausführung je Exemplar.
PDF aus dem Druckdialog übernimmt echte Seiteneinrichtung/Optionen und verwendet
nun temporäres Rendern plus atomaren QSaveFile-Commit. Ungültige Bereiche/Render-
fehler erhalten vorherige Datei; Symlink-/NUL-/aktuelle DB-Ziele abgewiesen.
Native Treiberkopien bleiben Treiberaufgabe, sonst kontrollierte manuelle Kopien;
PDF immer explizit. Mehrfache Ausgabe auf10000 Seiten begrenzt. Selection/current
page weiterhin deaktiviert. Gruppierte/andere nicht unterstützte Berichte bleiben
bewusst gesperrt; kein Original-/Hardware-/Windows-Ausgabenachweis.
App/Layout gebaut; bestehendes access_2003_layouts1/1 in10,79s bestanden. Echter
Dialog/PDF/Text beweist beide Kopiensortierungen rückwärts, nurSeite2 mit2/2,
aktuellen Datensatz wiederhergestellt und Bestandsschutz bei ungültigem Bereich.
Erste Fixturefehler waren fehlende Reaktivierung vor geschütztem Druckbefehl und
fehlende NFKC-Normalisierung der PDF-fi-Ligatur; nur betroffenen vorhandenen Check
wiederholt. Keine neue CTest-/Katalog-ID;456/386/70. Vorstufe009be44e privat/Baum
geprüft. DEB/Original/Inventar nochda72e727; gebündelte Aktualisierung folgt.
Ohne Zwischenabschluss weitere allgemeine Anforderungen fortsetzen.


Aktueller Paket-/Original-/Inventarbeleg: Quellcommit299b946ff0aa559f0e4a4893015a9f71b668f605,
Baumc9fa2d1f78ffd08b0102664f079e88743cb491dc, privat remote/Baum geprüft.
/workspace/work/report-print-options/package/linaccess-0.7.0-Linux.deb erstellt,
check-deb bestanden; bestehendes Paket-Integrationsbündel2/2 in168,22s:
Inventar47,60s, real_kassa_dao_event120,62s; sieben geschützte Hashes unverändert.
Dieses Paket enthält initiale Currency-Auswahl, unveränderliche Syntaxmuster,
native Currency-Erstellung/Regeln/Beziehungen/Indexsuche und Berichtsdruckoptionen.
Neues tatsächliches Inventar299b946f:100/78/22 (73 ohne und5 mit technischen
Parametern ausgeführt, jeweils ohne Access-Ergebnisorakel;3 Originalspalten,
9 fehlende Abhängigkeiten,10 Jet-/Bindungsgrenzen). Originalpfad hält bekannte
Makro-/Tastaturgrenzen; weiterhin kein positiver initialisierter Verkauf und kein
Originalbericht-/Hardwaredrucknachweis. Keine vollständige Parität/Distributions-
Matrix/Windows-Referenz neu behauptet.456/386/70 unverändert. Ohne Zwischenabschluss
weiterprogrammieren; nächster allgemeiner VBA-Umfang wird gesondert implementiert.


Native StrConv: NATIVE_STRCONV.md. Gemeinsamer Evaluator unterstützt dokumentierte
ASCII Upper/Lower/ProperCase, Worttrenner NUL/TAB/LF/VT/FF/CR/Space, benannte
String/Conversion/LCID-Bindung, VBA.StrConv und tatsächlichen String-Rückgabewert.
Nicht-ASCII/Codepages/East-Asia/expliziteLCID/Nichttext/Null-Quellvertrag bleiben
mit3251 ausdrücklich außerhalb dieses Umfangs; keine erfundene Windows-Parität.
Native skalare Objekt-Defaults werden einmal gelesen. Keine Host-/Symbol- /
Ergebniscaches neu. App/Core gebaut; bestehendes core_tests1/1 in0,52s bestanden,
inklusive echter VBA-Funktion, sieben Trenner, Getter-Zähler und Schutzgrenzen.
Keine neue CTest-/Katalog-ID;456/386/70. Letztes geprüfte DEB/Original/Inventar
bleibt299b946f (47,60/120,62s, sieben Hashes unverändert), vor dieser Funktion.
Dokumentations-Vorstufe0a045c94 privat/Baum geprüft. Ohne Zwischenabschluss weiter.


Native Rnd/Randomize: NATIVE_RANDOM_GENERATOR.md. Tatsächlicher projektweiter
24-Bit-Zustand im vorhandenen Modulkontext, geteilt von Formular-/Standardmodul-/
Editor-/Klassen-/nativen Modul-SQL-Hosts; getrennte Projekte/Reset/geschlossene DB
geprüft. Rnd liefert Single[0,1), fehlend/positiv nächste,0 aktuellen Wert, negativ
reproduzierbare Folge. Randomize ist Sub mit optionalem benanntem Number; fehlend/
Empty echter Systemtimer, sonst geprüfte Double-Konvertierung. Native Argumente
werden vor Zustandszugriff skalarisiert/konvertiert; fehlerhafte primitive Eingaben
verändern Zustand nicht. Keine Generatorzustände im Syntax-/Quellcache.
Öffentliche MIT-VBMath-Algorithmusreferenz mit Attribution/Lizenz und Installregel;
kein beobachtetes Access-Zufallsfolgen-Orakel, keine Behauptung vollständiger
volatiler Access-/Jet-Ausdrücke (direkt kalkuliertes/Jet-Rnd ausdrücklich3251).
App/Core/GUI gebaut. Bestehendes Bündel: GUI12,37s bestanden; Core fand fehlenden
benannten Randomize-Sub-Fallback. Binder und qualifiziertes argumentloses Randomize
korrigiert; nur Core neu gebaut/geprüft,1/1 in0,49s bestanden. GUI-20s unverändert.
Keine neuen CTest-/Katalog-IDs;456/386/70. Vorstufecde2c756 privat/Baum geprüft.
Letztes DEB/Original/Inventar ist299b946f, vor StrConv/RNG (47,60/120,62s, sieben
Hashes unverändert). Ohne Zwischenabschluss weiter; nächste Berichtsseitenkopf- /
Seitenfußoptionen anhand dokumentierter tatsächlicher Semantik implementieren.


Berichtsseitenkopf/-fuß: REPORT_PAGE_BANDS.md. Tatsächliche Report-Einstellungen
PageHeader/PageFooter0-3 werden streng gelesen, nativ erstellt und im echten
Berichtsentwurf gespeichert/wieder geöffnet. Dokumentierter Default AllPages0.
Bit1 unterdrückt jeweiligen Bereich bei Berichtskopf, Bit2 bei Berichtsfuß und
erzwingt für diesen eine eigene Seite (dokumentierte Access-Regel). Gemischte
Einstellungen bleiben unabhängig; ohne Berichtsfuß keine künstliche Fußseite.
Pagination/Detailteilung nutzen tatsächliche aktive Bereichshöhen, ursprünglicher
Datensatz/Layout werden restauriert; Ranges/Kopien nutzen weiter Quellseiten.
Keine neue Gruppierungs-/Format-/Print-Event-/VBA-Property-Mutation-Parität.
App/GUI/Layout gebaut. Bestehendes Layout1/1 in10,72s bestanden, tatsächliche PDF-
Texte für alle4 Modi und gemischte Kombinationen, Seitenzahlen/Inhaltsgrenzen/
Restaurierung und ungültige Quellen. GUI hatte alte ForceNewPage1-Erwartung,
Fixture verwendet bewusst0 und neue Fußseitenregel. Vorhandenen Wiederöffnungs-
Check angepasst statt zweiten neuen Dialoglauf; nur GUI wiederholt1/1 in12,07s,
20s-Watchdog unverändert. Keine neuen CTest-/Katalog-IDs;456/386/70.
Vorstufeace05fb7 privat/Baum geprüft. DEB/Original/Inventar noch299b946f, vorStrConv/
RNG/Seitenbereichen. Ohne Zwischenabschluss nächste allgemeine Gruppierung umsetzen.


Native allgemeine Berichtsgruppierung: NATIVE_REPORT_GROUPING.md. Echter Gruppen-
und Sortierdialog mit bis10 unterschiedlichen Quellfeldern, Auf/Ab-Reihenfolge,
Gruppenköpfen und Fußberechnungen; Speichern/Wiederöffnen/Quellgenerierung/Seiten-
vorschau/PDF. AST hält tatsächliche Gruppenbereiche/Controls getrennt mit belegter
Stufenbindung. Renderer verlangt alle7 expliziten Gruppenproperties, GroupOn0,
Intervall1, KeepTogether0; keine fehlenden importierten Defaults geraten. Andere
Druckgrenzen bleiben. Stabile temporäre Sortierung, Null-Gruppen, native Text-/
Zahl-/Date-/Currency-Schlüssel, eindeutige Gruppenaggregate und Berichtssummen;
Zeilen/IDs/Snapshot/Scope/aktueller Datensatz/Layout restauriert. Native Text-
Vergleiche sind keine volle Jet-Locale-Parität; binäre/gemischte/fortgeschrittene
Schlüssel/Intervalle/Repeat/KeepTogether/Events bleiben außerhalb dieses Umfangs.
App/GUI/Layout gebaut: vorhandene GUI12,68s bestanden (20s unverändert), gezielt
Layout nach Fixturekorrekturen1/1 in9,99s. Tatsächliche verschachtelte PDF-Summen/
Counts/Null/stabile Marker/First-Last/Restaurierung, Date-Rand, Grenzen/Dateierhalt.
Fixture nutzte zunächst bestehendes SQL-CStr3085 und fehlenden RecordSource-
Override; nun vorhandene Verkettung und tatsächliche Paketquelle wie die App.
Wichtig: Original-Toolbar-Prüffälle lesen nun bisher ausgelassene Gruppencontrols.
Ohne Datenbank verweigern alle3 Layout-PDFs wegen numerischem Format auf berechnetem
Text13. Exakter bekannter Maskenhash/Quelle fehlt/Typ/Error/Dateierhalt geprüft;
aktuell0 positive Layout-PDFs und3 belegte Verweigerungen. Drei Original-Handler
weiter ausgeführt; keine neuen privaten VBA-Aufrufe, keine Initialisierungs- /
Original-Berichtsparität behauptet. Frühere unvollständige positive Layout-
Beobachtung ist für diesen Stand ersetzt. Keine neuen CTest-/Katalog-IDs;
456/386/70. Vorstufe6578f056 privat/Baum geprüft. DEB/Original/Inventar noch299b946f,
vorStrConv/RNG/Seitenbereichen/Gruppierung; neue gebündelte Paketprüfung folgt.
Ohne Zwischenabschluss weiter. Allgemeine SQL-CStr-Lücke separat beheben.


Paket-/Originalbeleg nach nativer Gruppierung (2026-11-10): Anwendungscode
62d00929; korrigierte öffentliche Rechte der neuen MIT-Lizenzverzeichnisse und
strikte Paket-Allowlist. DEB unter work/native-report-grouping/package-retry/
package/linaccess-0.7.0-Linux.deb: Inhalt/Rechte/Build-ID/SQLite/Metadatenhelfer
bestanden. Vorhandenes Inventar49,42s:100/78/22, unveränderte Kategorien73/5/3/9/10,
kein Access-Ergebnisoracle. Vorhandener Originalablauf nach gezielter Korrektur
des Summenvergleichs1/1 in116,96s bestanden; sieben geschützte Hashes unverändert.
Neue Gruppencontrols hatten Annahmen des alten Prüfers offengelegt: direktes
SQLite umgeht die typisierte Quellausführung; die Regex umfasste auch kombinierte
Summen. Separate schreibgeschützte Quellzeilen-/Aggregatbindung rechnet vollständige
Ausdrücke mit unabhängiger Summenakkumulation nach, ohne Berichtscache/-Gruppenscope.
Keine Anwendungsgrenze gelockert, keine neuen CTest-/Katalog-IDs. Original-Layout-
PDFs bleiben0 positiv/3 belegte Verweigerungen, bekannte Makro-/Tastaturgrenzen
bleiben; keine initialisierte Verkaufs-/Windows-/Gesamtparität behauptet.456/386/70.
Weiter mit allgemeinem SQL-CStr, anschließend weiteren offenen Anforderungen.


Allgemeines natives SQL-CStr (2026-11-10): docs/NATIVE_SQL_CSTR.md. Gemeinsamer
VBA-Auswerter und belegte Feld-/Parametertypen erhalten Boolean/Single/Date/
Currency in typisierten Projektionen; Text einschließlich NUL, Null94, echte
BLOB-Felder13, nicht exakt darstellbare große SQLite-Integer6, Arity450.
Tokenbasierte Übersetzung, keine Geschäftsdatenersatzwerte oder Zustands-Caches.
App/DAO gebaut, bestehendes dao_execution1/1 in2,43s bestanden. Fixture erwartete
zunächst1 trotz vorherigem NativeScale-UPDATE auf2 und verwendete SQLite-X-Hex
statt eines echten Binärfeldes; gezielt korrigiert, keine Laufzeitgrenze gelockert.
Kein kompletter Jet-Locale-/Join-/Subquery-Typnachweis, keine neuen CTest-/Katalog-
IDs;456/386/70. PRIVAT/Baum geprüfte Vorstufe5ec76beb; Paket noch Anwendung62d00929
mit Verpackungskorrektur, vor SQL-CStr. Original116,96s/Inventar49,42s/7Hash gleich.
Ohne Zwischenabschluss weiter mit laufenden nativen Berichtssummen je Gruppe.


Native laufende Berichtssumme je Gruppe (2026-11-10): REPORT_RUNNING_SUM.md.
Echter Designer speichert running_sum_mode1/2; alte Definitionen bleiben2.
Ungültige Werte/Typen und fehlender tatsächlicher Gruppenkopf/-fuß werden
verweigert. Detailpräfixe je innerster tatsächlicher Gruppenstufe einmal in
Drucksortierung berechnet, bei Gruppenwechsel zurückgesetzt; Sortierstufen ohne
Gruppenband gelten nicht als Gruppenwechsel. Seiten/Neuzeichnen/Kopien zählen
nicht doppelt. Bestehende Null-/Detail-/CanGrow-/importierte Metadatengrenzen
bleiben. App/GUI/Layout gebaut: vorhandene GUI12,75s (20s unverändert), gezielt
Layout nach zeilenbegrenzter PDF-Regex1/1 in10,19s. Acht exakte Currency-Präfixe
bei verschachtelten/Null-Schlüssel-Gruppen, richtige Gruppen-/Gesamtsummen,
Speichern/Wiederöffnen/Restaurierung geprüft. Keine neuen CTest-/Katalog-IDs;
456/386/70. Vorstufe1579977d privat/Baum geprüft. Paket noch app62d00929 mit
Verpackungskorrektur5ec76beb; vor SQL-CStr/OverGroup. Original116,96s, Inventar
49,42s,7Hash gleich. Ohne Zwischenabschluss weitere offene Anforderungen umsetzen.


Allgemeine native SQL-Zeichenfolgen (2026-11-10): NATIVE_SQL_STRINGS.md.
CStr/Str/Left/Right/Mid einschließlich $-Varianten und Replace verwenden gemeinsame
VBA-Auswertung und belegte Typen aller Argumente. UTF-16-Positionen, optionale
Argumente, Str-Vorzeichen/Fraction, Null94/Stringvarianten, Replace Start/Count/
Compare; NUL bleibt erhalten, Binär/inexakte Integer verweigert. Typisierte
Argumente erreichen innerhalb der bisherigen Quellenresolver-Grenzen auch
Filter/Gruppier-/Sortierausdrücke; ASC/DESC/mehrere Schlüssel bleiben. Vorhandene
numericOnly-Guards für Date-/Currency-Operatoren nicht gelockert, keine allgemeine
Join-/Subquery-/HAVING-/Jet-Locale-Parität. App/DAO/Currency gebaut: vorhandenes
Bündel2/2 in3,02s (DAO2,95/Currency0,07); vorher StringfamilieDAO2,78s. Reale
DAO-Pfade inkl. Boolean-Gruppierung/Sortierung, Date-Filter und vierstelliger
Currency nahe2^39, UTF-16/Null/Optionen geprüft. Keine neuen CTest-/Katalog-IDs;
456/386/70. Vorstufec2c5be8c privat/Baum geprüft; DEB noch62d00929 mit5ec76beb,
vor CStr/OverGroup/Stringadaptern. Gebündelte neue Paket-/Original-/Inventarprüfung
folgt. Ohne Zwischenabschluss weitere offene Anforderungen implementieren.


Native SQL-Suche/Vergleich/Trim (2026-11-10): NATIVE_SQL_SEARCH_TRIM.md. InStr
mit UTF-16-Positionen liefert Long, StrComp Integer; Trim/LTrim/RTrim inklusive
$-Varianten, belegte Argumenttypen, Null/Varianten und optionale Vergleiche über
gemeinsamen VBA-Auswerter. ASCII-/Locale-/LCID-/Surrogat-/Quellenkontextgrenzen
bleiben ausdrücklich bestehen. App/DAO gebaut, vorhandenes dao_execution1/1
in2,79s (gesamt2,80) bestanden, inkl. vorheriger String-/Kontextprüfungen. Keine
neuen CTest-/Katalog-IDs;456/386/70.
Aktueller geprüfter Paket-/Original-/Inventarbeleg: Quellcommit678a7fd82320afe5a255158a4e5d7f9e26b58944,
Baum0fbe2a242c8c908120f6928ae6290d1bed7d1c30, privat verifiziert. DEB
work/sql-string-adapters/package-proof/package/linaccess-0.7.0-Linux.deb: Inhalt/
Rechte/App-/Helper-Build-ID/SQLite bestanden; vorhandenes Bündel2/2 in163,22s
(Original116,95/Inventar46,27).7 geschützte Hashes unverändert.100/78/22, bekannte
Kategorien73/5/3/9/10, kein Access-Ergebnisoracle; Original-Layout-PDF0 positiv/
3 Verweigerungen, bekannte Makro-/Tastaturgrenzen bleiben. Das Paket enthält
CStr/OverGroup/Stringadapter und liegt vor dieser Such-/Trim-Erweiterung.
Ohne Zwischenabschluss weitere allgemeine offene Anforderungen implementieren.


Wiederholte native Gruppenköpfe/Abschnitts-API (2026-11-10):
NATIVE_REPORT_REPEAT_HEADER.md. Echter Gruppendesigner speichert repeat_header,
Quellgenerator explizites RepeatSection und tatsächliche native Abschnittsnamen.
Planer wiederholt aktive Gruppenköpfe mit belegter Gruppenschlüsselbindung,
berücksichtigt Höhe/Seitenrollen und beendet geschlossene Bereiche. Beliebige
zeilenabhängige Headerausdrücke, andere RepeatSection-Kontexte/Spalten/Unterberichte
bleiben außerhalb dieses Umfangs; bestehende Guards/Dateiatomizität bleiben.
Section-Indizes5..24 für belegte Gruppenpaare, öffentliche erste4 AcSection-
Gruppenkonstanten, Name/Index-Identität, Boolean RepeatSection lesen/schreiben
mit Vorabprüfung/Neuzeichnung, Null94/geschlossen2450. Tatsächliche VBA-Umschaltung
und dreiseitige PDFs an/aus, Identitäten zweier Stufen, Restaurierung/Dateierhalt
geprüft. GUI12,75s bestanden (20s unverändert), nach gezielter Fixturekorrektur
Layout1/1 in10,45s; kein neuer kombinierter2/2- oder Windows-Nachweis.
Keine neuen CTest-/Katalog-IDs;456/386/70. Vorstufe00f8e4ef privat/Baum geprüft.
DEB/Original/Inventar weiterhin678a7fd8 (116,95/46,27s,7Hash gleich), vor Such-/Trim
und dieser Erweiterung. Ohne Zwischenabschluss weitere offene Anforderungen.


Typisierte native SQL-Groß-/Kleinschreibung und Rückwärtssuche (2026-11-10):
NATIVE_SQL_CASE_REVERSE.md. LCase/UCase samt $-Varianten, StrReverse/InStrRev
verwenden gemeinsame belegte Argumenttypen; Boolean-Text, exakte begrenzte
Currency-Ziffern und Long-Ergebnistyp von InStrRev. Bestehende ASCII-/Locale-/
UTF-16-Grenzen bleiben. StrReverse(Null)94, nullable Such-/Case-Varianten und
String-Result-Grenzen korrekt unterschieden; keine Windows-Parität behauptet.
App/DAO gebaut, vorhandenes dao_execution1/1 in2,96s bestanden. Keine neuen
CTest-/Katalog-IDs;456/386/70. Vorstufee64854cf privat/Baum geprüft. Paket weiterhin
678a7fd8 (Original116,95/Inventar46,27s/7Hash gleich), vor Such-/Trim-/RepeatSection/
Case-Erweiterungen; neue gebündelte Paketprüfung folgt. Danach echte modellose
DoCmd.OpenModule-Ausgabe für allgemeine VBA-Entwicklung nachvollziehen/umsetzen.
Keine Zwischenabschlüsse; weitere offene Anforderungen kontinuierlich bearbeiten.


Echtes natives modelloses DoCmd.OpenModule (2026-11-10): MODELESS_MODULE_EDITOR.md.
Formular-/Standardmodul-Hosts, optionale/benannte ModuleName/ProcedureName, belegte
Prozedursuche/Quellposition, Deklarationsansicht, vorhandene Fenster/Entwürfe
wiederverwenden. Gemeinsamer bisheriger Editoraufbau heap-eigen mit wertgebundenen
Callbacks; modaler GUI-Einstieg bleibt. Eigene Quellen bearbeitbar, importierte
Standardquellen hier schreibgeschützt ohne Speichern/Löschen/Run/Debug. Laden
führt Prozeduren nicht aus. Optimistische native Katalogrevision schützt vor
Überschreiben; Close mit Speicheroptionen, echter Abbruch, Busy-/Projektwechsel-/
Reentranz-/QObject-Lebensdauergrenzen. Keine volle VBE-/Klassen-/Formmodul-/Library-/
MDI-/Windows-Fehlergleichheit. App/GUI gebaut, vorhandenes mainwindow_documents
1/1 in13,50s bestanden, 20s unverändert. Tatsächliche VBA-Fortsetzung, Navigation/
Wiederverwendung, Revision/Save/Discard/Cancel, geladene Deskriptoren und unberührtes
technisches Importarchiv geprüft; keine neuen CTest-/Katalog-IDs.456/386/70.
Aktueller geprüfter Paket-/Original-/Inventarbeleg:5c82c36546f744621cb1523e2c1d37d3fef15258,
Baum20206cfc219378aadf8127937ceb3a9d2962cac8 privat verifiziert. DEB
work/sql-case-reverse/package-proof/package/linaccess-0.7.0-Linux.deb bestanden;
Original118,77/Inventar46,44s, Bündel2/2 in165,21s,7Hash gleich.100/78/22 mit
Kategorien73/5/3/9/10, kein Access-Ergebnisoracle; bekannte Original-Grenzen bleiben.
Dieses Paket liegt vor OpenModule. Ohne Zwischenabschluss weiterprogrammieren.


Offene native Modulfenster auswählen/neuzeichnen (2026-11-10):
MODELESS_MODULE_EDITOR.md. SelectObject acModule mit explizitem offenem Namen
über Formular-/Standardmodul-Hosts; keine Wiederöffnung oder Entwurfsersetzung,
kein falscher Datenblattkontext. Verbindungs-/Datei-/Generations-/Modal-/Closing-
Guards. Minimierung entfernen, maximierten/Vollbildzustand erhalten; explizites/
aktives Repaint erkennt modellose Module. App/GUI gebaut; vorhandenes
mainwindow_documents1/1 in13,18s bestanden (20s unverändert); tatsächliche VBA-/
Servicepfade, Zustand/Entwürfe/Neuzeichnen/geschlossenes Ziel geprüft. Keine neuen
CTest-/Katalog-IDs;456/386/70. Vorstufe65844569 privat/Baum geprüft. DEB bislang
5c82c365 (118,77/46,44s,7Hash gleich); neue Paket-/Original-/Inventarprüfung folgt.
Weiter mit der dokumentierten VBA-Bezeichner-Typzeichenlücke; keine Zwischenabschlüsse.


Native skalare Bezeichner-Typzeichen (2026-11-10):
VBA_IDENTIFIER_TYPES.md beschreibt %/&/!/#/@/$ an lokalen Dim-/Static-
Deklarationen, skalaren Parametern, Modulvariablen und Konstanten; kanonische
Variablenbindung, deklarationsbasierter Typabgleich und ByRef-Erhalt. Kein
Quelltext-Strippen, kein Variant-Subtype-Ersatz, kein veränderlicher Rahmen im
Syntaxcache. As-Doppelangabe wird verweigert. Vorhandenes core_tests nach
Korrektur der Konstantensyntax/ByRef-Bindung1/1 in0,53s bestanden. Funktionsnamen,
Array-/ReDim-Deklarationen, Schleifen, implizite Typzeichen und qualifizierte
Projektmember bleiben offen. Keine neue CTest-/Katalog-ID;456/386/70 unverändert.
Nächste Implementierung: Typzeichen an Funktionsnamen samt Signatur-/Rückgabebindung.


Typisierte lokale Funktionsnamen (2026-11-10): VBA_IDENTIFIER_TYPES.md ergänzt.
Function-Suffix/kanonischer Name/Rückgabevariable/Signatur-Typprüfung; direkte
Ausführung mit Typzeichen wird validiert, Metadateninventar behält kanonische
Namen und Quellzeilen. Long-Aufruf vor Klammer erst anhand lokaler Signatur als
Typzeichen gebunden; kein globales Entfernen von &. Vorhandenes core_tests1/1
in0,57s bestanden. Property-/Projekt-/Library-Gesamtnamensräume bleiben offen.
Privater skalarer Vorstand4aa7885f74844d3e41e95c0a2b28bb320f9a14f2,
Baum2cb0d4f0f474adce57e6595b59e3b24ab0fb7831 verifiziert.456/386/70 unverändert.
Weiter mit impliziten skalaren Typzeichen, danach Arrays/ReDim; kein Zwischenabschluss.


Implizite und Array-Typzeichen (2026-11-10): VBA_IDENTIFIER_TYPES.md ergänzt.
Skalare implizite Bindung ohne Option Explicit, Vorrang vor Def-Typen,
Literalbezeichner-Suffixe ignoriert; deklarationsbasierter Typabgleich bleibt.
Lokale/Modul-Arrays, ByRef-Arrayparameter, Elemente, ReDim/Preserve und Erase
verwenden kanonische Namen/Elementtypen. Neue ReDim-Ziele auch mit Option Explicit
prozedurweit kanonisch gebunden. Statischer Funktionszustand identisch bei
Aufrufen mit/ohne Suffix. Bestehender alter statischer Suffix-Guardtest aktualisiert,
weil dieses Verhalten nun implementiert ist; keine Wache pauschal entfernt.
Anwendung/Kern/DAO gebaut; bestehendes Bündel2/2 in3,54s (Kern0,57,DAO2,97s).
Keine neuen CTest-/Katalog-IDs, F040/F122 bleiben offene Vollanforderungen;
456/386/70 unverändert. Funktionsvorstand4e68489b948cf8344467e94ae42c60e4e7f44421,
Baum59d5837f3f8651ee9b31901701fbd9a7c6e9c1e4 privat verifiziert.
Nächste Implementierung: Typzeichen an For/Next, direkten Aufrufen und Stringzielen.


Typzeichen in Anweisungspositionen (2026-11-10): VBA_IDENTIFIER_TYPES.md ergänzt.
For/Next und unterstütztes For Each, direkte lokale Function-Aufrufe mit Call/
Anweisungsform, Mid$-Zuweisung und LSet/RSet auf denselben kanonischen Namen.
Typisierte Arrayquellen behalten Live-Mutationswachen. Lokale Signaturprüfung
nach Namensraumtreffer; eingebaute Dollar-Funktionen bleiben verfügbar (zunächst
bestehende Fixed-String-Prüfung fehlgeschlagen, genau diesen Filter korrigiert).
Double# an Namen/Zahlen verwechselt Anweisungsdoppelpunkte nicht mit Datumliteralen;
Uhrzeitliteral geschützt. Vorhandenes core_tests1/1 in0,53s bestanden.
Kein neues CTest/Katalog-Item.456/386/70. Vorstufe88fdf1f0b8a26c9e6769172dc34e9b7f4141a0e3,
Baum158eeec6c3759bd7be7aae583255916945dfdb24 privat verifiziert; App/Kern/DAO
auf dieser Vorstufe2/2 in3,54s. Nächster Paket-/Original-/Inventarbeleg folgt,
danach typisierte Konstanteninitialisierer und weitere offene native Anforderungen.


Konstanten-/Property-Typzeichen (2026-11-10): VBA_IDENTIFIER_TYPES.md ergänzt.
Getrennte quellgebundene Konstantentypen für Initialisierer/Enums/Optional/
Arraygrenzen/Stringlängen, mit Inline-Kontext und ohne Variant-Subtype-Inferenz.
Get-Namenssuffix und passende Let-Werttypen; kanonische Referenzen/Selektoren.
Schreibschutz typisierter Konstanten vor RHS-Effekten; Loop-Vorprüfung verwendet
kanonische Namen/Suffixtyp vor Def. Bestehendes core_tests1/1 in0,54s bestanden;
gezielte Fixes für Property-Inventarfilter und Syntax von Name&() enthalten.
Keine neue CTest-/Katalog-ID;456/386/70. F040/F042/F043/F122 bleiben breit offen.
Paketbeleg bleibt53e91932 (Original115,97/Inventar46,72s;2/2 in162,70s,7Hash gleich).
Weiter mit deklarationsbasierter Typbindung zwischen Standardmodulen.


Typisierte fremde Standardmodulfunktionen (2026-11-10): VBA_IDENTIFIER_TYPES.md.
Kanonische Hostauflösung, deklarationsbasierter Typabgleich, Ausdruck/Call und
ByRef-/Objektargumentmetadaten über tatsächliche Projektquelle; lokale/Instanz-
Priorität und Sichtbarkeits-/Mehrdeutigkeits-/SQL-Wachen bleiben. Fremde Properties
weiter3251. Statisches Inventar: threadlokal128/16MiB geschätzt, voller Quelltext
als Identität, kein Host-/Katalogzustand im Cache. App/Kern/DAO gebaut; bestehendes
Bündel2/2 in3,42s (0,42/3,00), echte getrennte SQL-Module/Integer2/ByRef/Call und
Verweigerung vor Zustandsänderung belegt.456/386/70, keine neuen IDs.
Vorstand89088dd8402cc62741d30fc1ae7a367f97015eac/Baum24a0fe404fa6fde6c1a7cae13ff38986bcb02b4e
privat verifiziert. Paketbeleg bleibt53e91932. Weiter mit Eval-/Event-Referenzen.


Typisierte Eval-/Event-Funktionsbindung (2026-11-10): VBA_IDENTIFIER_TYPES.md.
Eval: tatsächliche öffentliche Signatur/kanonischer Name, Dollarform; Access-/Jet-
Lexikmodus unverändert. Ereignisse: lokale/öffentliche fremde Signaturen und
Typzeichen, benannte Argumente geordnet, Optional-Missing erhalten, keine
fingierten Adressalias-Bindungen. Private fremde Funktionen/Bibliotheksrückfall
abgegrenzt. App/Kern/DAO2/2 in3,50s (0,41/3,09); nach Sichtbarkeitsabschluss
betroffenes core_tests1/1 in0,42s. Analyzer-Include beim Bau korrigiert.
Fester Katalog456/386/70 unverändert, keine neuen IDs. Weiter offenF040:
Single-!-Namen vor Wortoperatoren/Bang-Feldmehrdeutigkeit, vollständige Bindung.
Vorstand89824952415203a67e181c48fa6935cc96c3442c/Baumece81dd9b9b89507f8be5ad14334501906dddd3d
privat verifiziert. Paketbeleg bleibt53e91932, frischer Paket-/Originallauf folgt.


Single-! vor Wortoperatoren (2026-11-10): VBA_IDENTIFIER_TYPES.md erweitert.
Kontextbezogene Tokennamensbindung gegen echte Variablen-/Ergebnistypen, keine
Ausführung in Metadaten-/Syntaxprüfung; Bang-/Member-/Literalnamen erhalten.
Gebundene Syntax vor dem Körper; kontextbezogene Case/Argument/Bounds/Print-Gänge.
Implizite Single-Typen nach erster textlicher Verwendung, Zweige/Def/OptionExplicit
und vorheriger Variant-Deklaration; Intrinsics/Host-/Projektname nicht verdrängt.
App/Kern/DAO gebaut; finales bestehendes Bündel2/2 in3,69s (0,47/3,23s) bestanden:
Deklarationen/Const/Parameter/Modul/hoisted/implizit/Def/Fn/Get/Event/Call/For/Case/
ReDim, Funktion nur1x und echtes DAO-Feld[And] belegt; ungültiger Funktionsbang
vor Körpereffekten verweigert. Keine neuen IDs,456/386/70. F040 bleibt breit offen.
Paketbeleg780a709 (117,91/47,08s,7Hash gleich). Stand780a709/Baum579ff0b2 privat
verifiziert; neue Single-Erweiterung liegt nach diesem Paket.


Typisierte SQL-Kalenderfamilie (2026-11-10): DATE_RUNTIME.md erweitert.
17 Kalenderfunktionen mit nachgewiesenen Date/Boolean/Single/Currency/integralen
Argumenttypen, tatsächlichen Date/Integer2/Integer4/Text-Ergebnissen, geschachtelten
Ausdrücken und unterstützten Filter-/Gruppier-/Sortierpfaden. Sechs Vergleiche
zwischen nachgewiesenen Date-Werten liefern Boolean -1/0; Datumsliteral-Herkunft
bleibt im Lexer erhalten; Date IS NULL/IS NOT NULL behält gebundene Kinder.
Kein TEXT-Inhalt als Date-Typ geraten. Ungültiges UTF-8 im gemeinsamen Scalar-
Adapter vor ersetzender Dekodierung abgewiesen. Bestehende DAO-/Query-Modell-
Prüfung2/2 in3,47s (3,22/0,25s), App gebaut. DateAdd-Typverlust im Filter und
bestehende SQLite-1-Erwartung korrigiert; bestehender ungültiger UTF-8-Fall belegt.
Keine neuen CTest-/Katalog-IDs.456/386/70 unverändert; F067/F068/F071 sowie
regionale Kalender-/Format-/komplexe Quellenanforderungen bleiben offen.
Kein neuer Windows-/Access-Ergebnisoracle.

Letzter Paketbeleg vor dieser Kalendererweiterung: privat veröffentlichter
Stand9413428d938c27ef19cdeca1dcfc7d5bedaee49c,
Baum61740b61e261d6ea954b3347240193e3be67ac57.
work/single-bang-identifiers/package-proof/package/linaccess-0.7.0-Linux.deb:
Bau/CPack/check-deb bestanden, bestehendes Original-/Inventarbündel2/2 in164,29s
(117,39/46,90s),7 geschützte Hashes unverändert. Bekannte Referenzgrenzen bleiben.


Typisierte SQL-Mathematik (2026-11-10): TYPED_SQL_ARITHMETIC.md.
Sqr/Exp/Log/Sin/Cos/Tan/Atn erhalten nachgewiesene Date/Boolean/Single/Currency-
Argumente und deklarierte Double-Ergebnisse in Projektion/Filter/Gruppierung/
Sortierung. Tatsächliches DAO dbDouble ist7. Native Null94/Domain5/Typ13/
Overflow6 bleiben erhalten; kein Windows-Library-Orakel behauptet. App/DAO/
Query-Modell gebaut; Modell0,25s bestanden, betroffener abschließender DAO-Lauf
1/1 in3,36s. Technische Test-Erwartungen dbDouble und Null korrigiert; keine
Produktregel dafür gelockert. Keine neuen IDs;456/386/70, F067/F068/F071 offen.
Vorstand51b8e75ce1ee55827a3c7da243671a821d645fd5/Baumaa57c12dd3cdc7046009a39076fde7c00a7168f6
privat veröffentlicht und exakt verifiziert. Letzter Paketbeleg weiterhin9413428.


Typisierte SQL-Konvertierungen (2026-11-10): SQL_NUMERIC_CONVERSIONS.md.
CBool/CByte/CInt/CLng/CSng/CDbl erhalten nachgewiesene Date/Boolean/Single/
Currency/integrale Argumente, verschachtelte Ergebnistypen und echte DAO-
Metadaten. Projektion/Filter/Gruppierung/Sortierung geprüft; Null94/Overflow6/
Typ13 erhalten. CByte(True)=255 gemäß bestehendem VBA-Kern, nicht roher
SQLite-Integer1; technische Fehlererwartung dazu korrigiert, Produktregel
unverändert. App/DAO/Query-Modell gebaut; Modell0,29s bestanden, betroffener
abschließender DAO-Lauf 1/1 in3.35s bestanden. Keine neuen IDs;
456/386/70 und offene F068/F071/F105/F122 bleiben. Kein neuer Access-Oracle.
Vorstand82c0e230ad72024b7fd5a1df44fd0dae0ebde989/Baum9c6ffccfacd6d23f2afae0ddd119c295d28cb380
privat veröffentlicht und exakt verifiziert; Paket weiterhin9413428.


Typisierte SQL-Formatierungsfamilie (2026-11-10): NATIVE_FORMAT_RUNTIME.md.
Format/Format$, FormatNumber/Currency/Percent und FormatDateTime erhalten
nachgewiesene Date/Boolean/Single/Currency-Argumente einschließlich Optionen.
Projektion/Filter/Gruppierung/Sortierung sowie genaue int64-Eingabe mit typisierter
Nachkommastellenoption geprüft. Strikte Textmasken13/Nullmasken94 und getrennte
Null-Konventionen erhalten. Bestehende Single-Formatpolitik7 signifikante Stellen
und berechnetes TEXT ohne proven Text/Memo-DAO-Type (3251) bleiben; technische
Testannahmen hierzu korrigiert, keine Produktregel gelockert. Tatsächliche Text-
Variants geprüft. App/DAO/Modell gebaut, Modell0,35s bestanden, abschließender
betroffener DAO-Lauf1/1 in3.48s bestanden. Keine neuen CTest-/Katalog-IDs;
456/386/70 unverändert, F067/F068/F071/F121/F122 weiter offen, kein Access-Oracle.
Vorstandf856c97fa7eadc1df38f6e80fad0784b32bc8160/Baum3bec4863b184a479dc1b92d2ad5ceb0813074925
privat veröffentlicht/exakt verifiziert. Frisches Paket folgt; letzter Beleg9413428.


Paketbeleg für native typisierte Scalar-Familien (2026-11-10):
Privater veröffentlichter Stand2d7fa91e0ee8010c1bcc258bb6d84d0df829372e,
Baum812ba0beafe3a152b033afa6a4f8d61139c53938. DEB unter
work/typed-scalar-families/package-proof/package/linaccess-0.7.0-Linux.deb:
Integrationsbau/CPack/check-deb bestanden; vorhandenes Original-/Inventarbündel
2/2 in164,85s (Original118,63/Inventar46,21s). Alle7 geschützten Hashes
unverändert. Inventar100/78/22, Kategorien73/5/3/9/10 unverändert. Die78
Ausführungen haben weiterhin keinen Access-Ergebnisoracle; bekannte Original-
Makro-/Ablaufgrenzen bleiben. Paket enthält Kalender-/Mathematik-/Konvertierungs-/
Formatierungsbindung, liegt vor den neuen Zeichen-/Alias-/Längenänderungen.
456/386/70 unverändert, keine zusätzliche Vollanforderung geschlossen.


Typisierte Zeichen-/Basis-/Längenfamilie (2026-11-10):
CHARACTER_FUNCTIONS_RUNTIME.md und NATIVE_SQL_STRINGS.md erweitert. Hex/Oct,
Space/String, Asc/AscW/Chr/ChrW sowie Len/LenB binden tatsächliche Date/Boolean/
Single/Currency/Integer/Long-Quelltypen; Integer-/Long-Basisbreiten und tatsächliches
DAO dbInteger für Asc geprüft. Hex$/Oct$/Space$/String$/Chr$/ChrW$ mit echten
SQL-Callbacks und getrennten Null-Regeln ergänzt. Len(Date)=8 erhalten, LenB
nicht-textuell3251 statt ISO-Textinterpretation. Unterstützte verschachtelte
Projektion/Filter/Gruppierung/Sortierung; ASCII/Surrogat-/Output-/Quellgrenzen
bleiben. App/DAO/Query-Modell gebaut, bestehendes Bündel2/2 in4,54s (4,25/0,29)
bestanden. Keine neuen IDs;456/386/70 unverändert, keine Windows-Oracle-Behauptung.
Vorstand2d7fa91e0ee8010c1bcc258bb6d84d0df829372e/Baum812ba0beafe3a152b033afa6a4f8d61139c53938
privat verifiziert; dessen Paketprüfung/Originalbündel164,85s/7Hash gleich siehe
Paketbeleg oben. Neue Zeichen-/Längenbindung liegt nach diesem Paket.


Typisierte skalare SQL-Finanzfamilie (2026-11-10): NATIVE_FINANCIAL_FUNCTIONS.md.
FV/PV/Pmt/NPer/Rate, SLN/SYD/DDB und IPmt/PPmt erhalten tatsächliche Date/
Boolean/Single/Currency/integrale Argumente; deklarierte Double-Ergebnisse und
DAO dbDouble7 sowie unterstützte Verschachtelung/Filter/Gruppierung/Sortierung
geprüft. Unabhängige Nullzins-/Zweiperioden-Abschreibungswerte, Single-Rundung,
Null94, Text13 und negative Asset-/Timinggrenzen3251 belegt. Vorhandene Solver/
Domain-/Iterations-/Präzisionsgrenzen unverändert, keine SQL-Arrays hinzugefügt.
App/DAO/Query-Modell gebaut; bestehendes Bündel2/2 in4,19s (3,89/0,30) bestanden.
Keine neuen IDs,456/386/70 unverändert; kein neuer Access-Oracle.
Vorstandf8bb8a6b1016b5224ff19b86652c89f8b9e27885/Baum01cc89b8c8d4424e7778caba70eeb6a4b9685aa0
privat veröffentlicht/exakt verifiziert. Letzter Paketbeleg2d7fa91 (164,85s,
7Hashes gleich) liegt vor Zeichen-/Längen-/Finanzbindung.


Typisierte SQL-Farb-/Partition-Familie (2026-11-10): SQL_RGB_RUNTIME.md,
QBCOLOR_RUNTIME.md, NATIVE_PARTITION.md. RGB/QBColor/Partition binden reale
Date/Boolean/Single/Currency/integrale Quelltypen; Farbprojektionen INTEGER4,
tatsächliches DAO dbLong4/Long-Variant. Partition-Padding/Buckets, Date-Kanäle,
Boolean-Negativ-/Null-/Textfehler sowie Filter/Gruppierung/Sortierung geprüft.
Native Palette/Rundung/Null-/Extrem-/Quellgrenzen bleiben, keine Windows-Oracle-
Behauptung. App/DAO/Query-Modell gebaut; bestehendes Bündel2/2 in4,21s
(3,91/0,30) bestanden. Keine neuen IDs;456/386/70 unverändert.
Vorstande9051bce801df5dd03097e54512f435e80d2ac8e/Baum4475bc5f5bf527cb138642b11a99b65c5711c31d
privat veröffentlicht/exakt verifiziert. Letzter Paketbeleg2d7fa91 (164,85s,
7Hashes gleich) liegt vor Zeichen-/Längen-/Finanz-/Farb-/Partition-Bindung.


Typisierte SQL-Rundung/Normalisierung (2026-11-10): NUMERIC_FUNCTIONS_RUNTIME.md,
ROUND_RUNTIME.md und SQL_NUMERIC_CONVERSIONS.md. Abs/Int/Fix/Sgn/Round/CCur
binden nachgewiesene Quelltypen und geprüfte polymorphe Resultate; Int/Fix/Sgn
als native SQL-Callbacks ergänzt. Integer/Long-Abs-Overflow6, Currency-/Single-
Resultate mit tatsächlichen DAO-Typen, Null-/Präzisionsgrenzen und rohe exakte
SQLite-Abs-int64-Berechnung erhalten. Round(Single) bleibt3251. Unterstützte
Projektionen/Klauseln geprüft; sonstige Date/Currency-Operatorgrenzen bleiben.
Tatsächlichen Fehler13 bei CCur(Date) behoben: MS-VBAL5.5.1.2.3/6.1.2.3.1.3
fordert den Date-Double-Serienwert durch den vorhandenen Currency-Konverter.
Positive/negative Bruchteile im Kern und Date-Feld in SQL geprüft. App/Kern/DAO/
Query-Modell/Currency-Ziel gebaut; Modell0,31s und Currency-Ergebnisse0,07s
bestanden. Nach Date-Fix betroffenes Kern-/DAO-Bündel2/2 in4,47s (0,40/4,07)
bestanden. Keine neuen IDs,456/386/70 unverändert, kein Access-Oracle behauptet.
Vorstand3d2ff1f12bc33fa3b103cc8633f79bef7b5d27d6/Baum7da7aa9200277987742a3a1a6d06b5d7bc819051
privat veröffentlicht/exakt verifiziert; letzter Paketbeleg2d7fa91 liegt davor.


Frischer Paketbeleg für native SQL-Scalar-/Rundungsbindung (2026-11-10):
Privater Stand9f73a8c5f86f6cb1a28f2e86989c27f264f49245,
Baumebea651f8afd1f1db9e6689c4db3514f88e8cafc. DEB unter
work/sql-normalization/package-proof/package/linaccess-0.7.0-Linux.deb:
Integrationsbau/CPack/check-deb bestanden; bestehendes Original-/Inventarbündel
2/2 in167,90s (Original120,56/Inventar47,34s);7 geschützte Hashes unverändert.
Inventar100/78/22 und Kategorien73/5/3/9/10 unverändert. Weiterhin kein Access-
Ergebnisoracle für78 ausgeführte Abfragen, bekannte Original-/Makrogrenzen bleiben.
Paket enthält Zeichen-/Längen-/Finanz-/Farb-/Partition-/Normalisierungsbindung
und Date-zu-Currency-Korrektur; liegt vor den neuen SQL-Typabfragen.
Fester Katalog456/386/70 unverändert; keine Vollparität behauptet.


Native SQL-Typabfragen (2026-11-10): VARTYPE_RESULT_RUNTIME.md.
IsNumeric/VarType/TypeName als echte native SQL-Callbacks mit geprüfter Quell-
Typbindung und Boolean/Integer/Text-Ergebnisprojektionen. Date/Boolean/Single/
Currency/Integer und verschachtelte Byte/Long/Double-Konvertierungen, tatsächliche
DAO Boolean/Integer-Typen, Null-Laufzeittyp, Text statt geratenem Date sowie
Projektionen/Filter/Gruppierung/Sortierung geprüft. Boolean-Literalbindung bleibt
auch in Klammern erhalten; quoted Feldnamen unverändert. Grenzen unproven
Zahlen/BLOB/Object/Array/Locale sowie gesondertes Access-Eval-Profil bleiben.
App/DAO/Query-Modell gebaut; bestehendes Bündel2/2 in4,50s (4,14/0,35) bestanden.
Keine neuen IDs;456/386/70 unverändert, kein neuer Windows-Access-Oracle.
Vorstand9f73a8c5f86f6cb1a28f2e86989c27f264f49245/Baumebea651f8afd1f1db9e6689c4db3514f88e8cafc
privat veröffentlicht/exakt verifiziert, Paketbeleg167,90s/7Hashes gleich oben.
SQL-Typabfragen liegen nach diesem Paket. Kontinuierlich weitere Anforderungen
implementieren; kein Zwischenabschluss oder Abschlussanruf.


Native SQL-Boolean-Ausdrücke (2026-11-10): NATIVE_TYPED_LOGIC.md.
Nachgewiesene Boolean/NULL And/Or/Xor/Eqv/Imp/Not und Boolean-Vergleiche verwenden
native Werte (-1 auch bei physischerBOOLEAN1), deklarierte Boolean-Ergebnisse
und tatsächliche DAO Boolean/dbBoolean. Imp/Eqv/Xor/Or/And-Rangfolge und Not
unter Vergleichen/über And gebunden. Gewöhnliche Vergleichsprojektionen sowie
Gruppier-/Sortierausdrücke kodieren -1/0/NULL; einfache numerische Predicate und
predicate And/Or/Not behalten SQL-Struktur/den vorhandenen ID-Indexpfad.
Date/Currency-Operator-/Quellgrenzen, gemischte Bitlogik und ungemessene Imp/Null-
Windows-Semantik bleiben offen. Vier vorhandene Access-Null-And/Or-Beobachtungen
beibehalten; neue Fälle sind native technische Belege, kein neuer Oracle.
App/DAO/Modell gebaut; Modell0,31s bestanden. Fehler im strikten CBool(NULL)-Weg
für neue Vergleichsprojektionen behoben durch eigenen NULL-erhaltenden0/1-
Encoder; CBool(NULL) selbst unverändert94. Betroffener finaler DAO-Lauf1/1 in4,63s
bestanden. Keine neuen IDs,456/386/70 unverändert. Gezielter UI-Lauf folgt.
Vorstandc4c4af5945328bb99d3ce8254ab4952936d06a6e/Baum86fb8ce6124c1aaf788931a01b79a765d7cc6dec
privat veröffentlicht/exakt verifiziert; letzter Paketbeleg9f73a8c5 liegt vor
SQL-Typabfragen/Boolean-Ausdrücken (167,90s,7Hashes gleich).


Gezielter UI-Nachweis für SQL-Boolean-Ausdrücke (2026-11-11): App-/Desktop-
Abhängigkeiten und vorhandenes Hauptfensterziel gebaut. Query-Designer-Referenz
kodiert Vergleichsbuckets nun unabhängig mit CASE -1/0 statt rohem SQLite1/0;
Gruppier-/Summen-/HAVING-/Sortiervergleich unverändert vollständig geprüft.
Vorhandenes mainwindow_documents1/1 in15,10s (Test15,09) bestanden, unveränderter
20s-Watchdog. Keine neue CTest-/Katalog-ID;456/386/70. Veröffentlichter Stand
a599c2dd7a9c79f8eb1529de54bc288946f6d120/Baumaf0c0433d29584711ece0152e2467214fbd07b0f
privat/exakt verifiziert. Kein neuer Windows-/Gesamtparitätsnachweis.


Primitive SQL-Prädikatergebnisse (2026-11-11): NATIVE_TYPED_LOGIC.md.
LIKE/NOT LIKE, IN/NOT IN, BETWEEN/NOT BETWEEN und IS NULL/IS NOT NULL behalten
Boolean -1/0/NULL in unterstützten Projektionen, Verschachtelung, Gruppierung
und Sortierung; echte Predicate behalten SQL-Operatorstruktur. BETWEEN-
Bereichs-AND von logischem AND getrennt, gebundene Bounds/IN-Listen-Kinder
und negierte linke Ausdrücke erhalten. IS NOT NULL wird vor scalar Not gebunden;
vorher mögliche falsche Transformation IS Not(Null) korrigiert. Tatsächliche
DAO Boolean-Typen, NULL-Fälle, Präzedenz, Gruppierung/Sortierung/Filter geprüft.
Unveränderte Pattern-/Collation-/Membership-/Range-Coercion-/Date-Currency-/
Quellgrenzen; keine pauschale Jet- oder Oracle-Behauptung. App/DAO/Modell gebaut;
ein bestehendes Bündel2/2 in5,15s (4,75/0,39) bestanden. Keine neuen IDs;
456/386/70 unverändert. Vorstandf66aaa3e21e1bc19d766f25483e451b0a8ce9c26/
Baum3e2c56be109c5eebc94ff9ead0087f108d17c47a privat/exakt veröffentlicht.
Letzter Paketbeleg9f73a8c5,167,90s/7Hashes gleich; spätere SQL-/UI-Inkremente
darin nicht enthalten. Ohne Zwischenabschluss weiterimplementieren.


Native SQL-Wertzustandsfunktionen (2026-11-11): NATIVE_TYPED_LOGIC.md.
IsNull/IsEmpty/IsError als echte native SQL-Callbacks mit vorhandener Quell-
Typbindung, Boolean-Ergebnissen und tatsächlichem DAO dbBoolean. SQL-Null/
leerer String/gewöhnliche Werte, native Typabfragen, Filter/Gruppierung/Sortierung
und Arity450 geprüft. IsError ist kein SQL-Exceptionfänger: vorheriger1/0-Fehler
bleibt11. Keine zusätzliche Object/Array/Error-SQL-Kodierung oder Windows-Oracle-
Behauptung. App/DAO/Modell gebaut, ein bestehendes Bündel2/2 in5,03s (4,70/0,33)
bestanden. Keine neuen IDs;456/386/70 unverändert. Vorstand549adc3944b32c476de3d80206c7c00f882641a5/
Baum3982b798dc42fa3e8333560a6fd7cae3d5986c42 privat/exakt veröffentlicht.
Letzter Paketbeleg9f73a8c5 (167,90s/7Hashes gleich) liegt vor späteren SQL-Blöcken.


Native SQL-Nz-Bindung (2026-11-11): NAMED_BUILTIN_RUNTIME.md.
Echte SQL-Callbacks1/2 Argumente, ausgelassener SQL-Ersatzwert explizit leerer
String gemäß vorhandener Microsoft-Referenz; VBA-Nz-Empty unverändert. Eager-
Fallbackauswertung geprüft (Nz(1,1/0) bleibt11), keine COALESCE-Lazy-Substitution.
Nachgewiesene gleiche Date/Boolean/Single/Integer/Currency-Zweige bzw. passender
literalNULL-Zweig behalten Metadaten/tatsächliche DAO-Typen. Gemischte/unbekannte
Zweige bleiben ohne geratenen Ergebnistyp; vollständige heterogene Variant-
Identität in SQL bleibt offen. Null/default/retained-Werte, Typabfragen, Filter/
Gruppierung/Sortierung und Arity450 geprüft. App/DAO/Modell gebaut; ein bestehendes
Bündel2/2 in5,10s (4,72/0,37) bestanden. Keine neuen IDs;456/386/70 unverändert.
Vorstandcc9050249ecfb6c6c784b984b9ac5d23c76c7c79/Baumcd340f5b39fbb7d6dabeaa679abb69af252cccb1
privat/exakt veröffentlicht. Letzter Paketbeleg9f73a8c5 liegt vor späteren SQL-
Blöcken; gemeinsamer frischer Paket-/Originallauf folgt. Kein Gesamtabschluss.


Originale RIGHT-JOIN-Regression behoben (2026-11-11): NATIVE_TYPED_LOGIC.md.
Drei Originalabfragen und OnActivate liefen nach SQL-Boolean-Erweiterung in3035:
20M unveränderte Operationsgrenze, ca.4M Vollscans in130ms, kein Zeitlimitproblem.
Reiner boolean_result-Callback fehlte in geschlossener RIGHT-JOIN-Funktionsliste;
exakten Namen ergänzt, unbekannte Funktionen/Shapes weiterhin ausgeschlossen.
Anschließend13 bei erneutem Binden bereits kodierten SQLs: Encoder für0/1/-1/NULL
idempotent, bereits kompilierten Bool-Ausdruck nicht erneut wachsen lassen.
Kein Budget/Watchdog/Abdeckungsschwellenwert erhöht. Neues Szenario innerhalb
bestehendem Query-Modell-Ziel:2000 technische rechte Schlüssel, vollständige
matched/unmatched-Ergebnisse und tatsächliche Vollscans<10000. DAO/Modell2/2
in5,26s (4,93/0,33), betroffene Originalziele final2/2 in177,19s (Inventar51,53/
Original125,65) bestanden;100/78/22, Kategorien73/5/3/9/10 wiederhergestellt.
Alle7 geschützten Hashes unverändert. Diagnoseberichte vor unveränderter
Abdeckungsassertion sichern; Aktivierungsfehler nur Status/Error/Zeile, kein
privater SQL/VBA-Dump. Temporäre Budgetdiagnostik vollständig entfernt.
F067 um bestehende Detailverweise ergänzt;456/386/70 und IDs unverändert.
Vorstand57fb0600e7b79800c1f3de412f82de975b9701c5/Baume7ce83499342843ed3c1041dc88175eb4273555f
privat veröffentlicht/exakt verifiziert. Dessen DEB baute/checkte zwar, bestand
Originalziele aber vor Fix NICHT: kein gültiger neuer Paketbeleg. Letzter komplett
geprüfter DEB bleibt9f73a8c5 (167,90s/7Hash gleich), bis gemeinsamer Refresh folgt.
Kein Gesamtabschluss: vorbereitete IIf/Choose/Switch-Bindung weiter umsetzen.


Native SQL-Auswahlfamilie (2026-11-11): NAMED_BUILTIN_RUNTIME.md.
IIf/Choose/Switch erhalten nachgewiesene Date/Boolean/Single/Integer/Currency-
Argumente und kompatible Zweigmetadaten samt tatsächlichen DAO-Werten. Literal-
NULL-Zweige, fehlende Auswahl, Null-Bedingungen, Date-Bedingung, eager Fehler/
Arity und Verschachtelung/Filter/Gruppierung/Sortierung geprüft. Unbekannte/
gemischte Zweige ohne geratenen Typ. Roher IIf-int64-Transport bleibt erhalten,
wenn nur Boolean-Bedingung bewiesen; typisierter Transport mit bisherigen Grenzen.
IIf-Arity im nativen Callback statt SQLite-Vorabweigerung:450 statt3146. Falsche
technische Switch-Null-Erwartung korrigiert: bestehender Kern überspringt NULL,
Produktregel unverändert. Exakt bekannte reine typed-IIf/Choose/Switch-Callbacks
für RIGHT-JOIN-Adaption zugelassen; bestehender2000-Schlüssel-Fall prüft nun
rohen und typisierten IIf mit Ergebnissen und echten Vollscans<10000.
App/DAO/Modell gebaut; Modell0,38s bestanden; final betroffener DAO-Lauf1/1 in5,06s
(Test5,03) bestanden. Keine neuen IDs/Windows-Oracle-Behauptungen;456/386/70.
Vorstand49e695d2bb2bb134f2c24dff2e4be2f35f4ea33c/Baum72229b394f8c28b5df933d988a9a0e38b6a74ed6
privat/exakt veröffentlicht (Remote-Lesebestätigung verzögert, nachgeprüft/adoptiert).
Dessen Original-Fix2/2 in177,19s und7Hashes gleich; frischer gemeinsamer Paket-
/Originallauf folgt. Alter57fb-DEB bleibt nicht vollständig geprüft.
