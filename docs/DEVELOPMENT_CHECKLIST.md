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
- [ ] Per-artifact and per-procedure capability/execution classification, including failures.
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
