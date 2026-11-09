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
