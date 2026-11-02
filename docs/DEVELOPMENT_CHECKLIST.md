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
- [x] Reproduce 9/29 even with CLI 1.0.0 by changing only its loaded MDB library; tie all 20 rejected schemas to standalone PK clauses; current full test passes all four version/library combinations.
- [ ] Independently reproduce/explain the remaining historical intermittent EOF reports.
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
