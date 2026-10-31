# Bounded native OpenForm and option-group execution

DoCmd.OpenForm and Application.DoCmd.OpenForm accept FormName, View,
FilterName and WhereCondition as positional arguments. Empty positions use the
supported defaults. acNormal (0) opens an imported form, case-insensitively.
A nonempty WhereCondition is applied through the real record-source SQL before
Open and the initial record events. It is visible through Filter/FilterOn.
The existing native preview remains modal; this is not Access window-mode parity.

The parser preserves argument positions; syntax acceptance is separate from
runtime support. Unsupported views, nonempty FilterName and additional
DataMode/WindowMode/OpenArgs positions raise 3251. Null arguments raise 94,
missing imported targets raise 2102 and cancelled opening raises 2501.
Initialization failures and unavailable bound sources propagate rather than
returning successful opening. The same argument validation serves both dispatch
paths. Generic headless hosts reject a nonempty WhereCondition if they cannot
apply it.

## Reproduced original failure and fix

The unchanged original wizard handler in module
`6e8e0a0e41e1a1768cff4a0b79a07205c7b0f799eb8b465af5397843ce048ed8`
reached its four-position OpenForm call, but no form appeared. Its error handler
showed a message instead. The target form
`4480943bd1372986a103cc2400deca5b3a0305df6dd7fd8b77f701199363b167`
failed in Open at source line 102 and Current at line 84 with error 2465.
Both read an OptionGroup which was parsed but had no native runtime widget/value.

The parser now retains numeric DefaultValue, OptionValue and enclosing group
indices. Native QGroupBox/QButtonGroup choices use exclusive checkable buttons,
including the original nested ToggleButtons. Groups expose a numeric Value,
ControlType 103 and Null clearing. Programmatic selection updates real buttons;
user selection updates the group before supported update/click bindings. Bound
existing table records load and save their actual numeric values.

The original handler now opens that unchanged target and completes without an
unexpected error dialog. Its original Open and Current no longer fail on the
group read. This proves a specific execution path, not the whole application.

## Tests

- `core_tests`: original single-name host contract, omitted argument positions,
  supported defaults and explicit unsupported headless filtering.
- `access_2003_layouts`: real native dialogs, both dispatch paths, actual restricted
  row/value and Open-time Filter visibility; missing target, unsupported view,
  named filter/options, malformed SQL, raised initializer error and cancellation.
- Native option tests use a real mouse click, observe the new group value inside
  VBA, independently query its AfterUpdate SQL effect and verify persisted table
  data after closing. Default, Null, selection and rejected fractional values
  have separate checks.
- `real_kassa_dao_event`: all original artifact hashes are checked; a synthetic
  caller opens an unchanged original layout over actual MDB records with a
  WhereCondition, independently checks row counts and bound cells. This binding
  test intentionally overlays empty CodeBehind within a rolled-back test-copy
  transaction. Separately, the unchanged original root handler and target
  CodeBehind execute together without that overlay.
- Native modal callbacks have three-second deadlines; an unopened target or an
  unexpected original error dialog cannot count as successful execution.

## Remaining boundaries

No existing-open-instance reuse, modeless/MDI lifetime equivalence, designer or
datasheet view, named filter query, data/window modes or OpenArgs is claimed.
Only currently implemented SQL criteria and record-source binding are available.
Full reentrant Activate/error semantics require further work and an Access oracle.

Group defaults support numeric integer literals only. Expression defaults fail
explicitly when read. Bound new-record defaults, individual grouped child Value
assignment, fractional coercion and cancelled-update selection rollback remain
unsupported/incomplete. No exact Access rendering or event-order equivalence is
claimed. Unsupported individual assignments/fractional values raise 3251.
All mutations in regression tests affect disposable copies only.
