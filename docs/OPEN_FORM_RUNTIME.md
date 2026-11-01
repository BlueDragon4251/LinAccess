# Bounded native OpenForm and option-group execution

DoCmd.OpenForm and Application.DoCmd.OpenForm accept FormName, View,
FilterName and WhereCondition as positional arguments. Empty positions use the
supported defaults. acNormal (0) opens an imported form, case-insensitively.
A nonempty WhereCondition is applied through the real record-source SQL before
Open and the initial record events. It is visible through Filter/FilterOn.
The native target remains open after OpenForm returns, so the following VBA can
read/write Forms(target) and save its actual records. A uniquely open instance
in the same project is reused, including a supported new WhereCondition. Name
retains the stored catalog spelling rather than the caller's capitalization.
Imported Modal controls Qt application modality; it does not suspend VBA like
acDialog. Full Access window-mode/document lifetime parity is not established.

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
user selection updates the group before supported update/click bindings. Click
is dispatched to the group only, not its individual choices (see
ALLOW_EDITS_RUNTIME.md for the Microsoft contract correction). Bound
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
- The post-OpenForm Forms access regression failed before the modeless change.
  Tests now inspect the real persistent native window after VBA returns, check
  instance identity and retained CodeBehind state on reuse, and independently
  verify SQL writes made by the continuing caller.
- Immediate Screen.ActiveForm/ActiveControl after OpenForm/SetFocus are tested
  before processing delayed window-manager activation. Screen uses the logical
  active native document and its focused widget, not only OS activation timing.
- Project closing saves/validates independent opened documents and respects Unload Cancel.
  Closed object references invalidate. A self-closing child can show a real
  MsgBox before its procedure ends: physical QObject deletion waits until the
  outermost native VBA call unwinds, protecting nested Qt event loops.
- Both original calendar double-click bindings remain unchanged. Their tests
  now wait for the actual user click after the opener returns and verify the
  original date write-back and real MDB date-query count. The original calendar
  has no Modal property; the harness no longer assumes every target is modal.

## Remaining boundaries

Native opened forms now belong to the project window (or an independent top-level
window), and survive closing/destruction of their invoking form. Project shutdown
and Quit validate native documents; see [lifecycle evidence](DOCUMENT_LIFETIME_RUNTIME.md).
Bounded same-project active-document commands and keyboard routing are tested
(ACTIVE_DOCUMENT_COMMANDS.md). Full MDI, remaining activation/DoCmd semantics,
designer/datasheet views, named filters, data/window modes (including acDialog)
and OpenArgs remain open.
Only currently implemented SQL criteria and record-source binding are available.
Full reentrant Activate/error semantics require further work and an Access oracle.

Group defaults support numeric integer literals only. Expression defaults fail
explicitly when read. Bound new-record defaults, individual grouped child Value
assignment, fractional coercion and cancelled-update selection rollback remain
unsupported/incomplete. No exact Access rendering or event-order equivalence is
claimed. Unsupported individual assignments/fractional values raise 3251.
All mutations in regression tests affect disposable copies only.

Modeless increment validation: Release build, all 18 CTests passed (75.58 s),
fresh .deb payload and executable Build ID checked, four original SHA-256 checks
passed. GUI-disabled runtime code is unchanged from the preceding verified core
build. Debian 12/Ubuntu remote CI results remain unverified here.
