# Native With, keyboard delivery and make-table execution

This is a bounded runtime increment, not complete VBA or Kassa compatibility.
Original MDB/Bridge bytes remain unchanged; actions use disposable database copies.

## Implemented and verified

- `With` evaluates its object once, supports nested relative members and keeps
  the correct object through handled errors/Resume and Exit For. Native DAO
  Edit/Fields/Update effects are independently checked with SQL. Relative members
  outside a With block are rejected. CodeContextObject has the native instance
  context; full Access macro/class context remains open.
- `SendKeys` delivers actual Qt key events, synchronously for Wait=True and queued
  for False. Literal Unicode, supported named keys, modifier groups, repetition
  and escapes are preflighted before effects. Input stays in native documents
  of the invoking project (ACTIVE_DOCUMENT_COMMANDS.md). Unsupported key layouts/tokens fail explicitly; destroyed/hidden windows
  cannot receive queued events. This is not Windows desktop input injection.
- Native GoToControl resolves literal control names and selects editable text.
  Unbound TextBox.Requery preserves its value. Input is checked in actual widgets,
  including independent SQL persistence for bound synthetic records.
- Calculated controls retain their displayed typed result, including Currency,
  instead of rerunning an expression on every read. Input changes, record
  navigation and explicit calculated-control Requery recalculate. Parent values
  refresh after a changed child Requery snapshot; unchanged snapshots stay cached.
  Native Recalc is available; full Access dependency/event scheduling remains open.
  See READONLY_FORM_REQUERY_RUNTIME.md.
- Dotted generated control names resolve a unique punctuation-to-underscore
  spelling; exact names win and collisions fail. Bang/Controls.Item remain literal.
  Standard-module Me.Requery/SetFocus cannot bypass the instance scope guard.
- Option Compare Binary/Text/Database is scoped to the executing module for
  relational string operators. The Qt form host uses Unicode case-insensitive
  comparison for Text/Database; the generic host supports ASCII folding and
  rejects unresolved non-ASCII comparisons. Jet database/locale collation,
  Like/StrComp/InStr comparison options and full locale ordering remain open.
- Local `SELECT ... INTO [table] FROM ...` runs as a native make-table action.
  QueryDef parameters and RecordsAffected, existing-target error 3010 and nested
  transaction rollback are tested. TOP/RIGHT JOIN materialization and external
  database destinations remain explicitly unsupported. SQLite projection
  affinities are retained; full Jet AutoNumber, types, keys, defaults, indexes
  and validation propagation are not claimed.

## Original artifact evidence and remaining failure

The unchanged module with SHA-256
`6e8e0a0e41e1a1768cff4a0b79a07205c7b0f799eb8b465af5397843ce048ed8`
executes all three keyboard handlers containing physical lines 229, 244 and 1943.
The test supplies normal search input using actual imported records, observes
real queued events and independently checks text in all three actual receivers.
The third handler changes focus before queued delivery; the test follows that
actual focus instead of assuming the earlier GoToControl target.

Its original make-table statement at line 2143 is also executed independently
inside a savepoint. All 14 projected column names/counts and row contents are
compared with independent SQL. This fixture workflow has no active sale, so that
original snapshot is empty; positive/NULL row copying and parameter cases are
covered separately by the DAO regression. No positive sale snapshot is claimed.
The complete third handler now reaches its end. Reproduced blockers fixed along
this path were generated control alias error 2465, unsupported SELECT INTO error
3075 and case-sensitive table existence lookup despite Option Compare Database.

The later asynchronous AfterUpdate from the first two keyboard handlers still
fails with error 3075 at original line 1023 because this probe has no active-sale
identifier. The regression explicitly asserts both failures, separately from the
three completed keyboard handlers. This is **not** a successful checkout/sale
regression. A subsequent increment implements bounded modeless OpenForm
continuation (OPEN_FORM_RUNTIME.md). A valid full sale setup, independent document
lifetimes, actual report printing and Access event/result comparison remain open.

At this increment the syntax inventory accepted 185/187 procedures. Later
DoMenuItem omitted-argument parsing raises acceptance to 187/187; this is still
a syntax metric, never 187 executed procedures. The interpreter still has statement regex dispatch;
a full lexer/parser, classes, arrays and general VBA compatibility remain open.

REA 6.8.0 is an optional external analysis tool in the onboarding workspace. Its
CLI/ELF inspection was verified; no Ghidra/Windows/Access reference host is
available, and REA is not an application or package dependency. No fixtures were
submitted to external analysis services.

Validation for this increment: Release build; all 18 CTests passed (71.91 s),
GUI-disabled core test passed (3.00 s), fresh .deb payload/ELF Build ID checks
passed, and all four original fixture SHA-256 checks passed. Debian 12/Ubuntu
remote CI results are not verified from this environment.
