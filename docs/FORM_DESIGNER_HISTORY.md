# Native form designer Undo and Redo

The actual native form designer has Rückgängig/Wiederholen buttons and standard
Undo/Redo shortcuts. Ordinary text widgets keep their native text undo through
Qt ShortcutOverride. Designer history groups nearby input changes (250 ms), and
an immediate Undo flushes the pending group first. No timer delay is needed to
use the buttons.

History restores name/caption, source and actual field-list order, selected field
inclusion, labels, twip geometry, field event overrides, form events/timer, VBA
text and imported-original VBA bytes. Current field selection is restored by field
identity when possible. Geometry controls refresh and obsolete analysis output is
cleared. Source metadata is reread using the existing prepare-only source catalog;
Undo does not execute SELECT result expressions or save database objects. Runtime
form documents and imported original artifacts are not edited by this history.

Successful Save marks the current state as the saved baseline. Undo back to that
baseline clears the modified flag; new unsaved forms remain new. Undo across a
Save changes the editor draft only. Redo is retired by a new edit. History is local
to the open designer and is released on close; there is no new history file or
network export of VBA/source text.

The bounded history keeps up to 64 undo steps and 32 MiB of serialized steps;
a single state exceeding 8 MiB disables this history with a visible status while
preserving the actual editor contents. The saved baseline may additionally retain
one shared state. This is not full Access designer/toolbox, text-cursor/selection,
external schema mutation, report designer or cross-document undo parity.

The existing mainwindow target drives the real Save shortcut and Undo/Redo
buttons over source changes, field geometry/event overrides, caption and VBA.
It independently reads the saved SQLite object to verify that Undo/Redo has not
saved those drafts, and checks the saved-baseline dirty flag and restored original
field order/caption. No new CTest or catalog IDs are introduced; F080/F148 remain
open. Exact final validation is recorded in handoff/progress.
