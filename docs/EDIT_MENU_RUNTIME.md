# Native text editing menu

The main window has a German Bearbeiten menu with Undo, Redo, Cut, Copy, Paste,
Delete and Select All. It remembers the focused owned editor when the menu opens
and routes commands through actual Qt key events. Native line/plain/rich text
editors perform their own edits and normal validation/event paths remain active.
Shortcut labels do not register competing QAction shortcuts.

Enabled state follows the live editor selection, read-only state, undo/redo
availability and clipboard text. Password line editors disable Copy/Cut. A
QPointer protects a target closed by a dispatched event. Unsupported focus
contexts disable the commands.

This increment covers text editors belonging to the main window, including
live datasheet cell editors and SQL/source text areas. Native datasheet selections
also enable Copy, Select All and bounded clean-model Paste; see DATASHEET_CLIPBOARD_RUNTIME.md. Typed clipboard MIME, record append/paste over prior drafts, exact Access clipboard formatting, a shared MDI menu for independent form/report windows,
Access command-bar customization and exact Access 2003 appearance remain open.
The existing main-window regression exercises the actual QAction commands,
clipboard text, editor undo/redo and read-only/password enabled states.
