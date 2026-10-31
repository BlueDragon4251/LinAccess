# Native print-preview command bar subset

Imported Report layout previews now have a real Qt toolbar with current-layout
preview, printing, PDF output and close actions. Application.CommandBars exposes
that toolbar as the standard Print Preview bar. This adapts a small Office API
surface to native Qt; it neither starts COM nor implements the Office object model.

Supported:

- Application.CommandBars(name/index), Count and enumeration. Only the actual
  Print Preview bar of this preview context is exposed. Office indices start at 1.
- CommandBar.Name/Enabled/Visible and Controls(name/index), Count/enumeration.
- Command control Caption/Enabled/Visible and Type=msoControlButton (1), backed
  by actual QAction state. Unsupported Office bars, members and IDs fail.
- Weakly cached wrappers and QPointer guards for deleted actions/dialogs.
- Caption reads/writes for native label/button/check controls and form captions.
  Textboxes and shapes do not acquire a fabricated Caption property.

The print preview uses QPrintPreviewDialog and QPrinter. PDF writes are generated
in a temporary directory and committed with QSaveFile, require a .pdf destination,
and do not replace MDB/ZIP fixtures. The displayed canvas is scaled onto one page.
The UI explicitly calls this the current layout. It does not run Access group
formatting, multi-record pagination, subreports, aggregates or page events.

## Original source execution evidence

Three hash-fixed original Kassa Report_Open procedures execute in full through
the same Qt VbaHost path as application handlers, using the original report layout
and original unmodified module. They maximize the window, enable/configure the
Print Preview bar, enumerate its actual native actions and hide actions according
to Caption. The regression asserts that only Close remains visible, resets the
actions and reruns each complete handler. Their unused Cancel parameter is zero.

The test invokes the runtime entry point explicitly; this package does not yet
bind Report_Open/Close automatically from the object navigator. That wiring and
report data/event semantics remain a separate open step. This is not proof of
complete original report execution or print equivalence with Microsoft Access.

Each original layout also produces a temporary one-page PDF, checked for PDF
format and page count, then deleted. Tests verify that a non-PDF sentinel file is
unchanged on rejected export. No source data, VBA literals, PDF output or raw
SQL are emitted in test summaries or packaged artifacts. Additional synthetic
checks cover one-based indices, action caption lookup/identity, actual toolbar
state and access after action deletion.
