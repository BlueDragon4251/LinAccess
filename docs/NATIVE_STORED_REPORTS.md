# Owned native reports

Owned `_la_objects` report definitions now use the same native Report runtime as
imported report layouts. Double-click opens the report; **Objekte → Berichtsentwurf
bearbeiten** edits its title, SELECT source and VBA. New reports use the existing
create-report action. The legacy SQL-to-HTML preview and QTextDocument PDF path
have been removed for these definitions.

The adapter prepares source columns with `accessQueryColumns` without stepping
records, rejects action/make-table sources, empty/ambiguous columns and NUL SQL,
and generates an automatic fixed-band layout for 1–32 columns. Wider layouts can
exceed the selected printer page and are refused by the renderer rather than
clipped. This is not a drag/drop Access report designer or complete field layout,
grouping, CanGrow, subreport or print-event implementation.

The original multiline SQL is passed separately into the native layout model;
it is not flattened into SaveAsText string lines. The stored SQL and VBA remain
unchanged by opening. Existing extra definition keys, including archived original
VBA bytes, survive designer saves. The generated layout is an owned runtime
adapter, not an Access binary/export round-trip claim.

Report_Open/Load/NoData/Close procedures are automatically wired when declared.
Native controls, typed Report references, module state, real CurrentDb SQL,
readonly data binding and page/PDF output use the existing C++ engine. Unsupported
initialization results refuse full report rendering. Other event/property APIs
remain bounded as documented; code may still contain helper procedures invoked
from the wired events. The editor explicitly lists the supported automatic events.

Designer preview uses a nested native report window and returns to the retained
unsaved design. Saved reports opened from explorer or DoCmd have independent
project lifetime. DoCmd.OpenReport uses case-insensitive main-catalog lookup,
rejects ambiguity with imported reports and retains existing WhereCondition/view
limits. In particular a filter over an inline ranked TOP source remains 3251;
this adapter does not weaken that SQL guard.

Main-window object reads/writes now qualify `main._la_objects`, so a TEMP catalog
shadow cannot supply or overwrite report definitions. The existing main-window
case creates a report through the actual designer, checks exact multiline SQL/VBA,
executes TOP 2 over three SQL rows, checks code/typed controls/state/PDF, opens a
normal SELECT report with WhereCondition through actual VBA, closes its caller,
then edits/previews the report with four balanced Open/Close effects. No new
CTest target or work-catalog ID was added.
