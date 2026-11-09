# Native formatted search

Native table/query/form find dialogs offer an explicit formatted-display checkbox.
Raw typed search remains separate. Formatted searches compare literal String
needles against supported whole/prefix/substring presentations, with explicit
case sensitivity, directions and existing FindNext behavior. Their repeat state
includes the mode and the form control identity. Format errors are diagnosed,
never searched as success placeholders; NUL/malformed text is refused.

Tables validate checked field masks, then obtain the actual native delegate's
DisplayRole text. Queries use their materialized results and actual native text
delegate; expressions are not executed again to obtain a search candidate. A
native styled text delegate is required. Hidden rows/columns, Null, binary cells
and table checkbox presentations are omitted. Formatted-only query text search
does not invent a Jet datatype for an otherwise unproven expression column.

Direct-table forms format their bounded record snapshots using the actual bound
TextBox masks. Current-field VBA search retains the focused control's identity,
so differently formatted duplicate bindings use the selected mask. UI/all-field
search can examine each displayed bound textbox presentation. Hidden controls and
non-textbox/calculated/unbound controls are excluded. Password/NoEcho presentation
uses the native echo surface, rather than finding hidden plaintext as formatted
text. Existing raw search remains an independent mode.

DoCmd.FindRecord accepts SearchAsFormatted=True with literal String input and
supports OnlyCurrentField=acCurrent(-1) or acAll(0) in that mode. Raw all-field
coercion, wildcards and formatted typed non-String needles remain unsupported.
Formatted replacement refuses3251; masks are never reversed into stored values.
Native form/query/table activation and modal guards remain in force.

Formatted scans have a one-million-cell and five-second bound, checked per cell.
Existing form navigation can save an accepted current record before changing
records; validation failure preserves drafts and stops navigation. Table save/
unrelated-draft guards remain unchanged. Readonly query searching does not write
or reexecute its source. This is native literal display search, not full Jet/
Windows collation, wildcard, error-timing or FindRecord equivalence.

Explicit native wildcard search now extends the prior literal-only boundary;
see WILDCARD_FIND_RUNTIME.md. Full Jet/Windows behavior remains open.

Raw all-field FindRecord now uses matching native field kinds; see
ALL_FIELD_FIND_RUNTIME.md. Full cross-type/Windows traversal remains open.
