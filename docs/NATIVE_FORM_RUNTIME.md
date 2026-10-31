# Native embedded forms and calculated values

This is a tested execution subset, not complete Access form/VBA compatibility.
Original Bridge bytes stay unchanged in the artifact catalog.

## Embedded forms

Unlinked SourceObject forms now instantiate as actual embedded Qt child forms,
with their original layout, CodeBehind and record source. The control's Form
property returns that child; Requery reloads its records and preserves source
errors. Direct-table child edits persist through the same typed form editor.
Children do not become top-level Forms collection members.

Closing a parent saves the document tree, runs unload validation, and closes
children only after every unload is accepted. Cancellation leaves all object
references live. This behavior has native persistence/lifetime tests; exact
Microsoft Access ordering across every nested event remains unverified.

LinkMasterFields/LinkChildFields are preserved but currently fail explicitly
with 3251. Missing/cyclic sources fail with 2455; cancelled opening uses 2501.
There is no replacement button pretending to be an embedded form. Unsupported
sources display an error. Linked filtering, subreports, dynamic SourceObject,
parent/recordset properties and complete subform event semantics remain open.

## Columns and calculated controls

Table/query combo row sources retain every typed column, including NULL.
Column(column[,row]) uses zero-based indices. BoundColumn uses Access's one-based
field index, or zero for ListIndex. The selected bound value is separate from
the displayed first column. Metadata ColumnCount/BoundColumn is preserved.

Value/field lists, user row-source functions, column widths, full multicolumn
popup/list rendering and automatic popup loading remain open. ListBox currently
shares the Qt combo adapter; it does not reproduce Access listbox presentation.

Calculated textboxes execute bracket field/control references, explicit subform
Form references, basic scalar expressions and simple Sum(field), Count(field)
and Count(*) over the loaded record source. English and bounded German expression
names/separators are supported. NULL aggregate behavior is tested. Incomplete
recordsets produce an error instead of an incorrect partial total. Domain
functions, general aggregate expressions, Currency/Decimal/Date subtype behavior,
formatting, arbitrary module functions and all recalculation timings remain open.

Variant/Error is represented separately from NULL and exceptions. CVErr/IsError,
CStr error output, ByRef updates and function returns have execution tests.
Arithmetic on an Error Variant in VBA raises 13; calculated control expressions
preserve error values. Unbound textboxes can hold them; bound fields and SQL
scalars reject them. Circular formulas produce visible errors. IIf still evaluates
both branches: an unused division by zero does not become a successful value.
This does not implement every Error Variant conversion or locale behavior.

Runtime Tag properties are backed by real Qt object state and preserve initial
SaveAsText values. Long quoted property strings are reconstructed from adjacent
SaveAsText fragments, including doubled quotes, without changing source bytes.

## Original Kassa evidence and remaining failures

The complex original Form_Load completes on an isolated 29-table import with
100 original queries, form definitions and VBA modules. A separate prepared
reader verifies all 15 dynamic command-button Tag assignments against actual
records. Six embedded child forms have available record sources. Its SQL changes
are rolled back before the independent report inventory runs.

A second, simpler original Form_Load verifies six Enabled/Visible assignments.
The complex form still has 7 error values among 34 calculated controls. Its
OnOpen now completes, including writes to the actual enclosing Qt menu bar.
OnCurrent still fails with
13; OnActivate is not dispatched. Thus successful Form_Load is not successful
complete application startup. DSum/CCur and their dependent expressions are
among the next unresolved calculations.

Ignored private event/calculation inventories contain module hashes, event or
control positions, type/status/error numbers and source line numbers. They do
not contain field values, raw SQL or VBA literals. No Microsoft Access result or
rendering oracle is available.

Native project menus share the existing toolbar actions (file operations,
object creation/opening and refresh). CommandBars.ActiveMenuBar returns that
actual QMenuBar; its popup Controls expose the actual QMenu actions. Enabled,
Visible, Caption, one-based lookup, object identity and deleted-widget guards
are tested. No full Microsoft Office menu hierarchy is synthesized.

Mixed numeric/String Variant comparisons use numeric-before-string ordering,
including empty and numeric-looking strings; Empty and Null retain their
separate comparison rules. This fixes the original Form_Open condition after
menu access. Declared scalar type retention and all coercion combinations
remain outside this bounded comparison regression.
