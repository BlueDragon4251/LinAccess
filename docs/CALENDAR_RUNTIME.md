# Native calendar workflow from the original Kassa bridge

The original unbound date-filter form now executes its two function-expression
DblClick bindings. Its unchanged standard-module function reads real
Screen.ActiveForm and Screen.ActiveControl/Parent/Name values, stores the original
three public String globals and opens the original calendar form. This is actual
VBA execution through native Qt controls, not a procedure-inventory result.

A narrowly selected Qt adapter implements the calendar for ProgIDs
`MSCAL.Calendar.7` and `MSCAL.Kalender-Steuerelement.7`. Other OCXs/COM objects are
not mapped to this widget. The original NextMonth/PreviousMonth handlers navigate
actual displayed pages, including the December/January boundary. SaveAsText has
no OnClick property on this original custom control: its known source-interface
Click event maps to the existing CodeBehind `control-name_Click` procedure. Syntax
inspection only finds that binding; a real Qt click then executes the procedure.

The original click handler's default-member chain first attempts its subform path.
The expected missing-control error 2465 is handled by its unchanged single-line
`If Err = 2465 Then Resume Next`; it then writes the selected date into the focused
original filter control and closes the actual popup through DoCmd.Close. Repeating
this against the second original date field verifies that Screen and the public
global target follow the new focus rather than a stale control. Closed calendar
references are invalidated.

## Regression evidence

`tests/calendar_tests.cpp` automates real double clicks, original buttons and a
real mouse click in the native calendar table. Both original date fields receive
an actual midnight date selected from the imported original MDB copy. A compiled
SQL query reads both live Forms values and returns a positive record count equal
to an independent C++ count of the actual source records. No chosen date, row,
message text, SQL source or credentials are logged. Tests fail on unexpected error
dialogs or a callback deadline rather than waiting for a manual click.

The unchanged source artifacts are hash checked:

- Date-filter form: `b7a13cd0fb0bf009966c418ac5eaecaba184f4117c40777a8d3149b9f9d5f907`.
- Its CodeBehind declaration file (no procedures): `fd6af9e988bfaa03d9751c765621e3e63693bc7d13642852afae0f98399ba281`.
- Calendar form: `c41791f3730a70a0b537b26054893d3561d7a12c764ff291d0d96051ba15fd9b`.
- Calendar CodeBehind: `19aace6611cb21892249ec065d5f305eb862b3fb2669d3e374968a6f2aa0a897`.
- Standard helper: `9d73028aa3f4a747ececda13983e784d39d22e2fa4671098e8e6d583829baeee`.
- Public globals: `b58d91036b71523ee9217433bcf809539d8747b7ea7cb9bf459161e86f1c9922`.

The corpus loader verifies the helper against the manifest, and the regression
preserves all original source bytes. No synthetic helper
replaces these calendar procedures. Synthetic tests separately cover default
member calls on real native subforms and DAO recordsets, including actual SQL
updates, field scalar arithmetic/comparison and Null values.

## Runtime fixes exercised by this path

- Single-line If/Then/Else, colon-separated branch statements, closest nested Else,
  outer GoTo labels and Resume Next execute real instructions with original lines.
  Keywords inside strings/bracketed identifiers are ignored by branch scanning.
- Text literals such as `"("`, `"not"`, `"-"` and `")"` cannot act as syntax tokens.
  The original helper's error-message expression exposed this lexer defect.
- Repeated default calls such as `Forms(name)(control).Form(field)` retain value
  versus reference context. Form controls and DAO recordsets have explicit default
  invocation; native Fields/Parameters/controls and Err have scalar defaults where
  implemented. Object identity with Is stays reference based. This is not complete
  generic VBA default-property/Let coercion or a full parser/AST.
- Function-expression control bindings execute original functions with typed
  values and source failures. Full macro bindings, named arguments and ByRef
  write-back through arbitrary event expressions remain open.

## Boundaries

This is a native calendar adapter, not Windows COM/ActiveX execution. It does not
restore the opaque OCX persistence blob's exact defaults, fonts, all properties,
locales or pixel rendering, and has no Microsoft Access behavior/rendering oracle.
Value is currently an ISO date/midnight String Variant, not a complete VBA Date
Variant or OLE serial. Time-bearing values and bound calendar data sources are
explicitly unsupported; their data is not silently rewritten. The fixture's
calendar is unbound. Full Screen semantics, other events, linked-subform calendar
workflows and general OCX interfaces remain open. OpenForm now returns with its
native target still open (OPEN_FORM_RUNTIME.md). The original calendar is not
marked Modal; tests drive its real asynchronous selection instead of relying on
the former synchronous-preview behavior. Complete Access window-mode semantics
remain open. Report pagination,
full form compatibility and all 187 procedure executions are not established here.


Stage validation: all 18 CTests passed (91.96 seconds), the separate GUI-disabled
core build/test passed, a fresh Debian package matched the current executable
Build ID and passed permissions/content/fixture-exclusion checks, and all four
original fixture SHA-256 checks remained unchanged. Parsed layout coverage is
47 forms / 879 controls / 86 visible command buttons. The 169/187 syntax acceptance
measurement remains separate from actual procedure execution.
