# Native VBA module state and lookup scopes

Procedure locals, native CodeBehind instance fields and standard-module project
fields now have distinct storage. Class state belongs to its AccessFormPreview
instance. Standard-module state belongs to the SQLite connection/driver context;
closing one form does not clear it. Opening another project or destroying the main
window explicitly resets the project context. Reopening a form creates fresh
instance state. Separate project connections do not share public globals.

The bounded module-preamble decoder supports Dim/Private/Public/Global scalar
Boolean, String and Variant fields (including a per-name As clause). Boolean and
String assignments are coerced and invalid Null/type assignments raise errors
without corrupting the prior value. Other module field types, constants, arrays,
user-defined classes/types, As New, declarations after procedures and complete
VBA compile/binding rules remain unsupported. This is not a complete class engine.

Module fields share real references with nested calls. Writes are visible to
recursive calls and following expressions; ByRef aliases retain their field type.
Procedure Dim declarations bind throughout the procedure, even inside skipped
branches or after an earlier read. Locals/parameters shadow module fields. Existing
controls have priority over public project globals in their own CodeBehind;
standard-module procedures cannot implicitly access a calling form's controls or
Me. Public CodeBehind fields are exposed through the native form object. Private
fields are accessible internally with Me, but not through another form reference;
bang lookup continues to mean Controls rather than class fields.

External lookup searches current imported standard modules, selecting the latest
case-insensitive module-name version. It excludes other CodeBehind instances and
private standard-module procedures. Ambiguous public procedures/globals fail
explicitly. A private procedure remains callable within its own standard module.
The full Access compiler's diagnostic codes, qualification, properties, inheritance,
module/static initialization, class destruction and project reset corner cases are
not established by this subset.

## Regression evidence

`tests/core_tests.cpp` executes persistent and recursive module mutations, local
shadowing, declaration scope, typed alias failure recovery and String coercion.
`tests/module_state_tests.cpp` loads every original VBA artifact with hash checks.
The original public declaration module
`b58d91036b71523ee9217433bcf809539d8747b7ea7cb9bf459161e86f1c9922`
provides its actual three public String variables to generic native VBA callers.
Their values persist across separate real form instances and remain isolated from
another project. Project reset clears those standard-module values. Generic native
helpers separately exercise shared private standard-module state, private call
visibility, updated module selection, ambiguity, Me restrictions, instance field
visibility and control/global shadowing.

The original three CodeBehind declaration blocks initialize four Boolean fields.
Those fields are unused/commented in the supplied originals; their initialization
check is not additional original-procedure execution coverage. The native caller
and helper procedures used for persistence tests are synthetic and do not modify
original code bytes. The original calendar helpers' globals can now be bound, but
Screen, the calendar ActiveX adapter, default-member chains and its complete event
workflow are not implemented by this stage.

Stage validation: all 18 CTests passed (59.27 seconds), the GUI-disabled core
build/test passed, a fresh Debian package passed executable Build ID/content/
permission checks, and all four original fixture SHA-256 checks remained unchanged.
This is local validation, not a remote distro CI or Microsoft Access reference run.

The subsequent [calendar stage](CALENDAR_RUNTIME.md) now executes the unchanged
standard helper against these globals through native Screen and date selection;
its evidence and remaining ActiveX/default-member boundaries are documented there.


The common source catalog now includes separately stored native standard modules.
The real editor and current proof/limits are in
[NATIVE_STANDARD_MODULES_RUNTIME.md](NATIVE_STANDARD_MODULES_RUNTIME.md). Imported
archive bytes remain unchanged; native/imported collisions refuse explicitly.
