# Native macro object storage and deletion

A new macro cannot replace an existing native/imported macro name, including a
case variant: it fails3010. Open the existing native object to edit it. Object
names become readonly after save, as in the native module editor; full rename/
reference refactoring remains open. Group labels remain editable with the
previously documented caller-reference limits.

Save/Delete operate only on physical `main._la_objects(kind='macro')` and preserve
an exact name/definition revision. Conditional writes inside a savepoint refuse a
changed revision3251. UI project identity/generation is checked after confirmation.
Metadata TEMP shadows and extra main/temp catalog triggers are refused. Definition
materialization is bounded to1MiB. Backend readonly/busy failures retain native
3027/3260 mapping; complete Jet locking and concurrency semantics are not claimed.

The editor and explorer provide confirmed deletion of an owned macro object with
all its submacros. No/cancel preserves it. The explorer can delete a malformed owned
native definition without attempting to execute it, allowing repair after a
Shift-bypassed startup failure. Imported source-only objects are not deletion
candidates; archive bytes are never rewritten by these helpers.

`DoCmd.DeleteObject` with explicit acMacro4 and a native stored object name uses
the same guarded delete path. No implicit current-object/group-member deletion is
introduced. Running macro frames protect the whole stored macro object, including
named submacro calls, from replacement/deletion until the invocation unwinds.
The frame state is shared by the real form and standalone project hosts. Existing
project/query/modal guards still apply; this is not a global lock against arbitrary
external/raw SQLite writes or complete Access object lifetime parity.

Validation: existing layout1/1 passed (11.74s); after removing an old duplicate
case-only UI guard, existing GUI1/1 passed (7.41s), matching app build. Actual
second-connection updates, changed revision during confirmation, No/Yes, native/
archive collisions, TEMP/trigger refusal, malformed-owned deletion and active
macro protection in both hosts are exercised. Subsequent DeleteObject succeeds;
archive bytes and original upload3/3/fixture4/4 hashes remain unchanged. No new
CTest target, full suite, package or Windows Access2003 oracle is claimed here.
