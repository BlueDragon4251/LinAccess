# Native project AutoExec

Opening a native LinAccess project resolves a saved native macro named AutoExec
case-insensitively after schema, saved-query views, project title and explorer
initialization. It runs once for that actual open through the same guarded
project starter as the database-window macro action. Reopening the same file is a
new project generation and runs it again. Merely saving/editing a macro does not
run startup code. Native module application caches are cleared on project change;
held objects still fail the generation check rather than silently rebinding.

Holding Shift while confirming/opening the project skips AutoExec. This initial
native implementation provides the enabled bypass behavior; AllowBypassKey=False,
full Database.Properties startup settings, a startup form, trust/security-center
and the complete Access startup sequence remain unsupported. It does not claim
to observe a source setting that has not been imported/implemented.

Only `_la_objects(kind='macro')` participates. Archived Access macro bytes are
not parsed or automatically run. A native macro may explicitly invoke supported
public standard-module Functions, using the existing source and language guards.
Original MDB/private VBA archives are not rewritten by startup.

Startup uses actual DAO/SQL and native documents. Project/connection/generation,
modal and execution guards apply. Errors are visible with their actual number
and macro/action context; the project remains available. Earlier completed SQL
effects persist if a later action fails, matching the existing macro sequencing
contract. A following Shift-open bypasses a failing startup macro. Explorer data
is refreshed after successful/partially failed project macro execution.

Contract reference: Microsoft Learn, AllowBypassKey property,
https://learn.microsoft.com/en-us/office/vba/access/concepts/miscellaneous/allowbypasskey-property .
This documents Shift/startup/AutoExec bypass and next-open effect; it is not a
fresh Windows Access2003 oracle or proof of the unimplemented property setter.

Validation: existing GUI1/1 passed (7.30s), matching app build. Real file-project
opens exercise SQL/RunCode/table opening once per open, actual Qt Shift input,
archive-only refusal/unchanged bytes, failure2517 with prior persisted effect,
Shift recovery and stale Application after actual project replacement3420.
