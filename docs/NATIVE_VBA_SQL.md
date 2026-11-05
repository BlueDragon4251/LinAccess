# Native SQL calls into public VBA standard-module functions

Actual SQLite SELECT and UPDATE expressions can now call public functions from
archived VBA standard modules. The callback executes the same native C++ VBA
interpreter used by form handlers; it does not substitute predetermined values.
For example, a supported public `NativeScale(value As Long) As Long` body can be
used as `SELECT NativeScale(value) FROM inputs` or in an actual UPDATE assignment.
Current original-query coverage remains78/100; this extension is not claimed to
resolve the22 original dependency/context/column/join boundaries or supply an
Access result oracle.

## Binding and state

The checked main source archive supplies the latest entry per module name.
Private/Friend procedures and class modules are excluded from SQL registration.
Unqualified ASCII public function names up to128 characters are supported; ambiguous
names are refused rather than choosing a module. Public Sub procedures can be
resolved by native calls inside supported functions but are not registered as SQL
functions. Reserved native/sqlite names and existing SQLite functions are not
replaced. Module-qualified SQL names remain open. Static analysis only inventories
names; the native compiler/interpreter must actually accept and execute the body.

Callbacks use the existing per-connection standard-module state, preserving actual
module variables and scalar Static values across rows and query executions. A
fresh temporary host is created for each call; no Qt form host or persistent
QSqlDatabase is stored in callback registration. Weak driver/registry references,
SQLite-owned lifetime tokens, current native-handle equality and connection-thread
checks reject expired/reopened connections. Close/reopen resets the registered
module generation and shared state. Local arguments are copied SQL values; ByRef
changes to these temporaries do not become writes to the originating SQL field.

## Explicit first-stage scope

Arguments transport SQL NULL, finite numeric values and valid UTF-8 text including
embedded NULs. Integer values beyond exact Double transport are refused3251. The
native procedure binder applies supported declared types/defaults. Results support
Null, Boolean with True=-1, finite Number and valid String. Empty, Currency, Date,
BLOB, array/object/error results and complete Jet projection/Variant metadata need
further transport/type work and are explicitly refused3251. This is not complete
SQL/VBA type equivalence.

There is no GUI or DAO execution context inside the SQL host. Direct UI/database
host operations are refused. Existing non-null object and general array values in
the module/static/public-global state are also refused before body execution, so
an earlier form/DAO reference cannot silently enter this restricted call context.
Nothing slots and scalar state remain valid. Local native interpreter operations
retain their existing bounded behavior; these functions are not advertised as
pure. Registration has neither deterministic nor innocuous flags. Full reentrant
DAO, form context, COM/ActiveX, object/array state and class-instance calls remain
open. SQL binding and error precedence have no Windows Access2003 execution oracle.

## Source coherence and privacy

The registry checks source revisions and a per-connection SQL statement generation.
The native application owns the SQLite statement trace hook for this generation;
its callback ignores SQL text and statement contents and records no logs. Internal
source/state metadata reads do not advance the external statement generation.
Prepared queries therefore bind updated module bodies, including restoration after
ROLLBACK TO SAVEPOINT, without rescanning module bodies for every unchanged row.
Main/data/TEMP-schema revisions are checked; archive shadows are refused. Removed,
private or newly ambiguous definitions cannot keep executing a cached old body.
Other users of the same native SQLite trace hook, concurrent active source editing,
full transactional VBA-state semantics and large-project performance remain open.

Registration limits are256 latest modules,4096 public procedures,16 MiB encoded/
decoded source and60 SQL arguments. Failures emit generic native SQL error numbers,
not VBA source, arguments or private exception descriptions. Known native error
codes are propagated through the existing SQL error adapter, including3420;
complete custom Err.Raise mapping remains open. Sources stay in the private project
and runtime memory; packages contain neither modules nor fixture data.

## Verification

The existing DAO/QCoreApplication target executes actual SELECTs, a real UPDATE,
private-helper and cross-module calls, module counters, per-row Static progression,
Null/NUL text, Boolean and protected builtin abs behavior. It checks private/class
exclusion, GUI/Currency refusal, pre-existing DAO object-state refusal without body
effects, TEMP archive shadow detection, prepared source replacement, savepoint
rollback, ambiguity, removed definitions and fresh state after connection reopen.
No new CTest target or fixed-catalog item was added.

Production code full48/48 (87.35 s); additional actual UPDATE case in existing DAO
1/1 (0.74 s) passed afterward. Final targeted DAO/original-query2/2 (8.12 s), current
Debian/build-ID/SQLite3.54.1 and all four original fixture hashes pass. GUI-free core
runtime remains the unchanged fresh1/1 (1.66 s) from the string-mutation segment;
the SQL bridge itself is exercised by the Qt Core/SQL DAO target without Widgets.
F040/F068/F121 stay open; fixed catalog456/384/72 remains unchanged.
