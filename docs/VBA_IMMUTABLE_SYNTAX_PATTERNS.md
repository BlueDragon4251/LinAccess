# Reuse immutable native VBA syntax patterns

Sixty-two literal, input-independent std::regex patterns in the native VBA
compiler/declaration readers now initialize once, matching the existing static
pattern policy elsewhere. Pattern bytes and flags are mechanically unchanged.
No module source, symbol/type environment, argument, host, variable, object or
result value is added to shared storage. Match results remain local; module and
procedure state/context/caches keep their existing rules. Native grammar and
parser limitations do not change. Known primitive types also avoid scanning the
object vocabulary; Scripting lookups skip unrelated prefixes.

An unchanged20s GUI watchdog caught the real repeated-compilation bottleneck.
Optional LINACCESS_GUI_PROFILE diagnostics record phase names/times only; the
wizard took26ms, and no modal dialog was blocked. The final existing core/GUI/DAO
bundle passed3/3 in15.18s: core0.55, GUI12.29, DAO2.35. No checks or assertions were
removed and the watchdog was not relaxed. These are actual targeted run times,
not a general benchmark or Access compatibility percentage. No new CTest/catalog
IDs;456/386/70 unchanged, F040 remains the broader incomplete parser requirement.
