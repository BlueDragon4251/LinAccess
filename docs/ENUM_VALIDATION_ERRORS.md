# Enum declaration validation results

The procedure validator now converts an Enum-name/declaration runtime exception
into its existing unsuccessful validation result with the original error number.
It does not let an invalid primitive-type Enum declaration escape through the
capability-report UI as an uncaught C++ exception. Allocation failure uses error7.
The source-only Enum-name scanner remains a checked throwing metadata API.

The capability report can therefore label an unavailable project type environment
and still classify each invalid source procedure. No VBA or business procedure is
executed by classification. A source line is not invented when the declaration
scanner has no mapped source location; complete compiler/source mapping remains
open under F040/F051. Fixed catalog456/386/70 remains unchanged.

Existing GUI passed16.17s and final core passed0.55s, including a primitive
type keyword Enum declaration and the real invalid-source capability report.
The shared DAO target also passed5.29s. No new target/CTest ID.
