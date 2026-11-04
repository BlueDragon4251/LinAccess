# Bounded native VBA procedure translation cache (2026-11-04)

Successful native procedure translations can be reused within their execution
thread. The key includes the exact full module source, case-normalized procedure
name, label-resolution mode and inherited With parsing context. Hashes only
preselect candidates; exact source bytes are compared. Edited source therefore
cannot reuse an old plan under the same procedure name.

Entries retain source-only statement/parameter metadata, including immutable
scalar constants/defaults and folded fixed dimensions. Execution receives its
own statement vectors. Module/global variables, argument aliases, mutable
objects, DAO connections/cursors, form/report widgets, document focus and host
contexts remain outside the cache and bind again for each invocation. Plans with
live object/array metadata or execution return values are not admitted.
Direct-call resolution and runtime stop/lifetime checks continue for every call.

The per-thread LRU holds at most 128 entries and a 16 MiB accounted data budget
(source, vectors, strings and scalar payload estimates; allocator overhead is
additional). Oversized plans run uncached. Admission allocation failure does not
reject an already translated procedure. Failed/missing translations are not
cached. There is no persistent code cache, source dump or result memoization.
This caches the existing bounded native representation; complete VBA grammar,
AST/compiler semantics and general optimized execution remain open.

Existing core integration checks validation-before-execution metadata reuse,
case-normalized calls, unchanged per-host module state, separate host state and
changed-source/default behavior. No new CTest targets are introduced.

An external synthetic native runtime probe executed the same private counter
function 200 times against the immediately preceding runtime archive and the
cached runtime. Both produced final value 200. On this environment, sequential
runs measured 784.665 ms and 37.9019 ms. This is one bounded computation example,
not an app-wide benchmark or an Access 2003 performance oracle. Probe files stay
outside the repository and contain no original source/data.

Checkpoint: full desktop 48/48 (81.60 s), freshly rebuilt GUI-free core 1/1
(1.18 s), Debian payload/permissions/Build-ID/pinned SQLite and all four original
fixture SHA-256 checks pass.
