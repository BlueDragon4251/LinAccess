# Typed native SQL case conversion and reverse search

LCase/UCase, their $ variants, StrReverse and InStrRev now use the common
per-argument native SQL decoder/evaluator, preserving proven input kinds in
projections and scalar clauses. This fixes Boolean text conversion and bounded
Currency digits and gives InStrRev its proven native Long result metadata.
The prior basic SQL callbacks remain registered under the same internal names;
String-result variants and typed callbacks are added.

Existing bounds are retained: case conversion is ASCII only and rejects
non-ASCII3251; comparison requiring locale support remains guarded. Reverse
operations use the shared UTF-16 contract, including existing isolated-surrogate
limits. StrReverse(Null) raises94, while nullable case/search calls propagate
Null and $ case calls raise94. No new Windows case/collation claim.

App/DAO built; existing dao_execution passed1/1 in2.96s. Actual DAO paths cover
typed Boolean case/reversal/search, exact four-place Currency text near2^39,
UTF-16 reverse-search position and Long kind, optional ASCII comparison, Null
distinctions, String-result variants and unsupported/start errors. One fixture
initially expected a Null StrReverse result; it now checks the existing runtime
94 contract instead. No runtime guard was weakened, no new CTest/catalog ID;
456/386/70. Latest checked DEB remains source678a7fd8, before search/trim,
repeated group headers/Section access and these adapters.
