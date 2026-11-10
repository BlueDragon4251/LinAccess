# Native SQL search, comparison and trimming

InStr, StrComp, Trim/LTrim/RTrim and their String-result $ variants use the
shared native VBA evaluator and the per-argument checked SQL type decoder.
They participate in the same proven projection/scalar-clause context as the
other native string adapters. InStr returns native Long (UTF-16 positions),
StrComp native Integer; Null results retain their field metadata. The callback
only accepts the documented native text/null/integer result kinds.

Two-/three-/four-argument InStr, optional comparison, typed Boolean operands,
ASCII-space trimming, String-result Null94 and range/arity errors follow the
shared evaluator. Binary comparisons use native UTF-16 order. Text/database
comparisons remain the existing bounded ASCII/provider contract: non-ASCII
comparison requiring locale support and LCID search modes refuse3251. This is
not complete Jet locale or Windows string behavior. Isolated surrogate/storage
and existing source resolver bounds remain.

App/DAO built; existing dao_execution passed1/1 in2.79s (2.80s total). Actual DAO
cases verify a UTF-16 search position after a surrogate pair, start/comparison
options, Integer/Long result kinds, nullable search/compare/trim, typed arguments,
String-result variants and guarded invalid/unsupported modes. The existing
CStr/substring/Replace/context cases run in that same bundle; no new CTest IDs
or catalog points.456/386/70 remains.

Latest checked package/Original/query inventory is source678a7fd82320afe5a255158a4e5d7f9e26b58944,
tree0fbe2a242c8c908120f6928ae6290d1bed7d1c30: the DEB passed contents, rights,
application/helper build IDs and SQLite checks. Existing original/inventory
bundle passed2/2 in163.22s (116.95/46.27). Seven protected hashes unchanged.
Inventory100/78/22 with categories73/5/3/9/10; no Access result oracle. Three
Original layout PDFs remain known refusals; full original sale/macro/keyboard/
report parity remains open. That DEB predates this search/trim extension.
