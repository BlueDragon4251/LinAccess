# Native calculated query column labels

Native SQL adapter names are now shared between token rewriting and projection
label preservation. An unaliased supported scalar call, or an ampersand
concatenation, receives a readable label derived from its source expression
before backend conversion. Internal typed Currency/Date/VBA helper expressions
use the same source-label rule. Original SQL values and proven output types
remain on their existing execution path.

Explicit AS aliases and existing implicit aliases remain unchanged, including
quoted aliases. Mixed wildcard/scalar projections and first-leg UNION column
labels are supported. Literals are inspected as tokens, not text patterns.
A string merely containing an adapter name does not trigger name rewriting.

These are native source-expression labels. Exact Windows Access Expr1000
generation, duplicate-name numbering and every compound/implicit-alias grammar
case remain open; this change does not claim full query-field naming parity.

Existing query-model cases execute actual financial calls, mixed-star
projections, UNION, implicit/plain/quoted aliases and concatenation, checking
values and user-visible headers. Initial DAO/query/original3/3 (8.32 s) passed;
final current full desktop48/48 (86.70 s) passed after the last alias adjustment.
The latter also covers all current financial and document/UI changes. Updated
Debian/build-ID/pinned SQLite3.54.1/four original hashes pass. GUI-free runtime is
unchanged from fresh payment-components core1/1 (1.79 s); this label change is in
the Qt data layer. Original-query coverage remains78/100 without Access oracle.
No new test target or catalog ID; fixed456/384/72, F068/F080/F148 remain open.
