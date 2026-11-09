# Native wildcard Find

Table/query/form Find dialogs have an explicit wildcard mode separate from
literal matching. It shares the native VBA Like parser and two-row UTF-16
matcher: * matches any sequence, ? one code unit, # an ASCII digit, and []/[!]
classes use the existing range rules. Whole/prefix/substring modes add the
corresponding pattern boundaries. Case mode, direction and wildcard state are
retained by FindNext. Literal UI text keeps its existing behavior.

DoCmd.FindRecord String needles containing * ? # or [ select wildcard matching.
Raw searches require supported text fields; formatted search can use supported
numeric/date displays and preserves its actual control/delegate presentation.
Password-formatted display is still masked. Typed numeric/date/Boolean equality
remains independent; wildcard replacement refuses3251.

Pattern syntax and absolute pattern length are checked before record-navigation
save guards. Matching is bounded to16384 pattern units and one million input x
pattern units per candidate, plus existing scan cell/time bounds. A malformed
class reports93. Per-candidate bounds are checked when scanning candidates;
existing navigation may save an accepted current record before matching.

The expression evaluator keeps its existing Like operation; search calls share
its parser/matcher rather than inventing another wildcard grammar. Native text
comparison uses the existing UI case comparison; non-ASCII text ranges and
surrogate comparison restrictions remain explicit3251. Full Jet/database/Windows
collation, all-field coercions, regional behavior and FindRecord error/event
identity remain open under F373/F071/F130. No new Windows Access oracle is claimed.
