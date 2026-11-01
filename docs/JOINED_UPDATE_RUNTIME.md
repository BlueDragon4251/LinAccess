# Bounded native joined updates

DAO Execute/QueryDef.Execute supports one INNER JOIN with either direct table
as target. LEFT [OUTER] JOIN supports the left target; RIGHT [OUTER] JOIN
supports the right target. Targets must be live physical tables with an
unshadowed physical row identity. Qualified assignments, aliases, parameters
and WHERE use the shared SQL compiler and origin resolver.

A materialized source freezes candidate values before writing, including
self-joins. Missing outer-source matches produce actual NULL values, and
WHERE source.id Is Null can select those rows. Multiple source matches are
accepted only if every assignment has identical SQLite type/value; differing
values raise 3251. Required-field errors raise 3314. Both failures roll back the
whole statement, including earlier candidate updates. RecordsAffected counts
actual target rows. These paths have native SQL assertions on independent
expected values, including preserved LEFT/RIGHT rows and rollback.

Nullable-side targets, FULL joins, multiple joins/targets, saved-query targets
and unavailable row identities remain explicitly unsupported. No missing
target row is silently inserted. This conservative boundary is not general
Jet updateability or an Access result comparison.

The original full-data corpus still executes 77/100 without an Access oracle.
Its INNER action has conflicting candidate values; the other action uses the
nullable LEFT side of a RIGHT JOIN and a saved-query target, so remains 3251.
The separate unchanged INNER action on synthetic unique input/actual schemas
is still one prepared positive probe, not a full-data execution increase.
