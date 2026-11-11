# Exact Currency selector aggregates

Min/Max/DISTINCT Min/Max and native First/Last accept proven CURRENCY_TICKS
source types through the existing typed selector adapter. Decoded native
Currency comparisons preserve signed64 endpoints and last1e-4 digits. The
aggregate retains its tick encoding for final output and nested typed scalars.
Min/Max skip NULL; First/Last retain a selected NULL. Empty/grouped results keep
proven Currency result metadata. Existing execution-order bounds for First/Last
remain; no implicit Access ordering or new Windows oracle is claimed.

App and existing parameter/DAO/model targets built; bundled checks3/3 passed
in6.73s (0.20/6.07/0.45). Endpoints, DISTINCT, First/Last NULL, empty/grouped
results and nested VarType/CStr are covered. No new CTest ID. Fixed456/386/70 unchanged.
