# Exact typed native UNION

Bounded parameterless UNION/UNION ALL prepares each SELECT arm independently.
Proven Currency REAL/tick outputs normalize to signed64 Currency ticks before
compound comparison, duplicate removal and ordering. Proven physical Decimal outputs keep
their sortable keys. Output names follow the first arm; typed DAO metadata and
readonly results retain the exact type. NULL, UNION ALL multiplicity and the
existing SQLite execution-order bounds remain. Simple final ORDER BY output
names/ordinals and ASC/DESC are supported. No value-sampling type inference.

Bounds: at most32 arms/256 columns, distinct output names per arm, compatible
proven exact numeric output types, no parameters/TOP/make-table/CDec compound source syntax or other compound
operators in an exact context. Unsupported exact shapes refuse3251 rather than
exposing raw ticks as ordinary numbers. Other unproven columns retain ordinary
unknown metadata. This is no full Jet compound/collation/parameter oracle.

App/parameter/DAO/model targets built. DAO/model passed5.98/0.47s; the parameter
scope first exposed the existing explicit CDec compound-source compiler boundary.
The Decimal case now uses actual imported typed columns; that original boundary
stays closed. Only the affected parameter scope repeated and passed0.29s.
Currency endpoints/carrier unification, DISTINCT/ALL/NULL/order, star projections,
readonly DAO type5, physical Decimal key order and unsupported shape refusals
are covered. No new CTest ID/Windows oracle. Fixed456/386/70 unchanged.
