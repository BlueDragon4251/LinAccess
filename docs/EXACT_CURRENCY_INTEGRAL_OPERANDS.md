# Exact Currency with proven integral SQL operands

Known Currency tick operands now combine with proven Byte/Integer/Long/Boolean
fields or typed scalar results in +,-,* and comparisons. Result arithmetic uses
the existing native VBA Currency coercion/operation kernel, retaining scaled
signed64 storage. Typed IN/BETWEEN accept the same known integral operands.
Unknown plain literals, unproven integer SQLite columns, floating mixtures and
implicit dialect promotion are unchanged; explicit conversions remain necessary.

SQL UPDATE uses the same input-aware target adapter. Overflow6 aborts the SQL
statement without partial earlier-row writes. NULL and known Boolean -1 retain
the existing kernel semantics. This adds no Windows Jet numeric-promotion oracle.

App and existing parameter/DAO/model targets built; bundled checks3/3 passed
in7.05s (0.22/6.32/0.51). All integral source kinds, Boolean -1, NULL,
IN/BETWEEN, multiplication/update and atomic overflow are covered; the existing
unproven literal refusal is retained. No new CTest ID. Fixed456/386/70 unchanged.
