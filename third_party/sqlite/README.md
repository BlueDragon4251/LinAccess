SQLite 3.54.1 — private LinAccess SQL backend

Official source: https://www.sqlite.org/2026/sqlite-amalgamation-3540100.zip
Published SHA3-256: 92281c86e2957f9b9305890f9ac04204caa7b677cc868dc326d942689df28e30
Release: https://www.sqlite.org/releaselog/3_54_1.html (2026-10-16)

SQLite is in the public domain: https://www.sqlite.org/copyright.html
The source archive is unmodified and contains SQLite's public-domain dedication.
It contains public upstream source only, never Access fixtures or user data.

The GUI build compiles the archive offline as libsqlite3.so.0 with column
metadata and thread safety enabled. The application loads it from its private
lib/linaccess directory; no system library is replaced. Qt's system QSQLITE
plugin uses the same stable SQLite C ABI. Before native connection access,
LinAccess checks the SQL connection's source ID against the compiled source ID
and the linked native library. A mismatch fails explicitly with error 3251.

Ubuntu 24.04's SQLite 3.45.1 crashes while preparing nested JOIN views in the
original query regression, without stepping records. A standalone native
schema-only reproducer confirms the library fault. Both original failing
regressions pass with 3.54.1. Pinning the backend gives the supported platforms
the same SQL engine; it does not establish complete Jet/ACE compatibility.

To update: obtain an official amalgamation, verify its published checksum,
update the archive/version/checksum in this directory, and rerun backend,
query, DAO, full fixture and installed-package regressions on each platform.
