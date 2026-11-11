# Exact ACE2010 Numeric source import

The existing read-only metadata process accepts the checked17-byte Numeric
layout for MDBTools MDB_VER_ACCDB_2010 as well as Jet4. Precision/scale remain
validated, and raw/CSV/NULL/row/hash checks use the existing exact Decimal importer.
A separately generated/reopened Jackcess5.0.4 V2010 ACCDB with p28/s4 and p28/s28
provides positive/negative28-digit values, >2^53, 1e-28, zero and NULL.

The optional reference generator accepts V2000 (default MDB) or V2010 (ACCDB),
refuses replacement and verifies values after reopening. Neither database nor a
Java runtime/library is shipped with LinAccess. The existing optional
LINACCESS_JET_DECIMAL_REFERENCE import scope can exercise either file.

This does not enable other ACE versions, ACE field-display metadata, AutoNumber
source counters, complex/attachment fields, ACCDB forms/VBA, original-format
writes or complete ACCDB/Access parity. Those requirements remain open.
App/helper/import built. Existing import scope passed with both actual sources:
Jet4 0.46s, ACE2010 0.26s, including raw/CSV/value/DAO/constraint/source-hash checks.
The updated optional generator was compiled and its V2010 file reopened/verified. Fixed456/386/70 unchanged; F014/F017.
