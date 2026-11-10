# Proven Boolean DAO result metadata

Actual BOOL/BOOLEAN result declarations now map to DAO dbBoolean (Type1) in the
recordset field metadata path, with the existing one-byte Size. Native Boolean
Variant values remain Boolean, including False; this does not infer Boolean
from ordinary Integer or guess arbitrary expression result types.

The existing native SQL module case checks QueryDef.Fields metadata without
executing the module body, actual recordset Fields.Type/Size/Value for true and
false native Boolean function results, and its existing Currency/Date/Long and
stale-metadata guards. The existing dao_execution target is reused; no new
CTest or catalog ID. Broader metadata/type/oracle requirements remain F121/F122.
