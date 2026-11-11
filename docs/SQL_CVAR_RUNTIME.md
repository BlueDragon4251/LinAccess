# Native SQL CVar identity

MS-VBAL6.1.2.3.1.13 defines CVar as returning its argument data value. The native
SQL adapter keeps proven Date/Currency/Boolean/Byte/Integer/Long/Single/Double
source kinds through the existing checked field decoder and CVar kernel. Result
metadata retains the proven underlying kind, including NULL rows; surrounding
native type inspection and Date expressions receive that kind on later binding.

Unproven SQLite arguments use sqlite3_result_value to preserve their actual
storage, including full signed64-bit integers, binary bytes, NULL and empty text.
No VBA LongLong, Jet Text/Memo capacity or other unknown source type is invented.
This raw storage identity does not add heterogeneous SQL Variant tagging or
Object/array/SCODE transport. SQL errors while evaluating an argument still fail;
CVar does not catch them. The existing separate Access Eval CVar guard is retained.

App built; existing dao_execution passed1/1 in5.82s. Actual typed values/DAO field
types, Null rows, nested type inspection/Date predicate, raw integers above2^53,
binary bytes, empty text and arity/prior-expression errors were checked together.
No new CTest ID or Windows/Access result oracle. The preexisting core CVar language
implementation was reused. F041/F067/F068/F118 remain open; fixed456 IDs/386
confirmed partial scopes/70 open full requirements remain unchanged.
