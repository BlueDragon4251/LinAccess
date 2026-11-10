# Native non-Null logical result types

Native And/Or/Xor/Eqv/Imp now retain the supported effective primitive result
types from actual operand subtypes: Byte/Byte produces Byte, small integral
mixtures produce Integer, Boolean/Boolean remains Boolean, and wider supported
operands retain the existing Long conversion path. Byte Eqv/Imp use the actual
eight-bit result rather than a sign-extended Long value. No observed numeric
value is used to guess a source declaration or field type.

Non-Null logical results and Not carry the proven operand declaration category
through to subsequent arithmetic. Variant results can widen; fixed results
retain the existing checked overflow behavior. Object/default and unbound host
contexts retain their explicit binding limits. Access Eval's separate prior
logical path and all existing Null-result paths are unchanged.

MS-VBAL5.6.9.8 documents the type rules, but its general Or/Null table cannot be
used as a complete Access2003 oracle: the stored actual Access Eval reference
observes Null Or False as Null. The unmeasured numeric Imp/Null cases are not
changed or claimed resolved from that general table. Full logical Null, locale,
64-bit, interface and Windows equivalence remain open under F041/F118/F164.

The runtime's declaration-category lookup now reads existing type strings by
reference instead of copying them on every variable read/function return.
This preserves the same declaration selection; no benchmark gain is claimed.

Application/core/DAO/GUI targets built. The first existing bundled run passed
3/3 in27.29 seconds (core5.34/GUI18.72/DAO3.22). Checks cover the five binary
operators, Not, primitive subtype combinations, Variant/fixed follow-on overflow
and one actual native Byte-returning VBA function executed through SQL/DAO.
Its Eqv result is252 with DAO Type2/Size1 and an actual Byte Value. No new CTest
or catalog IDs. Fixed totals remain456/386/70; no full F041 parity claim.

The latest checked DEB predates this logical scope: source4cd0f848 at
/workspace/work/numeric-intrinsic-bindings/package/linaccess-0.7.0-Linux.deb.
Build-ID/SQLite3.54.1/MDBTools1.0.1/content/permissions/desktop/no-fixtures checks
passed. Its bundled GUI19.97/inventory46.77/original118.66 passed3/3 in185.40s;
query inventory100/78/22 and all seven original/fixture hashes unchanged.
The missing original macro-source boundary remains; this is not a positive
initialized original sale/report comparison or a complete Windows query oracle.
