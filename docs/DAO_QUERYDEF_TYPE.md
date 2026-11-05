# Native DAO QueryDef.Type

Anonymous/native owned QueryDefs now expose read-only Type as an Integer. The
Access compiler records the source kind before backend rewriting: SELECT0,
DELETE32, UPDATE48, INSERT/append64, SELECT INTO/make-table80, CREATE/ALTER/DROP96,
and outer UNION128. Top-level token inspection excludes strings/quoted identifiers
and nested set operations. Declared PARAMETERS are consumed first. Joined UPDATE
explicitly keeps48 even though its backend uses a materialized SELECT/CTE; TOP
and make-table rewriting likewise retain the source kind. Current plan copies and
owned SQL refreshes carry this metadata with the actual compiled query.

This is native statement-kind metadata, not a complete syntax-validation engine
or original Jet QueryDef metadata reconstruction. Non-owned named/imported sources
are explicitly3251 because missing Connect/passthrough/crosstab/procedure metadata
must not be guessed from SELECT text. Empty/unclassified native sources3251;
Type assignment383. Existing QueryDef workspace/connection/thread/closed and
named-generation guards still apply. Native short DAO QueryDefTypeEnum constants
are defined, including unsupported kinds; exposing constants does not implement
passthrough, crosstab, ODBCDirect, compound or procedure execution.

The existing DAO case verifies PARAMETERS/literal/nested-UNION/outer-UNION/TOP and
joined-UPDATE classification, then actually creates, inserts, updates, selects,
makes a table, deletes, alters and drops temporary native tables through one
changing QueryDef. Result9 is independently read through a real Recordset. An
owned VBA call checks dbQSelect/Integer identity; imported metadata refusal and
read-only assignment are checked. No extra CTest target/catalog point.

Targeted core/DAO/query3/3 (2.25 s), fresh GUI-free1/1 (2.75 s), current Debian
package/build-ID/SQLite3.54.1 and four original fixture hashes passed. Last full
baseline48/48 (85.02 s), not rerun after this bounded extension. F068/F121 remain
open; fixed456/384/72 unchanged. No Windows Access2003 execution oracle.

References: [QueryDef.Type](https://learn.microsoft.com/en-us/office/client-developer/access/desktop-database-reference/querydef-type-property-dao),
[QueryDefTypeEnum](https://learn.microsoft.com/en-us/office/client-developer/access/desktop-database-reference/querydeftypeenum-enumeration-dao).
