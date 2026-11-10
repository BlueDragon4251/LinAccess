# Native report grouping source metadata

The actual Access2003 artifacts contain42 active BreakLevel blocks across nine
reports, plus30 BreakHeader and29 BreakFooter sections. The earlier output bound
recognized only GroupLevel. Active BreakLevel ControlSource and unsupported break
sections now also prevent simplified report print/PDF output. GroupLevel spelling
is case-insensitive; an empty inactive BreakLevel source alone does not block.
The existing atomic PDF path preserves a prior target after refusal. Original
record-binding observations do not prove faithful grouped output.

The layout decoder also retains up to ten actual BreakLevel/GroupLevel records
in source order. Report.GroupLevel(Index) returns a cached native GroupLevel
object tied to the report's normal2450 lifetime guard. Available explicit source
properties return their actual documented primitive types: ControlSource String,
SortOrder/GroupHeader/GroupFooter Boolean, GroupOn Integer, GroupInterval Long and
KeepTogether Byte. Typed As Access.GroupLevel locals and function results and
named Index bind actual objects. Repeated access retains identity while held.

No property absent from the source is filled with a guessed default; it refuses
3251. Metadata mutation refuses383, bad arity450, Null index94 and an index absent
from the actual source5. This does not implement group creation/deletion, every
unused ten-slot/default-template rule, property mutation, group sorting/rendering,
GroupOn intervals, GroupLevel enumeration, Windows error precedence or full
report grouping/pagination/events. It exposes actual source state and keeps
unsupported output guarded. F050/F085/F087 remain broad open requirements.

App/core/layout targets built. Before adding metadata, the existing layout check
passed1/1 in13.67s for the actual BreakLevel output bug. The combined typed-object/
metadata change then passed the two existing core/layout checks2/2 in18.77s.
Actual VBA gets a typed GroupLevel through a named argument/function result,
checks identity and VarTypes, and reads actual stored values. C++ checks all
seven properties, absent-property/argument/write/index/closed guards. Native
PDF cases verify active/modern/lowercase/break-section refusal and unchanged
prior output. The existing original corpus still parses/binds all nine real
report artifacts without adding private VBA execution. No new CTest/catalog IDs;
456/386/70 unchanged. No Access grouping-render oracle has been supplied.

Latest checked DEB/Original is sourcea23c7009, Original1/1 in119.54s, seven protected
hashes unchanged. That package predates field capabilities/container keys/group
metadata. Latest actual query inventory remains4ebc3f6c100/78/22, not an Access
result oracle. No fresh full-suite/package/Original claim for this scope.

Public property contracts, not Windows execution-oracle evidence:
https://learn.microsoft.com/en-us/office/vba/api/access.report.grouplevel
https://learn.microsoft.com/en-us/office/vba/api/access.grouplevel.controlsource
https://learn.microsoft.com/en-us/office/vba/api/access.grouplevel.sortorder
https://learn.microsoft.com/en-us/office/vba/api/access.grouplevel.groupheader
https://learn.microsoft.com/en-us/office/vba/api/access.grouplevel.groupfooter
https://learn.microsoft.com/en-us/office/vba/api/access.grouplevel.groupon
https://learn.microsoft.com/en-us/office/vba/api/access.grouplevel.groupinterval
https://learn.microsoft.com/en-us/office/vba/api/access.grouplevel.keeptogether
