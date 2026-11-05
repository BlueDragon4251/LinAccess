# Native LSet, RSet and Mid string assignment

LSet/RSet now compile native string-alignment instructions for existing String
variables or Variants currently containing a String. Target width is its current
UTF-16 code-unit length. LSet pads on the right, RSet on the left; longer inputs
use a left prefix up to that width. An empty String is not expanded. Normal
native String RHS coercion applies. The common assignment path preserves typed
slots, module aliases and actual ByRef updates. Constants are refused before
execution effects. UDT memory copying, fixed-length declarations, object/control
lvalues and non-String target variants remain explicitly outside this scope.

Writing Mid(stringVariable,start[,length])=replacement now shares indexed syntax
but dispatches to native string mutation only when no actual variable named Mid
is bound. Existing array/object default-member assignment paths retain priority.
The number of UTF-16 units replaced is the minimum of requested length, source
length and remaining target length. Omitted length uses the source length; zero
length does not change the target. Prefix/replacement/tail use the existing native
UTF-16-aware string functions, then the common typed/ByRef assignment path.

Supported Mid start positions are1..current nonempty target length. Start<1 and
negative length raise5, Null94, invalid argument count450; empty/out-of-range
starts and non-simple/non-String targets explicitly3251 pending reference evidence.
A required truncation or boundary inside a surrogate pair is refused3251 before
mutation, retaining the original target. This native UTF-8-backed implementation
does not pretend to preserve arbitrary unpaired UTF-16 surrogates. MidB, fixed
Strings, general lvalues, complete coercion/error precedence and Windows Access2003
execution equivalence remain open. Source-only cached plans contain no host state.

Existing actual VBA core flows check alignment/padding, Variant String truncation,
whole Unicode pairs, refusal without mutation, zero-width/zero-length behavior,
Mid's documented dog/fox/cow example, actual ByRef changes and a two-dimensional
String array named Mid. Desktop core1/1 (1.82 s), fresh GUI-free1/1 (1.66 s),
current Debian/build-ID/SQLite3.54.1 and all four original fixture hashes pass.
Last full baseline48/48 (85.02 s), not rerun after this bounded language extension.
No new CTest target or catalog ID. F040/F068 remain open;456/384/72 stays fixed.

References: [LSet](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/lset-statement),
[RSet](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/rset-statement),
[Mid statement](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/mid-statement).
