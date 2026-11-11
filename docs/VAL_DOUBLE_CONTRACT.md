# Native Val Double result contract

MS-VBAL6.1.2.3.1.21 declares Val(Value As String) As Double and states that the
result is a Double. Native radix parsing and supported decimal type suffixes
now keep their existing checked value conversion but return a Double value.
They no longer expose an incidental Integer/Long/Single/Currency subtype.
VarType/TypeName and subsequent Hex therefore see the function's declared
Double result. Existing overflow, Null, arity, unsupported radix suffix and
regional-prefix bounds remain unchanged; this is not full Val suffix/locale
or VBA7 LongLong support.

The bound SQL function preserves proven source kinds before native Val's string
conversion, including Boolean text and Single rounding/text representation.
Unproven sources keep the existing raw path. Result metadata is Double/dbDouble,
including empty results; it is not guessed from sample rows. The existing native
Single string conversion uses seven significant digits: Val(CSng(16777217))
therefore follows that policy and yields16777220, whereas the checked string
suffix Val("16777217!") yields16777216 as a Double. No new Windows observation
is claimed for that native formatting policy.

This extends the confirmed bounded F313 scope; it does not close the broad
F041/F067/F071 requirements. The fixed456 IDs/386 confirmed partial scopes/70 open
full requirements remain unchanged.

App built. Existing DAO passed5.29s and GUI passed16.17s in the shared contracts
bundle; final affected core-only check passed0.55s after removing an unsupported
Len storage-width expectation. No Len runtime change. Actual Double values/tags,
Hex width, SQL field types/values, empty results and existing error limits passed.
No new CTest ID or Windows oracle.
