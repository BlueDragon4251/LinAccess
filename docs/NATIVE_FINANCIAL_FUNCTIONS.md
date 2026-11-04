# Native FV and Pmt

FV(Rate,NPer,Pmt,[Pv],[Type]) and Pmt(Rate,NPer,Pv,[Fv],[Type]) now execute
native C++ floating-point calculations. The VBA binder supports the documented
named arguments, including Type (not the differently named .NET Due argument),
and omitted optional arguments default to zero. The SQLite adapter calls the
same expression-engine implementation for real Access SQL calls.

Supported scope: finite Double-coercible inputs, whole payment periods0..32767,
rate greater than-1, and Type0/end or1/beginning. Pmt with zero periods raises5.
Fractional/negative/larger periods, rate<=-1 and other timing values explicitly
raise3251 rather than invent compatibility outside the proven scope. Null raises94;
missing required/wrong arity raises450; nonfinite arguments/results raise6.
The result retains Double identity. There is no complete financial library claim:
PV/NPV/NPer/Rate/IRR/MIRR/depreciation and Windows error/edge-case oracle remain open.

The computation uses long-double intermediates and log1p/expm1 to avoid the
usual near-zero-rate loss from subtracting1 from a rounded power. Rate0 uses
the linear formula. Payment timing multiplies the annuity factor by1+rate.
The signs follow cash-flow direction, rather than absolute-value correction.

Existing core cases compare FV with independently accumulated monthly deposits,
and Pmt with an independently amortized balance including a future cash balance.
They check beginning/end timing, omitted/named arguments, zero/tiny rate and
zero-period rejection. Existing DAO cases execute both functions over actual
SQLite financial rows and independently check the balances. No extra CTest target.

Validation: targeted core/DAO/query3/3 (2.38 s), fresh GUI-free core1/1 (1.54 s),
Debian package/build ID/pinned SQLite and four original fixture hashes passed.
Last full baseline immediately before this change48/48 (82.86 s); no full-suite
rerun is represented as current. F068 remains open; fixed catalog456/384/72.

References: [VBA FV](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/fv-function),
[VBA Pmt](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/pmt-function).
No execution against Windows Access2003 was available.
