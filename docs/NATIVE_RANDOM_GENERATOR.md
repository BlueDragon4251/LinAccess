# Native project random generator

Rnd returns Single values in[0,1). Missing/positive Number advances the sequence;
zero reads the current value; a negative Single starts a reproducible sequence.
Randomize is a Sub with optional Number (including named binding). Missing/Empty
samples the actual local system timer; other values use checked Double coercion.
VBA-qualified calls use the same generator. User procedure/local resolution is
retained; Randomize cannot be used as a value expression.

The generator uses a24-bit state and the public MIT-licensed VBMath algorithm
reference. The initial state and implementation-defined details are a native
contract; this is not an observed Access2003 random-sequence oracle. Algorithm
attribution/license is in third_party/visualbasic-random and installed with the
application. No mock business values or test-specific product generator exist.

Actual mutable state is owned by the existing connection/project module context,
shared by forms, standard/module-editor hosts, native classes and native standard
module SQL callbacks. It is separate from syntax caches, module variables and Err.
Independent projects get independent state; project reset replaces it. Native
lifetime/thread/database guards apply when obtaining state. Scalar argument
coercion happens before state acquisition; failed primitive arguments leave state
unchanged. Native classes forward to their real service owner. Standalone hosts
have their own instance state, without process-wide mutable generator storage.

Direct calculated/Jet-expression Rnd and Randomize remain explicitly unsupported;
this does not implement every Access volatile-query/evaluation rule. A standard
VBA scalar function used from SQL executes with its actual project generator and
its existing non-deterministic SQL callback registration. No Windows sequence/locale/timing equivalence is claimed.

The existing core/GUI bundle checks repeatability, Single/range/current semantics,
negative seeds, numeric/timer Randomize, qualified/named calls, ordinary symbol
resolution, primitive error retention, two real forms plus standard-module and
SQL sharing, project isolation/reset and a closed database. App/core/GUI targets built. The initial existing core/GUI bundle passed the GUI
check in12.37s, including the actual cross-host SQL/session cases. Core exposed a
missing named-Sub fallback for Randomize; that binder path and qualified bare
Randomize were corrected. Only core was rebuilt/repeated and passed1/1 in0.49s.
No GUI assertion or20s watchdog was relaxed. No new CTest/catalog identities; broad
F041/F042/F043/F068/F071/F118/F164 requirements remain open.

Public semantic contracts: MS-VBAL6.1.2.10.1.6 and6.1.2.10.2.1.
https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/rnd-function
https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/randomize-statement
Public algorithm reference (not Access execution evidence):
https://github.com/dotnet/runtime/blob/main/src/libraries/Microsoft.VisualBasic.Core/src/Microsoft/VisualBasic/VBMath.vb
