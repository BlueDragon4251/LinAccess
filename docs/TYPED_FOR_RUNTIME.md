# Native numeric For counters

Native simple-variable For/Next supports declared Byte, Integer, Long, Single and
Double counters. Bound and step expressions are evaluated once in order. Their
numeric values are retained independently of the counter's declared type; only
the start/Next assignment coerces to that type. Loop conditions compare the
assigned counter, including Single rounding, rather than the temporary sum.

This permits fractional limits with Integer/Long, signed steps with Byte and
zero steps with an explicit Exit For. Without an exit, the existing execution
budget still bounds a zero/nonadvancing loop. Failed Byte/Single/counter overflow
reports error6 before replacing the accepted counter value. Body writes to the
counter and zero-iteration start values remain observable. Existing exact native
Currency loops retain their separately documented bounded arithmetic path.

Implementation source: Microsoft [MS-VBAL], v20260818, section5.4.2.3 (For), with
numeric-type definitions in5.5.1.2.1. The public specification defines the original
uncoerced evaluated bounds/steps and treats zero as a nonnegative step. Declared
Date is outside its numeric counter types and retains preflight refusal3251;
unsupported counters do not execute a preceding SQL prefix.

- [Microsoft VBA language specification](https://learn.microsoft.com/en-us/openspecs/microsoft_general_purpose_programming_languages/ms-vbal/)
- Downloaded current PDF SHA256: `7cabb2af078a3ffaa41ad5833acedf99373d7fa9f62a2af8abc32d27d0bc0329`.

No PDF or private source is committed. This is specification-based native
execution, not a new Windows Access2003 observation. Full mixed Variant/Decimal/
Currency promotions, property/array counter targets, all nested Next-list error/timing cases and
all compiler/error-timing equivalence remain open under F045/F041/F122.
