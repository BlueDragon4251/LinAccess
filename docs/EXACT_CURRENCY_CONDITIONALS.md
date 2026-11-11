# Currency encoding through conditional SQL expressions

Nz/IIf/Choose/Switch unify proven native Currency REAL and tick value branches as
Currency ticks when either value branch uses ticks. Branch selection does not
change the VBA Currency type or the signed64 value. Typed reparse/derived result
binding uses the same common type. Existing eager argument evaluation remains.

The shared scalar adapter derives Currency encoding from value-producing
operands. A tick-valued Choose index, IIf/Switch condition or Round precision no
longer changes a REAL Currency result into raw ticks under REAL metadata.
Known uniform REAL branches keep their existing encoding. This does not infer
uniform types for unproven literals or change mixed VBA type semantics.

App and existing parameter/DAO/model/core targets built; bundled checks4/4
passed in7.36s (0.53/0.32/5.93/0.57). Full-range/fallback/null branches, all four
conditional functions, nested CStr, SQL update and Currency-valued controls
with REAL results/round precision are covered. No new CTest ID. Fixed456/386/70.
