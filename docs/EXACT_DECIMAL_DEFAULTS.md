# Exact Decimal literal defaults

Imported exact Decimal columns accept numeric decimal literal defaults that fit
their proven precision/scale without rounding. Native DDL stores the 24-byte
Decimal key as a bounded SQL blob literal. Source default spelling stays in
source metadata. No numeric SQL evaluation or Double conversion constructs the
stored value. Invalid scale/range or nonnumeric defaults refuse before creation.

DAO default metadata decodes only the bounded exact key spelling on a proven
Decimal field with its actual physical check and precision/scale. Pending AddNew,
new datasheet records and form defaults receive the encoded native value; the
existing literal-default query never executes arbitrary expressions. Field's
DefaultValue reports a canonical numeric expression. Existing NULL records stay
NULL and merely showing a default does not insert a record.

General expressions, Decimal field creation/property mutation, rounding to a
field scale, original-file writeback and Windows lexical equivalence remain open.
App/helper/import/DAO/GUI targets built; existing bundled checks3/3 passed
in25.62s (GUI18.97/DAO6.33/import0.32). The optional independent Jet generator
reopens and verifies actual p/s, values, defaults and NULL; the existing import
scope also reads those actual source defaults and checks pending DAO values.
Jackcess applies defaults to null insert arguments; the generator installs
default properties after writing explicit NULL reference rows, preserving them.
An omitted Decimal header include was corrected before checks. No new CTest ID. Fixed456/386/70 unchanged.
