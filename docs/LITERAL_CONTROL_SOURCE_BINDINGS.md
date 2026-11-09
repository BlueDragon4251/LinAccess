# Literal field names in native form bindings

Bracketed ControlSource identifiers decode doubled closing brackets before matching
the actual record-source columns. The same decoding is used for loading/saving,
bound-control field references and OldValue. Aliased controls therefore expose the
same current draft through their real field reference instead of falling back to
an unrelated stored-row read.

The native form generator brackets/escapes field names that begin with =, are
already enclosed in brackets, or contain outer whitespace. Their names remain
unchanged; the ControlSource carries the literal binding. These fields no longer
become formulas or lose their literal brackets/spaces. Ordinary field bindings
retain their prior representation. Imported source bytes are not rewritten.

The existing navigation fixture reads and writes actual SQLite columns containing
closing/literal brackets, outer spaces and an initial equals sign. It verifies
aliased Name/current/OldValue and independently checks the saved SQL values.
The final existing layout/GUI bundle passes 2/2 (12.39 / 14.16 seconds), and the
current DEB payload/Build-ID/MDBTools1.0.1/SQLite3.54.1 check passes. No new CTest,
full-suite/Windows oracle or original-workflow rerun is claimed.

This extends the bounded native F155/F056/F160 behavior. It does not complete
Access prefix-name rules, arbitrary nested source qualification, name ambiguity,
VBA source refactoring, query updateability or Windows event semantics. The fixed
catalog remains 456 / 386 confirmed partial scopes / 70 open full requirements.
