# Exact native Currency datasheet editing

Physical CURRENCY_TICKS columns use exact text display and editors. Text/numeric
input converts through the existing Currency codec; SQL storage remains signed64
scaled units. Empty numeric text means NULL. The typed clipboard marks Currency
units explicitly, so copying a tick field retains its Currency meaning rather
than becoming an ordinary integer dollar amount. Range undo/redo restores stored
ticks without scaling them a second time. Ordinary user integers retain ordinary
numeric conversion semantics.

Updates/deletes use the existing prepared raw-original identity path also used
for Decimal fields. This bypasses Qt's textual original-record cache. Ambiguous
physical matches refuse3251. No original MDB write-back, source MONEY import,
arbitrary field design or broad Access parity is claimed.

App and existing GUI target built; bundled GUI scope passed1/1 in18.46s.
The first technical case needed to reselect a cell after Save resets selection;
only the affected GUI scope was repeated. No new CTest ID/Windows oracle. Fixed456/386/70 unchanged.
