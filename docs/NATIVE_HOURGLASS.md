# Native busy cursor

`DoCmd.Hourglass HourglassOn:=True/False` now controls a visible Qt application
cursor override. `Screen.MousePointer` observes the supported default0/hourglass11
states and accepts those two assignments. vbDefault/vbHourglass constants are
available. Required argument450, Null94 and normal Boolean/Long conversion paths
apply; other pointer shapes and unmapped foreign overrides explicitly3251.

The cursor is an original, small native raster hourglass, not an assertion that
its pixels or animation match Windows Access2003. The copied custom QCursor is
retained as a Qt application property so equality can identify the owned override
rather than treating every standard WaitCursor as ours. Repeated enable does not
push another override. A pre-existing foreign override is refused without popping
it. Disabling/resetting only removes the matching owned top cursor. If another
Qt override temporarily obscures it, mapping/mutation is refused and cleanup does
not pop the unrelated override; complete foreign-stack interoperation remains open.

Native execution-depth cleanup resets the owned cursor after the outer native VBA
execution leaves, including errors; nested procedure/module and modal/event frames
retain it until that boundary. No form/document pointer is kept in cursor state.
Direct API assignments outside VBA can explicitly reset with MousePointer0.
This is the supported native GUI context, not GUI-free cursor emulation.

Existing core/mainwindow/layout3/3 (14.78 s), fresh headless1/1 (1.67 s), current
Debian/build-ID/SQLite3.54.1 and all four original hashes passed. Actual VBA checks
observe the cursor across a nested call; success/error exits clear the actual Qt
override. Direct assignment and preservation/refusal of an existing CrossCursor
are checked. No new CTest target or catalog ID. Last full baseline48/48 (83.85 s);
F068 remains open and catalog456/384/72 stays fixed.

Reference: [DoCmd.Hourglass](https://learn.microsoft.com/en-us/office/vba/api/access.docmd.hourglass).
