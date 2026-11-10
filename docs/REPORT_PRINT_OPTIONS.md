# Native report page ranges, ordering and copies

The real report print dialog now offers page ranges bounded by the current
preview page count. The renderer accepts AllPages or an explicit valid PageRange,
clamps the last requested page to the actual report and refuses an empty range
before opening the output device. Selection/current-page modes remain disabled.
Original report page numbers and total pages are retained in selected output.

First/last page order, 1-100 copies and collation use the recorded source-page
pictures. Calculations and events are not rerun per copy. Native drivers that
support copies retain their driver copy count; otherwise LinAccess renders the
copies and temporarily disables driver duplication. PDF copies are rendered
explicitly. Multiple-copy output is bounded to10000 pages. Printer failures are
propagated through the existing document-print failure dialog.

PDF output chosen in the print dialog now uses the existing temporary-render /
QSaveFile commit path and transfers the actual chosen page layout, range, order,
copy count and collation. A failed range/render leaves a previous PDF untouched.
Symbolic-link destinations, NUL names and the current database path are refused.
The direct PDF export action still uses all source pages and the saved report
page setup. Unsupported grouped/formatting/event report contracts remain guarded.

The existing report/layout bundle verifies actual PDF page count and text order
through the real print dialog, both collation modes in reverse order, second-page
only output retaining2/2, restored current record and a failed-range prior-file
retention case. App and report/layout targets built. The existing access_2003_layouts bundle
passed1/1 in10.79s. Initial fixture failures were a missing reactivation before
the guarded print command and PDF text ligature normalization; only this existing
check was repeated after those fixture corrections. Rendering safeguards and
actual assertions remain active. No new CTest or
catalog identity and no printer hardware/Windows/original grouped-output claim.
