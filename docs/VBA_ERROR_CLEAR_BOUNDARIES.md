# Automatic Err.Clear boundaries

MS-VBAL6.1.3.2.1.1 requires Err.Clear when executing any OnError statement or an
explicit ExitSub/ExitFunction/ExitProperty. The native interpreter now clears
Number, Source, Description, HelpFile and HelpContext at these actual boundaries.
Only active statements clear the object; a skipped conditional OnError leaves its
current state unchanged. Existing Resume resets and explicit Err.Clear remain.
The same actual ErrObject is cleared, including retained object references and
object-valued Function/Get results; this is not a replacement error snapshot.

Existing core/GUI targets passed2/2 on the first bundle23.07s (4.76/18.30), app
built. Shared technical source exercises all three OnError variants, all five
properties, inactive OnError, an actual GoTo handler with error704, and retained
ErrObject references across ExitSub/Function/Property. Existing rich error,
Resume/debugger and native property cases remain in those targets. No new CTest,
Original-event rerun, Windows oracle, full suite or package refresh is claimed by
this check. Full environment-wide singleton Err identity across independent host
frames, nested/global error histories and complete diagnostics remain F051.
Catalog456/386/70 is unchanged.
