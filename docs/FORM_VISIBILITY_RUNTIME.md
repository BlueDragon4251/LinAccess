# Native Form.Visible and hidden-document context

Explicit Me.Visible/Forms(...).Visible now reads the native form's own hidden
flag and can show/hide already activated top-level forms after initial Current.
This changes a real Qt window; it does not unload/close the form or emit finished.
Open-form collections retain hidden documents. Scalar Boolean input is evaluated
once and Null raises 94. Showing a hidden form preserves programmatic drafts and
does not implicitly save the source record.

Explicit native Hide deactivates/clears a hidden active document; queued/native
activation cannot make an explicitly hidden form active. Formerly activated
hidden callers cannot receive default record commands or unnamed Close when no
valid active native document exists (native 2046). Explicit named Close remains
available, including from Deactivate callbacks. Native SendKeys diagnoses missing
visible keyboard focus (2110) without delivering to a hidden editor.

Visibility changes during initialization, record/save/navigation/load/close or
reentrant phases diagnose 3251. Modal changes and hiding pending/dirty edits are
also explicitly unsupported until validation/save/focus semantics have a
reference. Showing dirty hidden forms is supported. Embedded/report visibility,
acHidden/open modes, complete Windows activation/MRU/MDI/error/event identities
and unqualified implicit property binding remain open.

The fixture-independent form_visibility_documents Qt/VBA/SQL regression covers
actual hide/show, explicit Me.Visible Boolean values, scalar-once/Null, hidden
collection membership, no early Unload/Close/finished, pending-hide rejection,
draft-preserving show, no implicit source save, all-hidden command/SendKeys/Close
rejection, hidden project-close Unload cancellation/retry, self-close from
Deactivate and modal rejection. Source table data remains unchanged.

Targeted visibility/mainwindow/original-layout/Kassa-DAO regressions pass 4/4
(48.48 s). GUI suite now has 33 CTests. Full validation is in PROGRESS.md. This
adds a bounded native property/lifecycle path, not full Access window parity or
additional complete original business workflows.

Reference:
https://learn.microsoft.com/en-us/office/vba/api/access.form.visible
