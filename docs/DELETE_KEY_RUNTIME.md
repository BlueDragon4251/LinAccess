# Native Ctrl+Minus record deletion

Native forms now route Ctrl+Minus through the existing active-document Delete
Record command, including SendKeys "^-". KeyPreview can cancel before the
command; edit/deletion/source/new-record policy guards and pending-field
validation retain their existing behavior. No separate SQL-delete shortcut or
confirmation bypass is introduced.

The existing native Delete/BeforeDelConfirm/AfterDelConfirm pipeline uses a
savepoint and real Qt confirmation dialog with No as its default. Cancelled
callbacks and No restore the record/position; Yes releases the deletion and
updates the form. Foreign-key failure preserves the record. Full batch deletion,
query updates, Windows event/focus/layout/keypad identity and complete keyboard
parity remain open. This is native form behavior, not full datasheet parity.

The new fixture-independent delete_record_key CTest uses an actual Qt form,
physical keys, native VBA callbacks, real confirmation buttons and SQL records.
It verifies KeyPreview, AllowDeletions/AllowEdits, pending edit validation, Delete
cancellation, confirmation-event cancellation, user No/Yes, native SendKeys and
an attempted deletion rejected by a real foreign key. Default No, after-status
values, restored position, remaining record data and no false confirmation after
constraint failure are checked. The test explicitly activates the form before
subsequent keys in the offscreen environment after modal dialogs. Targeted
shortcut/deletion tests: 2/2 (0.90 s). GUI suite now has 25 CTests; full validation
is in PROGRESS.md. Original application business execution coverage is unchanged.

References:
- https://support.microsoft.com/en-us/office/keyboard-shortcuts-for-access-70a673e4-4f7b-4300-b8e5-3320fa6606e2
- https://learn.microsoft.com/en-us/office/vba/api/access.form.allowedits
The AllowEdits reference explicitly notes that Delete Record is unavailable for
existing records when AllowEdits is No; existing native policy guards remain.
