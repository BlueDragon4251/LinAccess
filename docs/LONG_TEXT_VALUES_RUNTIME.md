# Native long TextBox value preservation

Native TextBox editors no longer use QLineEdit's default 32767 UTF-16-unit limit.
Their widget capacity uses the native int limit, preventing truncation while
loading longer bound Memo/text values. This matters because saving a record
serializes its bound controls: an unchanged but truncated Memo could otherwise
be written back when another field was edited.

A fixture-independent native regression first reproduced the failure in SQL
after changing a separate field. With the capacity correction, a generated text
value over 40000 UTF-16 units, containing Unicode, emoji and CRLF/LF, remains
complete through widget/Value access, another-field update, native VBA append,
Ctrl+apostrophe pending copy and explicit SQL save to a second record. The
shortcut capacity-guard test restores the actual editor capacity after its
intentional small-limit scenario. Targeted arithmetic/copy/delete/long-text
tests pass 4/4 (0.99 s). GUI suite now has 26 CTests; full validation is in
PROGRESS.md. Original fixture/business execution coverage is unchanged.

This fixes data preservation, not complete Memo UI parity. TextBoxes still use
the existing single-line adapter. Multi-line editing/rendering, full field-size
validation and Access-specific UI/format/large-memory limits remain open. The
widget capacity does not replace database field constraints or establish full
BSTR/type/encoding compatibility.
