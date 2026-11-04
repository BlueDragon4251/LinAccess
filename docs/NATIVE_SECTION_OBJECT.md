# Bounded native Section object

Form/Report.Section requires one existing section index or retained imported
section name. Indices 0–4 map to detail, report/form header, report/form footer,
page header and page footer. Group sections are not fabricated. Name lookup is
case-insensitive and rejects ambiguity; absent source names remain unavailable
rather than receiving invented localized names.

The non-owning Section wrapper is weakly cached per document/section, preserving
identity while held and rejecting reads/writes after the native document closes.
Section/Access.Section are native TypeOf and typed-slot identities; no TypeName
string guessing is involved. acDetail/acHeader/acFooter/acPageHeader/acPageFooter
constants are available. Me.Section and a retained Report/Form reference use the
actual property call; Me!Section remains an explicit control lookup.

Supported reads are imported Name, source Height (Integer with overflow guard),
CanGrow metadata, BackColor (Long) and Parent, plus ForceNewPage (Byte) for report detail sections.
ForceNewPage writes validate Byte conversion and the 0–3 range before mutation.
They work during Report_Open and rebuild/invalidate an existing native page view
after later VBA calls. Mutation during rendering is refused. These are runtime
properties, not persistence into original MDB files.

BackColor writes validate the Long/OLE encoding before changing either model or
native canvas. They paint the full section content width and rebuild report pages;
prepared print bands use the same updated color. Explicit source/runtime OLE values
are returned unchanged. An unset color returns the actual native canvas base RGB,
not a fabricated original Access system-color token. Only the existing supported
RGB/system encodings are accepted. Full live system-palette re-evaluation and a
Windows color-profile oracle remain open.

Other section writes, Visible, group sections, direct named-member
shorthand, complete property collections and Format/Print/Page scheduling remain
open. Source Height is not claimed to be a fully formatted row's dynamic height.
The read-only Section property does not imply all Section members are read-only.

The existing report case runs actual typed VBA during Open, checks native class/
source metadata, writes detail ForceNewPage and observes four generated pages.
A second procedure resets it and immediately observes one real page. Native name/
index/Parent identities, invalid-write state preservation and closed-object guards
are checked without a new CTest target or catalog point.

References: [Report.Section](https://learn.microsoft.com/en-us/office/vba/api/access.report.section),
[Form.Section](https://learn.microsoft.com/en-us/office/vba/api/access.form.section),
[Section.ForceNewPage](https://learn.microsoft.com/en-us/office/vba/api/access.section.forcenewpage).
