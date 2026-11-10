# Native repeated report group headers and Section access

The actual group designer persists repeat_header as a Boolean. It requires a
group header; existing definitions default false. Native source emits explicit
RepeatSection and real generated section names. The parser retains the setting
on the distinct, proven group section. Other RepeatSection contexts, columns,
subreports, report events and advanced grouping remain guarded.

The planner tracks active group scopes, inserts repeating headers in outer-to-
inner order on continuation pages, accounts for their height, and removes
closed scopes. Repeated headers use the group's representative key row. Only
checked group-key TextBoxes (no running/growing or arbitrary row-dependent
expressions), labels and existing supported static graphics are within this
repeat contract. Remaining renderer control/format/event guards still apply.
The planner checks available body space and does not keep opening empty pages
when a detail slice cannot fit beneath repeated headers. Atomic PDF target
replacement and record/layout restoration remain.

Section(index) now supports proven header/footer pairs5/6,7/8 and consecutive
pairs through the ten-level native bound, as documented by Report.Section.
Only the four published AcSection group constants are added; no extra enum
names are invented. Name/index access shares section identity. RepeatSection
is Boolean read/write on proven loaded report group headers. Enabling an
unsupported repeat context rejects3251 before mutation; rendering-time mutation
is refused. Existing guarded refresh/lifetime behavior is reused; Null94 and
closed-object2450 remain. No metadata or original MDB is persisted by a live
property assignment.

App/GUI/layout built. Existing GUI passed12.75s with the20s watchdog unchanged.
After correcting the negative fixture to change only its header (not its group
key), the targeted access_2003_layouts run passed1/1 in10.45s (10.46 total).
Actual three-page PDFs verify enabled/disabled repetition, each record once,
restoration and retained prior PDF on refused arbitrary headers. Actual VBA
uses the new constants and Access.Section type to toggle the property; native
index/name/parent identities across two group levels, invalid/Null settings
and closed lifetime are checked. This is not a fresh combined2/2 run or a
Windows pagination/event/rendering oracle. No new CTest/catalog identity;
456/386/70.

Latest checked DEB/Original/inventory is source678a7fd8 (116.95/46.27s; seven
hashes unchanged), before search/trim and this feature.

Public contracts, not Access execution observations:
https://learn.microsoft.com/en-us/office/vba/api/access.section.repeatsection
https://learn.microsoft.com/en-us/office/vba/api/access.report.section
https://learn.microsoft.com/en-us/office/vba/api/access.acsection
