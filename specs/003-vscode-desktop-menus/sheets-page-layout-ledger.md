# Sheets Page Layout parity ledger

The plugin mounts the desktop `SheetsRibbon` Page Layout band. Connected controls
use the shared validated page-command decoder and one canonical transaction.
This ledger records a checkpoint, not full desktop-menu completion.

| Group         | Controls                   | Current status                                                                                                         |
| ------------- | -------------------------- | ---------------------------------------------------------------------------------------------------------------------- |
| Themes        | Themes, Colors, Fonts      | Visibly disabled; canonical workbook-theme editing/save/projection remains open                                        |
| Page Setup    | Margins, Orientation, Size | Shared menus connected; original settings echoed where representable; margin presets recorded after edits              |
| Page Setup    | Print Area                 | Set from current selection / clear; saved defined name and reopen verified                                             |
| Page Setup    | Breaks                     | Insert, remove and reset; imported breaks retained and structural coordinates mapped; empty actions do not add history |
| Page Setup    | Print Titles               | First row, selected rows and clear; shared desktop row cap; saved defined name verified                                |
| Scale to Fit  | Width, Height, Scale       | Shared controls and desktop fit-axis defaults; native menu closes after selection                                      |
| Sheet Options | Gridlines: View / Print    | Separate screen and print flags; source metadata and canonical echo connected                                          |
| Sheet Options | Headings: View / Print     | Separate screen and print flags; canonical save/reopen connected                                                       |

Asynchronous preparation freezes the document root and owning sheet. Save waits
for preparation and commit. Overtaken source reads reject without changing the
newer document. Page-only edits preserve cell XML and existing chart caches,
including pending formula results, and do not read unrelated chart cell ranges.
Source protection metadata is still indexed before granting write access.

Broad installed-platform, locale, appearance and print/export acceptance remains
open. The other base tabs and DOCX menus also remain part of the active goal.
