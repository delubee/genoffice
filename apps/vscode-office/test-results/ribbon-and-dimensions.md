# Ribbon and worksheet dimensions — 0.1.3

## Delivered controls

Seven accessible ribbon tabs replace the flat toolbar: Home, Insert, Page
Layout, Formulas, Data, Review and View. The README contains the available
controls and remaining desktop-parity limits. The UI supports keyboard tab
navigation and preserves the Excel / VS Code appearance preference.

- Home retains formatting, copy/fill and sheet/row/column controls, and adds
  numeric row heights (points) and column widths (pixels).
- Insert manages session column/bar/line/pie charts.
- Page Layout saves paper/orientation/margins/print area and scaling.
- Formulas inserts explicit formulas and common summaries without overwriting
  an occupied result cell; recalculation uses the existing formula engine.
- Data sorts the first selected column and applies bounded value filters.
- Review reads/writes/removes text notes and replaces text within a selection.
- View persists gridlines, headings, zoom and frozen panes.

Header resizing intercepts the native row/column commands before mutations and
submits validated shared operations. Hover over the middle of a separator to
use the native resize grip. Dimensions persist in canonical roots and follow
supported row/column shifts. Original sizes, hidden rows and custom heights load
from the source. Viewport pages adapt after row sizing, filtering and zoom.

Page, note and filter edits use the existing shared OOXML writers. These edits
participate in the same revision checks, durable undo/redo and save/reopen flow
as native agent operations. Saved note drafts clear after commit so later
AI/other-editor note updates remain visible. Value filters honor 1900/1904 date display values. Text sorting uses an
explicit zh-CN collation in both executions, so OS/UI language differences
do not select different row orders.
Original image notes, custom filters and unsupported reference transforms fail
closed. Pixel-to-column conversion currently uses the default 7px digit width.

The product skill is version 0.1.3. Tool confirmations now describe view zoom,
gridlines, headings and dimension effects. The desktop adapter rejects the new
VS Code-only view fields instead of silently ignoring them; its existing View
ribbon remains available. This is the first common-control migration, not full
desktop feature parity or a release-candidate acceptance claim.

## Verification

- Linux and Windows: three VS Code workspace typechecks passed.
- Linux and Windows unit suites: session 59, editor 10, extension 31 passed.
- Shared gateway: 252 tests passed.
- Desktop compatibility: 70 workbook-DSL/page-setup tests and its typecheck passed.
- Chromium/Univer: all 15 browser tests passed. Tests include actual keyboard
  input, actual column and row header drags, chart editing, all seven ribbon
  tabs, sorting/filter visibility, notes, replacement, formula results, view
  settings and appearance. A chart-form projection race found during regression
  was fixed by suppressing intermediate install notifications.
- Linux / VS Code 1.109.5: two actual extension-host tests passed. Original
  sizes load; resize undo/redo and exact save/reopen succeed. Page/view/freeze,
  notes and filter state survive XLSX save/reopen, including hidden rows without
  explicit heights. Presentation replies retain the finite-JSON boundary.
- Windows 11 / VS Code 1.109.5: four actual extension-host tests passed: new
  dimensions and ribbon settings, previous merged cells and appearance. Final
  packaging passed three tests including Chinese sorting through the actual
  Webview and host candidate, plus the two save/reopen scenarios.
- The focused native-tool confirmation test passed on Linux and Windows after
  adding explicit view-setting effect details.
- Changed-source ESLint, the theme-color gate and git whitespace checks passed.

Rust code and dependencies did not change. The previously validated static-CRT
Windows sidecar is reused. Real-model Chat, macOS, actual Remote SSH, print
preview and the remaining desktop dialogs are outside these verification claims.

Browser screenshot:
`packages/sheets-editor/test-results/browser/manual-intents-seven-ribbo-8b297--review-and-view-operations/excel-ribbon.png`.

## Final Windows package

`genoffice-workbook-0.1.3-win32-x64.vsix`: 6,786,484 bytes.
SHA-256: `52645eb71b092e417151b646b3ff05fb287e57f758375e37db302b6c095a7e66`.

The actual installed VSIX smoke passed in an independent ordinary VS Code
1.109.5 profile. It edited/saved/reopened value 731, read through a native tool,
and installed the updated product skill with Cargo absent from runtime PATH.
The downloaded package hash matches this exact installed package. Audit confirms
AMD64 PE, the unchanged static-CRT sidecar and no VC++ redistributable imports.
See `ribbon-package-audit.json` and
`ribbon-installed-win32-x64-proof.json` for the exact evidence.
