# Sheets Insert control ledger

Reference: the `activeTab === 'Insert'` branch of
`apps/sheets/src/renderer/SheetsRibbon.tsx`. The plugin now renders that exact
component branch in its Insert panel. This inventory records command families,
including generated options, rather than treating visible controls as completed
features. The remaining document/tab/contextual inventories are still required.

| Group         | Desktop control / variants                                    | Command                                                   | Plugin status                                                                                                                                                                                                                                                                                                                                                                                                                                     |
| ------------- | ------------------------------------------------------------- | --------------------------------------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| Tables        | PivotTable                                                    | `pivot-open`                                              | Open                                                                                                                                                                                                                                                                                                                                                                                                                                              |
| Tables        | Edit PivotTable                                               | `pivot-edit`                                              | Open                                                                                                                                                                                                                                                                                                                                                                                                                                              |
| Tables        | Table                                                         | `format-as-table`                                         | Open                                                                                                                                                                                                                                                                                                                                                                                                                                              |
| Illustrations | Pictures                                                      | `insert-picture`                                          | Shared desktop picker, validation and frame sizing connected; canonical insertion/format/history and repeated byte-preserving saves verified; clipboard/drop and broad acceptance remain open                                                                                                                                                                                                                                                     |
| Illustrations | Shapes, all `SHAPE_GALLERY_GROUPS` presets                    | `insert-shape:<prst>`                                     | Shared gallery and desktop drawing connected; canonical text/format/history and repeated mixed-object saves verified; broad preset/platform acceptance remains open                                                                                                                                                                                                                                                                               |
| Illustrations | Icons                                                         | `insert-icons`                                            | Open                                                                                                                                                                                                                                                                                                                                                                                                                                              |
| Illustrations | Screenshot                                                    | `insert-screenshot`                                       | Open                                                                                                                                                                                                                                                                                                                                                                                                                                              |
| Checkbox      | Checkbox at selection                                         | `insert-checkbox`                                         | Open                                                                                                                                                                                                                                                                                                                                                                                                                                              |
| Charts        | Recommended Charts                                            | `recommended-charts-open`                                 | Shared recommendation dialog and frozen asynchronous insertion connected; complete acceptance remains open                                                                                                                                                                                                                                                                                                                                        |
| Charts        | Column, bar, line, area, pie, scatter, radar, doughnut, combo | `insert-chart:<type>`                                     | All nine controls and desktop in-grid frames connected; original/session move, resize and deletion use transactions; session/original inline title/type/colors and bounded cross-sheet series ranges connected; shared context Format/labels/delete, format pane, Select Data series replacement/transpose and actual Chart Design ribbon, Alt Text/placement dialogs and layer menus connected; full contextual/property acceptance remains open |
| Charts        | PivotChart: column, bar, line, pie, doughnut, radar           | `insert-pivot-chart:<type>`                               | Open                                                                                                                                                                                                                                                                                                                                                                                                                                              |
| Sparklines    | Line, column, win/loss                                        | `sparkline:line`, `sparkline:column`, `sparkline:stacked` | Open                                                                                                                                                                                                                                                                                                                                                                                                                                              |
| Filters       | Slicer                                                        | `slicer-open`                                             | Open                                                                                                                                                                                                                                                                                                                                                                                                                                              |
| Filters       | Timeline                                                      | `timeline-open`                                           | Open                                                                                                                                                                                                                                                                                                                                                                                                                                              |
| Links         | Insert/Edit Link; URL/internal address; remove; cancel        | `link-open`, `link-set:<encoded target>`, `link-remove`   | Dialog, current-link read and atomic edits connected; see verification checkpoint                                                                                                                                                                                                                                                                                                                                                                 |
| Comments      | New Comment (desktop legacy-note popup)                       | `note-open`                                               | Open; preserve desktop threaded-note collision behavior                                                                                                                                                                                                                                                                                                                                                                                           |
| Text          | Text Box                                                      | `insert-textbox`                                          | Desktop cell defaults and shared text editing connected; canonical Save flushing connected; broad acceptance remains open                                                                                                                                                                                                                                                                                                                         |
| Text          | Header & Footer                                               | `header-footer-open`                                      | Shared dialog, frozen transaction and native save/history/reopen verified; full print acceptance remains open                                                                                                                                                                                                                                                                                                                                     |
| Symbols       | Equation                                                      | `insert-equation`                                         | Open                                                                                                                                                                                                                                                                                                                                                                                                                                              |
| Symbols       | Symbol: currency, math, arrows, Greek and miscellaneous grids | `insert-symbol`                                           | Shared palette and atomic append connected; calculation freshness and full acceptance remain open                                                                                                                                                                                                                                                                                                                                                 |

The source component provides the original icons, labels, ordering, dropdowns
and locale dictionaries. Pivot-sensitive hints read the current selection against
source/current pivot output ranges. Rendering a control or returning an explicit
unsupported-command error does not complete its row. Full insertion acceptance
also requires projection, authoring constraints, native host actions, save/reopen,
shared manual/AI history, contextual transitions and appropriate rejection paths.

Link editing uses the source/canonical target at the frozen selection's top-left
cell. Source hyperlink coordinates map through structural edits and canonical
removals/overrides; the host returns only in-range targets. The desktop target
normalizer is shared, and the actual desktop dialog retains synchronous desktop
behavior while the plugin awaits the authoritative transaction before closing.
Activation/navigation and grid decoration for links, every original target form,
all appearances/locales and the other rows remain acceptance work.

Symbol insertion reuses `SymbolDialog`, `appendSymbol`, and the same Univer
`covertCellValue` conversion invoked by desktop `FRange.setValue`. This includes
currency number formats and formula-looking output, rather than forcing every
result into plain text. Each click reads the frozen current anchor and writes
through the ordinary manual transaction. A subsequent click advances only to the
state/revision in that operation's committed receipt. Pending or unverified
formula results reject without a write; complete calculation-cache freshness,
including saved/reopened caches, remains shared acceptance work with Find.

Source-object contextual formatting is now connected separately from new-object
insertion: existing file shapes/pictures use the actual Shape Format/Picture
Format bands for paint (shapes), arrange, rotate/flip and numeric size/properties.
Their image bytes are read through the host-scoped media bridge. New shape/text-box
insertion is now connected through the shared gallery and drawing/text components.
New floating pictures also use the shared picker, file validation and cell sizing.
Graphical hyperlinks, clipboard/drop/embed and wider acceptance remain open; the
contextual checkpoint does not mark those Insert controls complete.

The shared Home clipboard group now has an authoritative Copy path as well as
Paste. Copy captures complete typed source cells (including cells outside the
initial grid window), delegates HTML/text serialization to the desktop service,
and retains a detached snapshot keyed by its copy ID. A subsequent same-editor
paste after source changes or sheet switching retains the copied values/formats,
shifts formulas and replaces/tile-translates merges through one canonical typed
transaction. Literal errors, dates, formula-looking strings and independent
borders are covered, including native save/reopen. This does not complete Cut,
Paste Special, copied metadata/rich text, filter semantics, cross-editor snapshot
portability, denied-copy reporting, clipboard fallback or installed OS acceptance.

Home Cut now uses authoritative native clipboard capture and canonical atomic
range relocation. Same-editor menu/key Cut supports overlap, cross-sheet formula
and absolute-reference movement, plain notes/links and contained merges, native
history, Save/reopen, stale-root rejection and Escape cancellation. This is partial
clipboard parity: affected native chart/defined-name metadata and unsupported
feature families still reject pending their transforms. Paste Special,
cross-editor cut, clipboard permission/cache fallback and full contextual/base
ribbon acceptance remain open.

Home Paste Special now uses the actual shared desktop dialog and all five quick
entries. Canonical selective fields support the shared arithmetic, skip blanks,
transpose, links, number formats, source themes/all variants and column widths
without flattening large styled ranges into scalar operations. Source freshness
follows desktop cached values/current format and own-layer live-source rules.
Comments/Validation retain the shared desktop disabled state. Formula-result
freshness, CF/rich metadata, filtered/merged target edge cases, source lifecycle,
cross-editor ownership and wider locale/platform acceptance remain open.
