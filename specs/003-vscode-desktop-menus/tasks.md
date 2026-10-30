# Work remaining for full desktop menu parity

- [x] Extract the desktop Sheets command band without maintaining duplicate UI.
- [x] Expose browser-only office chrome and isolate Paste Special's UI model.
- [x] Add controlled locale scopes for desktop component reuse in embedded hosts.
- [x] Remove the AI UI barrel from the Docs ribbon and its reachable UI helpers.
- [x] Verify the shared Sheets component's embedded interactions and browser boundary.
- [x] Integrate the desktop Font dialog in DOCX with validated font effects and advanced character settings, persistence and shared history.
- [x] Integrate the desktop Paragraph dialog in DOCX with validated spacing, indents, outline and pagination fields, persistence and shared history.
- [x] Reuse desktop Docs styles and controlled locale for the embedded dialogs; preserve both appearance modes.
- [x] Connect shared Open, Save, Save As, document chat and selection chat entry points to explicit native host actions for DOCX/XLSX.
- [x] Add the shared file/chat entry labels in all 21 supported languages.
- [x] Verify file/chat routing in the built DOCX browser UI and real native XLSX save with inactive-target rejection.
- [ ] Build a complete control/variant ledger for both document types, including contextual menus, dialogs and panes.
- [x] Generate a reproducible source inventory for both desktop ribbons and directly named context menus, recording source hashes, action locations, constraints and dynamic variants without inferring plugin completion.
- [x] Review the full desktop Data command group against the current plugin panel and record missing workflows in sheets-data-ledger.md.
- [ ] Map and implement every Sheets ribbon command through authoritative session operations.
- [ ] Reuse Sheets dialogs and supporting panes; preserve desktop parameters and state echoes.
- [ ] Extend Sheets operations/projection/serialization for uncovered desktop capabilities.
- [ ] Integrate the actual shared Sheets ribbon and desktop chrome styles in the plugin.
- [ ] Extract Docs privileged host ports throughout ribbon, tabs, dialogs and editor helpers.
- [ ] Extend DOCX schema policy, conversion, session validation and serialization for rich desktop features.
  - [x] Policy rewritten around raw/anchor attributes: raw XML, source anchors, revisions, fields and comment ids stay immutable (a new or split owner drops them); every other schema attribute and mark may be authored within shape bounds (strict JSON, image data URLs, no script links). Only redacted content and host objects stay protected placeholders. Mid-paragraph splits no longer fail on the copied identity.
- [ ] Integrate the actual Docs ribbon, dialogs and supporting panes in the plugin.
  - [x] The desktop `Ribbon` is mounted on the DOCX runtime (DocsRibbon.tsx): editor commands, styles gallery, Insert table picker, Table Layout, font name/size and file/chat quick actions verified in the built browser page; view toggles are local; host capabilities without a VS Code port report explicitly (page setup, headers/footers, notes, comments, revisions, protection, mail merge, find, export).
  - [x] Desktop dialog shortcuts Ctrl+D (Font) and Ctrl+Alt+M (Paragraph) at the window, as the desktop shell.
  - [x] Desktop editor context menu (Font…/Paragraph… and editor entries), Find/Replace panel with Ctrl+F/H, and Insert → Pictures through a typed Docs host port (docs-host.ts; desktop keeps its preload bridge, the view reads the file in the webview). New pictures are accepted by the policy only in the desktop's exact shape and save as new media parts.
  - [x] Ribbon bullets / numbering / multilevel galleries: the desktop numbering actions allocate list definitions in the view; they travel with the steps that use them, are validated by the session (no XML-bearing marker text, fresh identities, known abstract lists) and join the overlay in the same history entry; presented, undone and saved to numbering.xml.
  - [x] Layout / Design document settings: final-section page setup (margins, orientation, size, columns via the shared tabs), page color, mirror margins and text watermark through a typed `docx.settings` request — validated (declared sectPr keys only, identifier-safe words, bounded numbers), one coalesced history entry per gesture without document steps, presented to the page geometry, saved to sectPr / w:background / settings.xml / the header watermark, reopened.
  - [x] Links: the desktop link dialog (Ctrl+K, context menu) and the Insert tab dialog author link marks; saved as external hyperlink relationships; script/data/file schemes refused.
  - [x] Default header/footer: Insert → Header/Footer galleries and Edit Header/Footer mount the desktop strip editor on the page; edits are settings entries validated as plain JSON (raw run XML only as the file's own values), saved as regenerated parts with PAGE fields, reopened.
  - [x] Comments: Review → New Comment / Delete / Previous / Next / Show Comments and the context menu run the desktop review actions; the desktop Comments panel composes, replies, edits, resolves and deletes. The comment list travels with the anchor steps (one history entry) or alone as a settings entry; anchors must name listed comments; saved to comments.xml with resolved state, reopened.
  - [x] Footnotes / endnotes: References → Insert Footnote/Endnote open the desktop note prompt; the reference and the note list commit together (references must name listed notes); notes show under the page body and edit by double-click; saved to footnotes.xml / endnotes.xml.
  - [x] Styles pane and Navigation pane: the desktop panes mount from Home / View; Styles pane definitions travel with the steps that apply them (or alone), are validated (identifier-safe ids/names, plain properties), merged like the desktop upsert, resolved into live styles before presentation, and saved to styles.xml.
  - [x] Table dialogs (Table Properties, Split/Insert/Delete Cells, Cell Margins) and Define New Bullet / Number Format / Multilevel List dialogs from the Ribbon and context menu; ListDialogs now imports the browser-only chrome Dropdown.
  - [x] Track Changes: the desktop recorder runs in the view (ins/del marks with author/date), Accept/Reject (current/all) run the desktop revision actions, Previous/Next select the revision; revision records may be cleared, never authored; saved as w:ins/w:del.
  - [x] Design theme fonts/colors as settings entries (presented through the theme CSS, saved to the theme part); View zoom applied to the page with the desktop Zoom dialog and fit modes; Spelling toggles the browser's native check on the editor.
  - [x] File → Convert to HTML / Markdown through a typed `docx.export` port: the desktop exporters render in the view, the host writes via a VS Code save dialog (Markdown pictures into `<name>_files`). PDF / image export and printing remain explicit (no print pipeline in the webview).
  - [x] Insert → Field: the desktop field insertion with whitelisted, in-document field types (PAGE, NUMPAGES, DATE, TIME, FILENAME, AUTHOR, MERGEFIELD …); INCLUDETEXT / INCLUDEPICTURE / DDE / LINK and markup are refused by the policy.
  - [x] Review → Protect Document: the desktop dialog sets editing restriction, password-to-modify (hash only) and remove-personal-information as settings entries; saved to settings.xml, reopened read-only when enforced. Password-to-open encryption remains explicit.
  - [x] Picture Format on file pictures: image blocks are no longer sealed placeholders; size, rotation, wrap and other picture attributes are edited as the desktop does, while anchors and raw XML stay immutable; resized extents saved. Charts, shapes and other objects stay sealed.
  - [x] Set Numbering Value (restart/continue via restart definitions) and Adjust List Indents (level edits of file or session lists, step-less entries) with the desktop dialogs; presented and saved to numbering.xml.
  - [x] References: sources (Manage Sources / Insert Citation / Bibliography) as a settings list saved to the sources part; Table of Contents, Index and Insert Caption blocks accepted only when their generated field XML matches the engine templates exactly (escaped text, in-document TOC / INDEX / SEQ instructions).
  - [ ] Picture bullets.
  - [ ] Remaining: comments, notes, headers/footers, page setup, revisions, protection, drawing, mail merge, print/export.
- [ ] Route all ribbon AI actions into native VS Code chat with document context.
- [ ] Implement real file, clipboard, picture, export/print and window host actions.
- [ ] Verify native Open/Save As dialogs and real-model chat/skill flows after complete ribbon integration.
- [ ] Investigate the view-closed cancellation diagnostic observed in the native DOCX revert/close regression.
- [ ] Verify both appearance modes, 21 locales, keyboard/focus behavior, popover layering and responsive ribbon densities.
- [ ] Prove manual/AI edits, shared undo/redo and save/reopen for every command family.
- [ ] Build and test installed Linux and Windows VSIX packages, including junction paths.
- [ ] Audit every ledger entry against rendered behavior and persisted output; finish only when all entries are verified.

## Partial Home checkpoint

- [x] Render the shared desktop XLSX Home ribbon with desktop styles.
- [x] Add a validated batch adapter for the initial Home command subset.
- [x] Connect the real Format Cells dialog with atomic apply and stale-state rejection.
- [x] Verify rich formatting projection, saved XLSX and native Undo/Redo.
- [ ] Complete every Home command and remove temporary legacy controls after parity checks.

## Relative Home commands

- [x] Apply mixed-cell indentation steps atomically using composed cell styles.
- [x] Share decimal stepping with desktop and verify General/explicit patterns, projection, save and history.
- [x] Match adjacent-donor geometry and formula offsets for directional fills.
- [ ] Copy source formatting for ribbon fills, including explicit removal of destination-only formatting, through projection/save/history.
- [ ] Support larger decimal selections with efficient grouped operations and the authoritative host range limits.

## Formatting fill checkpoint

- [x] Copy modeled donor cell styles for Home directional fills as one authoritative transaction.
- [x] Persist explicit style replacement so destination-only formats stay cleared after save and later edits.
- [x] Verify formatted/blank donors, date formats, formula shifts, display, saved XML and native Undo/Redo.
- [ ] Audit inherited row/column/default formats and source features outside the modeled cell-style schema before claiming full fill parity.

## Clear commands and axis dialogs

- [x] Connect Home Clear Formats with persistent format replacement and filtered-row exclusion.
- [x] Connect selected note/hyperlink removal and one-batch Clear All for modeled cell data.
- [x] Verify clear operations through actual menus, saved notes/hyperlinks and native history.
- [ ] Audit Clear All for conditional formats, validation, note images, filtered objects and large selections before full parity acceptance.
- [x] Extract and reuse the desktop row-height/column-width dialog with async apply and stale-state rejection.
- [x] Support point/character units, zero/maximum boundaries, character-width provenance and column insertion shifts.
- [x] Share Normal-font calibration and verify dialog sizes after actual save/close/reopen.
- [ ] Connect desktop AutoFit Row Height and AutoFit Column Width.
- [ ] Audit headless pixel-width requests for fonts requiring browser measurement, fractional character widths and all default-font variants.

## AutoFit measurement checkpoint

- [x] Route original Home AutoFit entries through desktop skeleton measurement and shared wrap/pitch corrections.
- [x] Hydrate required original pages, preserve source cells, and commit measured sizes in one host transaction.
- [x] Reject AutoFit results overtaken by concurrent edits, including before queued submission.
- [x] Click actual native VS Code webview menus and verify saved XML and independent native Undo/Redo.
- [ ] Finish automatic-row-height mode after subsequent edits, large-axis hydration and off-viewport measurement-extent parity before full AutoFit acceptance.

## Go To checkpoint

- [x] Reuse the desktop Go To dialog, address/name validation and frozen-pane reveal logic.
- [x] Navigate named regions and sheet-qualified references without authoring data or history.
- [x] Extend display coordinates for distant references and retain the selection through later resynchronization.
- [x] Share the desktop defined-name priority so live definitions precede REF residues.
- [x] Verify real keyboard editing after a jump and native VS Code menu navigation without document revisions.
- [ ] Audit scoped-name collisions, full-row/column/grid-edge targets, native Ctrl/Cmd+G, IME/focus/cancellation and disposal before full Go To acceptance.
- [ ] Reuse the desktop Find/Replace panel with authoritative complete-source search and replacement sessions.

## Find/Replace query and panel checkpoint

- [x] Query the complete immutable-source index with current overlays, clears and structural coordinate/formula-text transforms; do not fall back to viewport-only matches.
- [x] Bind paged query tickets to a document view and state, cancel obsolete queries, reject stale pages and keep source handles inside the host.
- [x] Reuse the actual desktop Find/Replace panel, input history, options and results list behind a presentation contract shared by both applications.
- [x] Navigate source matches through the shared Go To reveal path, preserve the query input focus and bind result-row selection to the replacement target.
- [x] Submit complete-text replacements through one authoritative manual transaction, preserve scalar types and reject stale, duplicate, protected and over-capacity writes.
- [ ] Complete large Replace All parity beyond the current ordinary 2,000-cell transaction limit; keep one atomic history action and full validation.
- [ ] Prove fresh calculated-value queries after edits and cover calculation dependencies, formatted values, hidden/protected cells and all desktop matching edge cases.
- [ ] Finish native keyboard shortcuts, selection-relative navigation, multi-view lifecycle/IME/focus audits and full search acceptance in both appearances/locales.

## Workbook shortcut and dialog placement checkpoint

- [x] Register native workbook Find/Replace/Go To commands and scoped Ctrl/Cmd+F/H/G keybindings through a strict presentation-only bridge.
- [x] Revalidate the active custom-editor tab, panel and session after waiting for readiness before dispatching a native command.
- [x] Render shared dialogs independently of the Home panel's visibility, preserving the currently selected ribbon tab.
- [x] Focus/select the query on repeated Find and retain an already selected matching cell during query research.
- [x] Verify actual Ctrl+F/H/G in Linux VS Code from Insert, two-match A80 anchoring, unchanged document state, and native Find after switching to a text editor.
- [ ] Audit macOS/Windows, F5 and other desktop shortcut variants, chat/terminal/modal focus precedence, loading/disposal races and non-matching selection-relative navigation before complete keyboard acceptance.

## Insert ribbon and hyperlink checkpoint

- [x] Render the actual shared desktop Insert command band; keep one persistent dialog adapter while portalling the header into the selected Home/Insert panel.
- [x] Record Insert command families and generated option sources in `sheets-insert-ledger.md`; retain incomplete rows as open work.
- [x] Extract/reuse the actual desktop Link dialog and target normalizer, with authoritative async apply and pending/cancel behavior in the plugin.
- [x] Read bounded current hyperlinks from source plus session overrides/removals and coordinate transformations, without mutating the source snapshot.
- [x] Bind link edits/removals to the dialog's frozen selection/state and verify validation, cancellation, concurrent-edit rejection, saved XML and exact native Undo/Redo.
- [ ] Complete link activation/grid decoration, every target form/length boundary, all appearances/locales and all other Insert commands before full Insert acceptance.
- [x] Close the original custom document, open a distinct new session, and prove the saved internal link and numeric cell value read back with revision zero and no dirty state.
- [ ] Trace the bare `Canceled` diagnostic observed in the close/reopen Link UI scenario before claiming clean cancellation/disposal acceptance.

## Insert Symbol checkpoint

- [x] Reuse the exact desktop Symbol categories/grid and append helper with async pending/close protection for hosted transactions.
- [x] Reuse Univer's desktop facade conversion, including number-format output and formula-looking text; apply value/format changes as one host edit.
- [x] Read the frozen current anchor, revalidate after the read and inside the queue, and advance repeated inserts only to their own committed receipt.
- [x] Verify ready/error/empty cached results and refusal of pending results or oversized output before producing writes.
- [x] Verify browser pending clicks/Escape, single-anchor range behavior, repeated inserts and concurrent-edit rejection.
- [x] Verify actual VS Code source text/number/boolean/date/blank/ready-formula cases, native history, delayed-acknowledgment race, authored-style preservation and save/close/reopen.
- [ ] Complete verified formula-result freshness after arbitrary edits/reopen, all input-conversion/locale cases, palette-wide acceptance and installed cross-platform behavior before full Symbol acceptance.
- [x] Resolve a selected merged-cell subordinate to its host-provided content anchor and prove only that anchor changes in real VS Code.
- [x] Include the full Symbol read/commit operation in editor flushing, avoiding a self-wait, and prove native Save waits for the accepted currency format.

## Insert Header & Footer checkpoint

- [x] Reuse the desktop six-section dialog and field-code insertion, with synchronous desktop and asynchronous hosted apply compatibility.
- [x] Read original odd sections using the shared desktop decoder and overlay canonical edits, including explicit null clears.
- [x] Validate before queueing a frozen-sheet/revision page-setup transaction; permit correction after invalid input and reject overtaken dialogs.
- [x] Verify browser prefill, cancel, selection-position field insertion, invalid-length retry, clearing and concurrent edit rejection.
- [x] Verify native pending/duplicate/close guards, exact Undo/Redo, saved XML preservation and distinct-session reopen alongside Link and Symbol regressions.
- [ ] Complete broad print preview/code-format fidelity, installed Windows/macOS, locale/appearance and lifecycle acceptance.

## Insert chart presentation and authoring checkpoint

- [x] Extract the desktop chart presentation without media, document or editor services, preserve the original desktop exports and replace the plugin's four-type BasicChart renderer.
- [x] Use shared drawing for column, bar, line, area, pie, scatter, radar, doughnut and combo charts; keep desktop rendering helpers, palette, legends, labels, axis text/layout and series selection callbacks.
- [x] Share desktop manual chart construction with authoritative add_chart, including numeric-year and scatter category parsing, source references, anchors and unlabeled insertion defaults.
- [x] Remove the session four-type insertion gate and serialize all nine supported plot sets through the existing drawing writer, rejecting unrecognized plot sets.
- [x] Connect every ordinary desktop Insert chart button and the actual Recommended Charts dialog with captured sheet/state/revision, complete bounded source reads and pending formula rejection.
- [x] Await chart source-read/commit work during editor flush; guard recommendation pending clicks, close, errors and stale submissions.
- [x] Render saved source charts using the shared presentation when opening a new session.
- [x] Verify browser chart insertion beyond the initial viewport and recommendation cancel, pending/flush, duplicate and concurrent-edit behavior.
- [x] Verify native all-type save/history/reopen, repeated saves without duplicate charts, full browser regressions and invalid-input recovery.
- [x] Trace normal-editor selection synchronization stealing ribbon-tab focus; guard the owned render instance, restore it on disposal and verify pointer release and explicit Go To focus.
- [ ] Finish in-grid chart anchors/drag/resize, source-chart editing/deletion/hydration, contextual design/format menus, complete chart options, formula freshness, locales/appearances and installed cross-platform acceptance.

## Floating chart frames and original drawing transactions

- [x] Reuse the actual desktop floating installer, frame selection, drag/resize handles, marker conversion and delete affordances inside the workbook grid.
- [x] Add the drawing preset without pulling an embedded AI UI; expose the existing shape path helper through browser chrome and verify the complete visual module boundary.
- [x] Add validated set_visual_anchor transactions for session charts and original drawings, with sheet ownership, protection, immutable source metadata and atomic inverse state.
- [x] Resolve original drawing locators in the host and persist source anchor/removal patches through the surgical drawing writer.
- [x] Support atomic deletion of mixed original/session chart selections and reject a drag overtaken by a concurrent transaction.
- [x] Wait for drag completion and visual transactions during editor flushing; cancel gestures through Escape, pointer capture loss and window blur.
- [x] Remove duplicate sidebar chart previews and retain temporary chart property controls in a collapsed section.
- [x] Reset native Webview body padding so the grid's measured viewport matches the actual editor area.
- [x] Verify native move/resize, exact Undo/Redo, byte-preserving move save, distinct-session geometry and cascading deletion with no dangling references; pass the final 29-scenario browser regression.
- [ ] Finish Chart Design/context menus and panes, original chart property editing/cache hydration, RTL/zoom/frozen-pane/large-anchor acceptance, cross-Webview pointer lifecycle and installed platform verification.

## Shared inline chart properties checkpoint

- [x] Replace the plugin's temporary chart property forms with the actual desktop inline ChartEditor for session-added charts.
- [x] Route title, supported chart conversions, series colors and pie/doughnut point colors through frozen authoritative edit_chart transactions; retain advanced chart types during title/color changes.
- [x] Keep rejected drafts editable, disable controls and cancellation during apply, prevent duplicate submissions and include the full commit in editor flushing.
- [x] Release the inline-editor installation guard during accepted hosted applies, preserving drafts during unrelated updates and rejecting overtaken submissions.
- [x] Preserve authored series and point colors in the session chart writer and verify repeated save/reopen through the native XLSX sidecar.
- [x] Verify native shared-panel pointer interaction, exact history, saved XML and distinct-session reopen; pass the final 31-scenario browser regression.
- [ ] Connect original-file chart property edits, source ranges/cache hydration, Chart Design/context menus and panes, all remaining chart options, locales/appearances and installed-platform acceptance.

## Original chart property transactions

- [x] Resolve original edit_chart ownership from trusted metadata and reject unknown/ambiguous parts, protected sheets, removed visuals and invalid series/points.
- [x] Store validated source property overlays with nested color merges, leaving original cells, metadata and chart parts immutable until saving.
- [x] Apply source overlays in the shared grid presentation and use the actual desktop inline editor for title, supported type conversion and colors.
- [x] Replay source chart overlays through surgical XML patches and omit patches for deleted drawings.
- [x] Verify original-file editing and repeated-panel state echoes in the browser, including the complete 32-scenario regression.
- [x] Verify real VS Code original-chart manual/agent-origin history, surgical save/reopen and deletion alongside native regressions.
- [ ] Complete source range/cache hydration, contextual chart menus/panes, AI tool discovery of current original charts, all chart parameters, appearances/locales and installed-platform acceptance.

## Chart series ranges and cache hydration

- [x] Reuse the desktop inline series-range inputs for original and session-added charts, with bounded authoritative reads and canonical sheet-qualified references.
- [x] Resolve separate values/category sheets, merge same-series entries and preserve references together with host-derived caches; never trust browser-supplied point arrays.
- [x] Reject non-vector/oversized ranges, missing material, pending formula results and outdated range submissions atomically.
- [x] Keep journaled original series caches live across later scalar edits, preserve nested property edits and stop obsolete name references from overwriting explicit series renames.
- [x] Track the entire inline range-read/apply lifecycle during flushing, including duplicate clicks, held reads, pending cancellation and disposal cleanup.
- [x] Hydrate missing original chart values/categories/names for presentation without authoring source metadata or revisions.
- [x] Verify actual native Save waits while a clean document's range read is held, then saves cross-sheet references/caches; prove native history, distinct-session reopen and cascading removal.
- [x] Pass final browser (35), session (121), editor (18), boundary (10), native VS Code (3), TypeScript and shared desktop/plugin build checks for this checkpoint.
- [ ] Complete live refresh for all untouched/reopened original chart references, broad formula freshness and typed-value cases, full Select Data/Chart Design/Format panes and chart parameters, AI tool chart discovery, locales/appearances and installed-platform acceptance.

## Desktop chart context and format pane

- [x] Reuse the actual desktop chart context menu and persistent ChartFormatPane through browser-only chrome imports.
- [x] Route legend, labels/position/number format, gridlines, value-axis bounds, gap/hole size and series/point explosion settings through validated canonical edits and source/session serialization.
- [x] Validate effective partial axis bounds and point extents, including uncached references; clear incompatible settings when converting chart families.
- [x] Keep checkbox/dropdown proposals visible while pending; disable overlapping edits/close, preserve rejected text drafts, and flush active pane inputs before Save.
- [x] Preserve pane input focus during hidden-editor synchronization and restore chart-element picking when the floating frame captures the pointer.
- [x] Add the axis-range validation message to all 21 locale dictionaries; keep existing desktop control labels.
- [x] Verify targeted context/pane/doughnut keyboard and real pointer tests, original native format/history/save/reopen and repeated session-chart format serialization.
- [x] Pass final browser (38), session (124), native VS Code (3), editor (18), boundary (11), type/lint/format and desktop/plugin build checks for this checkpoint.
- [ ] Connect full Select Data, Chart Design/Format ribbon bands, visual alt text/arrange/size commands and remaining contextual controls; disabled entries stay explicit pending work.
- [ ] Audit noncanonical index keys, all plot/parameter/input edge cases, live original cache/formula freshness, appearances/locales and installed platforms before full chart/menu acceptance.

## Desktop Select Data dialog

- [x] Mount the actual shared SelectDataDialog for original and session-added charts, with series add/remove/rename/reorder and Switch Row/Column.
- [x] Submit declarative source indices and range references; generate all replacement caches in the host from the effective current chart and verified material.
- [x] Preserve source-series colors through reordering, discard obsolete index/point overrides and refresh replacement references after later cell or inline series edits.
- [x] Freeze the dialog's root/revision, await the full canonical commit, prevent duplicate/cancel/edit actions while pending, retain rejected drafts and make Save wait.
- [x] Verify source/session candidate operations and actual browser interactions, and eight supported original-chart families through two independent saves and sidecar reopens.
- [x] Verify real VS Code Select Data add/reorder/rename, exact native history, Save/XML and distinct-session reopen; keep the pre-existing intermittent format-focus/cancellation audit open.
- [ ] Complete broad plot/parameter/input edge cases, untouched/reopened original live-cache/formula freshness, actual AI discovery/chat flows, all locales/appearances and installed-platform acceptance.
- [ ] Continue Chart Design/Format contextual ribbon bands and remaining visual commands; this checkpoint does not complete the full menu-replication goal.

## Desktop Chart Design contextual ribbon

- [x] Show the actual desktop Chart Design command band when a floating chart is selected; retain dynamic keyboard tab navigation and return to Home after deselection/deletion.
- [x] Share the desktop chart command mapping with the plugin instead of maintaining separate quick-layout, palette, grouping and title implementations.
- [x] Connect chart elements, legend/label presets, quick layouts, per-series/per-point palettes, grouping, supported conversions, Select Data, transpose, format pane and deletion to canonical transactions.
- [x] Extract and reuse the actual desktop title/axis-title dialog; freeze its target/root, await commit, retain rejected drafts, guard pending edits/cancel and participate in Save flushing.
- [x] Verify original/session real browser interactions, current-state menu checks, palette/family capabilities and chart title concurrency protection.
- [x] Separate background backup flushing from explicit chart-draft commit; prove an actual native backup retains focus/draft/root, then native Save commits the text.
- [x] Pass the final native combined suite (3), complete browser suite (45) plus backup-refinement browser checks (3), shared boundaries (12), editor units (18), type/lint/format and desktop/plugin builds.
- [ ] Complete the remaining contextual Format/arrangement/size/alt-text controls, broad family/parameter edges, locales/appearances and installed-platform acceptance; full menu replication remains open.

## Shared visual properties and layer menu

- [x] Add bounded canonical `edit_visual` metadata/layer edits for editor visual IDs, including protection/ownership/removal checks and source drawing/session model persistence.
- [x] Mount the actual desktop Alt Text and Size and Properties dialogs for source/session charts; retain desktop chart placement controls and existing chart resize handles.
- [x] Await dialog commits, disable pending controls/cancellation, retain rejected drafts, freeze root/revision and join native Save flushing.
- [x] Reuse desktop layer-order projection and connect Bring to Front/Forward and Send to Back/Backward menu commands as one transaction.
- [x] Preserve prior metadata through later geometry edits; replace deleted visuals with a removal tombstone only.
- [x] Verify canonical source/session metadata/layer changes and atomic rejection cases, real property/order browser interactions and repeated source/session XLSX saves/reopens.
- [x] Verify actual overlap/occlusion and mouse hit targets, native property dialogs/history/save/XML/reopen and agent-origin visual editing history.
- [x] Pass full browser (49), session (131), native VS Code (3), desktop visual units (41), boundary (13), CLI (12), type/lint/format and desktop/plugin build checks.
- [ ] Finish broader shape/image presentation and commands, align/distribute/rotate/flip, applicable size controls, hyperlinks/clipboard and the remaining menu/DOCX/locales/platform acceptance.

## Source shapes, pictures and contextual formatting

- [x] Display source shapes and pictures with the actual shared desktop visual renderer and enable the matching Shape Format/Picture Format contextual bands.
- [x] Read floating-source media through a validated identity/root request; keep source sessions/paths in the host and reject stale, foreign, removed and ambiguous identities.
- [x] Connect shape fill/outline, clockwise/counterclockwise rotation, horizontal/vertical flip and the shared numeric size/properties dialog to canonical source drawing edits.
- [x] Extract and reuse desktop frame-center/rotation/anchor calculations; commit geometry and frame properties together as one revision and preserve prior visual metadata.
- [x] Connect multi-selection align/distribute through desktop geometry and one canonical batch; reject unavailable/stale selections instead of partially arranging them.
- [x] Keep unsupported clipboard/hyperlink/embed context entries disabled while their host paths remain unfinished.
- [x] Verify actual shape/picture rendering and decoding, paint/transform/size failure-retry-Save behavior and physical align/distribute geometry in browser tests.
- [x] Verify native source shape/picture history, saved transforms/frame dimensions, unchanged source media bytes, independent-session reopen and existing native chart/shortcut regressions.
- [ ] Complete links/clipboard/drop/embed paths, remaining source families/parameters, other ribbon tabs, DOCX, locale/appearance and installed-platform acceptance.

## New session shapes and text boxes

- [x] Reuse the desktop Shapes gallery and drawing gesture, including click defaults, Shift square, Escape/window-blur cancellation and text-box cell defaults.
- [x] Store session shapes and charts in one ordered canonical visual list, with unique allocated IDs, chart compatibility projections and shared formatting/history.
- [x] Validate explicit shape anchors and ownership/protection; reject drawing gestures overtaken by a newer root.
- [x] Await shape text commits, retain rejected drafts and include focused text in explicit Save while background backups preserve drafts.
- [x] Save authored text/paint/transforms/frame properties and mixed layer order; split additions into bounded batches and verify 54 objects through repeated saves/reopens.
- [x] Pass full browser (53), native VS Code (4), session (138), mixed-object save/reopen, type/lint/format and desktop/plugin build checks; record results in verification.md.
- [ ] Complete clipboard/drop/links/embed, remaining ribbon tabs, DOCX and broad locale/platform acceptance; full menu replication remains open.

## New floating pictures

- [x] Extract and reuse the desktop picture picker/read/decode path, accepted types, 20MB per-file validation and aspect-ratio cell sizing.
- [x] Add canonical inline `add_picture` operations and ordered session picture models with stable allocated IDs; connect the actual Insert Pictures control.
- [x] Reuse the shared Picture Format renderer/menu for session pictures, canonical geometry/transforms/properties/layers, deletion and manual/AI undo history.
- [x] Execute picture-bearing candidate roots with the same shared interpreter in the host, avoiding duplicate large-byte candidate transfers; test the 20MB boundary and subsequent transforms/deletion.
- [x] Verify actual browser file selection/cancellation, invalid files, delayed commit/Save, image decoding, formatting and selection overtaken by another root.
- [x] Verify mixed pictures/shapes/charts across two independent saves and sidecar reopens with exact picture bytes, authored transforms/frame dimensions and layer order.
- [x] Pass final full browser (54), session (140), new native picture flow and three existing native regressions, type/lint/format/schema checks and desktop/plugin builds; record results in verification.md.
- [ ] Complete graphical clipboard/link/embed paths, additional Insert commands, remaining shared ribbon tabs, full DOCX menus and broad locale/platform acceptance.

## Shared picture transfers and native external cell paste

- [x] Reuse the desktop picture paste/drop installer and its image-versus-table/text rules in the plugin grid; scope native edit targets to the embedded container.
- [x] Extract and share the actual desktop scene/scroll/scale hit test for drop anchors; freeze the owning sheet, anchor and revision before asynchronous image reads.
- [x] Include clipboard/drop image reads and commits in Save flushing; remove listeners on disposal and reject overtaken roots without an extra image.
- [x] Initialize the desktop clipboard controller after workbook creation, so hidden native toolbars do not leave the default copy/paste hooks absent.
- [x] Delegate table/text payloads to the desktop parser without preview image files; capture native mutation plans and commit supported values/styles/merges as one canonical batch without native projection/history writes.
- [x] Preserve boolean cell types and computed CSS clipboard colors through the shared neutral-style converter; localize new clipboard errors in all 21 app shards.
- [x] Verify real Chromium clipboard Ctrl+V, exact drop anchors, input-focus exclusions, mixed table previews, stale drops, delayed Save and rejected external-cell paste retry.
- [x] Verify transfers/parser in 5 native scenarios, editor/shared units, boundary, type/lint/format and desktop/plugin builds; record the full browser attempt and its isolated picker repeats.
- [ ] Resolve the intermittent picker timeout seen in the 57-test full browser run (56 passed); four standalone repeats pass without a timeout increase.
- [x] Connect the actual Home Paste button to the same image/table transaction paths, track clipboard-read Save work and freeze the target/root; verify read denial and native Ctrl+V recovery.
- [x] Verify the Paste button in the native transfer scenario, including exact history and three-image/table save/reopen; pass final builds, type/lint and targeted browser checks.
- [ ] Finish cut/internal-source freshness, clipboard-cache fallback, Paste Special, mixed rich text, additional native mutation families/automatic row sizing, other Insert/base tabs, DOCX and broad locale/platform acceptance.

## Typed same-editor Copy snapshots

- [x] Connect the actual shared Home Copy button and grid Ctrl/Cmd+C to an authoritative bounded source read and the desktop clipboard serializer/cache/marquee.
- [x] Retain a detached typed snapshot identified by the native copy ID across source edits and projection reinstalls; keep Copy free of document history entries.
- [x] Add bounded `paste_cells` transactions preserving literal formula-looking strings, errors, dates, replacement styles and independent border edges; shift formulas and invalidate copied caches.
- [x] Replace contained destination merges, retain outside merges and tile compatible selected targets; reject partial merges, missing source data, protected targets, mixed date systems and grid/write-budget overflow.
- [x] Verify snapshot/menu/native-key paths in unit/browser tests and cross-sheet native VS Code Copy/Paste with exact Undo/Redo, saved XML and independent-session reopen.
- [ ] Complete reference-aware Cut, Paste Special, copied metadata/rich text, filter/hidden-row semantics, cross-editor snapshot portability, copy-denial reporting and clipboard-cache fallback.
- [ ] Complete broader base/Insert/DOCX ribbons and locale/appearance/installed-platform acceptance; full menu replication remains open.
- [x] Pass the final 61-scenario Home/manual browser suite, 142 session tests, 303 gateway tests, focused shared/CLI/boundary checks, final TypeScript/lint/format and desktop/plugin builds; retain the earlier intermittent picker audit as open.

## Atomic cell Cut and reference movement

- [x] Connect shared Home Cut and grid Ctrl/Cmd+X to authoritative source reads and the native CUT serializer without modifying the document before Paste.
- [x] Add a canonical range relocation operation with complete dependency materialization, source/destination/dependent protection, source capture before clearing, overlapping moves and one reversible transaction.
- [x] Move typed cells, replacement formats, contained merges, plain notes and hyperlinks; follow absolute/relative contained A1 dependencies and retain outside dependencies for cross-sheet moved formulas.
- [x] Preserve formula string/table/external/3D tokens and the desktop partial-range behavior; keep Copy's fill-offset semantics separate.
- [x] Reject stale cut roots atomically, consume successful cut identities, handle grid Escape cancellation and verify subsequent Copy recovery.
- [x] Verify actual browser Cut and native VS Code cross-sheet Cut/Copy controls, exact Undo/Redo roots, saved merge/note/link/error/formula XML and distinct-session reopen.
- [ ] Complete affected defined-name and native-chart transforms, table/pivot/sparkline/filter/CF/validation movement, rich notes and broader clipboard fidelity; interim rejects remain feature gaps.
- [ ] Complete Paste Special, cross-editor cut, clipboard denial/fallback, remaining base/Insert/DOCX ribbons and full locale/appearance/installed-platform gates.
- [x] Pass the full 64-scenario browser regression and four final Cut/error-retry cases, 148 session tests, 307 gateway tests, editor/shared/CLI/boundary checks and desktop/plugin builds; verify saved source clearing and dependent references in an independent native session.

## Shared Paste Special dialog and selective fields

- [x] Extract the desktop pure transform with compatible re-exports and reuse the actual dialog, defaults, source availability, arithmetic, transpose, skip-blanks and link logic.
- [x] Connect five quick entries, the full dialog and grid Ctrl/Cmd+Alt+V; freeze source/target/root, track read/apply work for Save, retain failed drafts and reject stale dialogs.
- [x] Extend bounded canonical typed paste with selective values/formats/number formats/except-borders and skipped addresses; preserve unrelated target fields and support more than 100 differently formatted cells in one operation.
- [x] Match desktop cached-value/current-format, advanced live-source and current-column-width rules; keep external basic HTML choices on the native clipboard hooks and external transformations on the shared text parser.
- [x] Cover each quick entry, arithmetic/transpose/skip, links, source edits, failed retry, delayed Save, stale roots, external source availability and computed-value-independent format/link combinations.
- [x] Verify native cross-sheet dialog use, exact Undo/Redo roots, selective style/number-format preservation, saved column-width values and independent-session reopen; run shared-transform, canonical and boundary regressions.
- [ ] Complete verified computation freshness, CF/validation/rich metadata, filtered/merged target parity, source deletion/structural edits, cross-editor ownership and all locale/appearance/platform gates.
- [ ] Finish the remaining base/Insert and DOCX desktop menu groups; the full goal remains active.
- [x] Pass the final 73-scenario browser run, 3 native clipboard regressions, 151 session tests, 307 gateway tests, 29 editor units, 131 shared desktop units, CLI/boundary/type/lint/format checks and both builds; keep lifecycle/performance warnings and earlier picker audit open.

## Shared Page Layout band and print metadata

- [x] Reuse the actual desktop Page Layout band, share command decoding and preserve desktop defaults for orientation, margins, paper, scale/fit, print area/titles and page breaks.
- [x] Read indexed source page settings and break metadata, map structural coordinates, freeze owning sheet/root and include asynchronous page work in Save flushing.
- [x] Keep screen grid/headings separate from print flags; preserve imported page breaks and make empty break actions history-free.
- [x] Fix shared fit-menu wrappers so normal native VS Code clicks close the menu and reach subsequent print controls.
- [x] Verify native Undo/Redo, saved print names/setup/breaks, unchanged cell XML/package sentinel and a distinct-session reopen.
- [x] Keep page-only transactions independent of chart calculation/cache refresh and unrelated chart-reference material; retain indexed protection checks.
- [ ] Complete document theme/color/font editing and the other base/Insert/DOCX menu groups, plus broad locale/appearance/installed-platform acceptance.
- [x] Pass full browser (77), final Page Layout browser repeats (5), final native save/history/reopen (1), session (154), gateway (307), shared/editor/boundary/type/lint/format checks and desktop/plugin builds; record history timeout/recovered-host observations without closing the performance audit.

## Shared Function Library and Function Arguments

- [x] Extract the actual desktop Function Library group and reuse live function descriptions, shared category buttons, catalog/search/recent history and Function Arguments UI.
- [x] Support awaited canonical writes with pending-action gating, retained rejected drafts, frozen owning target/root and Save waiting for commits.
- [x] Read bounded authoritative static precedents before preview, including off-page/cross-sheet values; decline unverified formula/name/dynamic/external previews rather than supply zeros.
- [x] Share the desktop per-column AutoSum geometry and submit selected columns as one canonical operation batch.
- [x] Scope argument pointer picks and Shift+F3 to the embedded grid; initialize/dispose the native fx listener with each ready workbook projection.
- [ ] Finish Defined Names, Formula Auditing and Calculation settings/menu groups; the shared Function Library does not complete Formulas.
- [ ] Complete verified computation freshness, full caret/cancel/formula-language behavior, locales/appearances/installed platforms and remaining base/Insert/DOCX menus.
- [x] Pass the completed full 87-case browser run, final native function save/history/reopen, 34 editor units, 51 shared units, boundary/type/lint/format checks and desktop/plugin builds; keep the remaining Formulas groups and broad acceptance open.

## Canonical defined-name foundation

- [x] Enable authoritative add/delete names with stable logical sheet scope, source preservation, per-scope case-insensitive duplicate checks, workbook protection and atomic batches/inverses.
- [x] Serialize current name state through the declarative writer; preserve hidden/built-in/unmodeled names and verify real XLSX save/reopen with unchanged source cells and package sentinel.
- [x] Share desktop name installation with the plugin projection and read current journal names in Go To; mark old source formula caches pending and keep metadata-only edits independent of chart cell material/cache refresh.
- [x] Reject logical-scope variants in standalone adapters rather than silently applying global scope; regenerate native tool schemas.
- [x] Verify native agent-origin name transactions, Go To synchronization, exact native Undo/Redo and independent-session reopen.
- [x] Implement ordered Name Manager rename/removal reference plans with action-time local/global bindings; connect shared dialog actions, Define Name, Create from Selection and Use in Formula.
- [x] Preflight hidden/empty source name bindings before accepting reference-aware edits; reserve source scopes without exposing hidden formula content.
- [ ] Complete destination-shadow qualification, scoped calculation and broader locale/platform/appearance acceptance; full menu replication remains active.
- [x] Pass the final 88-case browser suite, 158 session tests, 309 gateway tests, 14 CLI tests, 17 native-tool/boundary units, native name save/history/reopen, type/lint/format checks and desktop/plugin builds; retain outstanding Name Manager reference/UI and broad acceptance tasks.

## Shared Defined Names group and source reference replay

- [x] Reuse the actual desktop Defined Names group and Name Manager with awaited canonical actions, retained rejected drafts and disabled pending controls.
- [x] Rewrite existing overlay formulas at action time; replay original-source references on read/save while preserving later formulas, literals, table/external tokens and hidden local shadows.
- [x] Preserve hidden/empty definitions and foreign worksheet/cell extensions; decode CDATA formula text before rewriting.
- [x] Share Create from Selection label cleanup and A1 helpers; read formatted authoritative labels and commit over 100 definitions as a single bulk operation.
- [x] Reject overtaken reads/dialog roots and make Save await pending definition work; exercise retry, scope, left/top labels and blank-only Use in Formula in browser tests.
- [x] Verify 98 full browser cases plus the interrupted startup and seven name cases in an 8-case repeat; pass final native UI/save/reopen, 161 session tests, 313 gateway tests and desktop/plugin builds.
- [x] Retain structural protection for reference-change logs and file-only bindings after visible names are removed; preserve extension content during formula-cache invalidation.
- [ ] Complete qualified destination shadow handling, reference plans across structural edits and calculation parity without reducing the full desktop-menu goal.

## Shared Formula Auditing

- [x] Extract desktop Formula Auditing controls and pure formula-reference/precedent/dependent parsing into shared modules.
- [x] Add an independent validated source.audit bridge for complete-source dependency/error queries without consuming Find/Replace tickets or mutating history.
- [x] Apply ordered name-reference replay in original formula searches; disclose pending formula values after definition changes.
- [x] Connect actual arrows, error navigation, removal and Watch Window to frozen current-state reads, including off-page/cross-sheet sources.
- [x] Implement per-sheet Show Formulas through canonical page settings, shared display interception and grid-scoped keyboard handling; serialize it in session and standalone adapters and explicitly reject it in the older desktop DSL adapter.
- [x] Verify six browser cases, 165 session tests, 313 gateway tests and shared parser/search regression; record and correct the pre-workbook service registration error.
- [x] Verify native UI/history/save/reopen and final desktop/plugin builds, types, lint/format, editor units and CLI/tool/boundary checks.
- [x] Establish clean complete current browser acceptance: 106 pass after the protected-view/bridge changes, including the shape case.
- [ ] Diagnose the historical intermittent default-shape click failure; a later clean full run does not prove its root cause resolved.
- [x] Align protected-sheet Show Formulas/gridlines/headings/zoom with a separate indexed host view grant, strict bridge validation and normal content/readonly restrictions.
- [ ] Complete automatic/manual calculation settings, verified formula-result freshness, large-file/performance and all other menu/locale/appearance/platform requirements.

## Protected worksheet view permissions

- [x] Distinguish host-verified view write access from ordinary sheet write access; keep the new grant optional and denied by default when missing.
- [x] Wait for complete source protection indexing before granting view access; keep readonly, mixed print/view patches and content batches under their ordinary restrictions.
- [x] Validate viewWritableSheetIds explicitly in the strict candidate bridge and verify browser results by authoritative host execution.
- [x] Prove real source/save/reopen retains the protection element/password, source sheetData and package sentinel; verify native Undo/Redo on both protected and ordinary sheets.
- [x] Pass 170 session tests, 313 gateway tests, 34 editor units, 7 targeted browser cases, 106 complete browser cases and 2 native scenarios; pass plugin build/types/lint/format checks.
- [ ] Continue broader desktop permission mapping for print/layout and readonly presentation under the remaining Review/View work; the view grant does not redefine those requirements.
- [ ] Implement the actual Calculation group with canonical cache/freshness semantics; retain the concrete Node experiment and unresolved result/error/array findings in calculation-notes.md.

## Host calculation adapter foundation

- [x] Implement a host-only, DOM-free adapter over authoritative hydrated Univer snapshots, using the same function engine and real bundled Node/CommonJS tests.
- [x] Capture rich results during actual engine evaluation; preserve error/text provenance, booleans, numbers, blank-reference semantics, full spills, computed format hints and link/image metadata without modifying input.
- [x] Mirror desktop workbook/Calculate Sheet requests, including cross-sheet dependents and exclusion of unrelated cached formulas from a scoped result.
- [x] Isolate per-job engine IDs to prevent module-level index/cache cross-contamination; keep stable caller IDs/order in returned data.
- [x] Reject cancelled/over-budget work without publishing partial results; preserve START cancellation across engine resets and dispose runtime observers.
- [ ] Complete frozen-root source hydration, hidden/scoped names, table/external resources, array ownership, date-system/time-zone and function coverage before treating results as document-authoritative.
- [ ] Implement worker/resource bounds and job lifecycle/stale-result rejection, then canonical cache/read/search/watch/chart/preview/save integration.
- [ ] Connect the actual shared Calculation group, automatic/manual mode, F9/Shift+F9, projection/undo/cache behavior and saved calcPr/results with native history/reopen and platform acceptance.

## Complete-source calculation loading

- [x] Add native internal stored-cell enumeration, including styled blanks, source array ownership, complete-index bounds and pagination, without changing ordinary Find semantics.
- [x] Load all live original sheets and new-sheet overlays through bounded sparse current-state reads; preserve exact typed records, explicit clears/styles and coordinate/name rewrites without mutating the caller root.
- [x] Validate scope/order/cursors/grid/completeness, reject missing enumerated source records and stale/cancelled jobs, and enforce explicit operation cell/byte budgets with no partial success.
- [x] Keep source-space features/array ownership distinct from logical overlays; retain ownership even for overridden/deleted array anchors.
- [x] Carry source table ranges and column names into the private calculation source; register structured-reference tables through the formula engine service and reject table layouts it cannot represent exactly.
- [ ] Verify structured-reference formulas against native XLSX tables, including totals/header variants, table calculated columns and row/column edits.
- [x] Verify real native source loading beyond row 100000 and beyond the 10000-result page boundary after row insertion; pass 202 session tests, 313 gateway tests and 276 native tests.
- [x] Compose private hidden/scoped definitions and representable table resources into a normalized source snapshot; preserve forced text values and reject currently unsupported source resources.
- [ ] Bind legacy array ranges and external workbooks; complete row/column/default/date/time-zone and full style semantics, then verify the normalized snapshot against real workbooks.
- [ ] Add worker isolation, root/job-owned derived cache publication and complete shared Calculation controls/modes/history/save/reopen behavior. The overall desktop-menu goal remains open.

## Private calculation resource context

- [x] Add a session-scoped host-only native resource command without adding hidden definition bodies to public open metadata or renderer/AI routes.
- [x] Preserve hidden/empty/built-in/global/local definitions with stable original scope IDs, raw calcPr attributes and ordered unresolved external relationship IDs.
- [x] Parse namespace/path-aware XML, CDATA/entities and prefixed SpreadsheetML; reject invalid scopes, nested definition markup, DTDs and foreign relationship namespaces.
- [x] Require and validate private context in the complete-source loader, enforce ownership/cancellation/budgets, and include context reads in client close cancellation.
- [x] Verify real source context and public-metadata separation after sheet reordering/renaming, plus cancelled/closed and malformed native reads.
- [ ] Complete calculation-time local/global name resolution and transformation replay; installed engine name caching currently lacks a worksheet scope key.
- [ ] Complete table/array/external/date/default binding, worker isolation, canonical cache/freshness and actual Calculation controls/mode/history/save/reopen acceptance.

## Normalized calculation name binding

- [x] Isolate resolved global/local name bindings with private aliases; verify qualified/quoted/Unicode/workbook names, global fallback, nested expressions and local range ownership through the actual engine.
- [x] Preserve LET initializer/sequential scope, LAMBDA parameters/named calls, built-in functions, strings, unary operators and percentages during computation-only binding.
- [x] Resolve direct dynamic INDIRECT name strings by runtime sheet/case/qualifier through the public address-function lookup and restore the hook on disposal.
- [x] Freeze caller definition values; verify same-ID concurrent jobs and mutation from progress callbacks cannot alter another calculation's bindings.
- [x] Expand shared followers before binding/provenance capture using public offset APIs; reject missing/ambiguous masters and verify no private aliases appear in returned formulas or authored input.
- [x] Compose private source definitions/scopes with canonical definition/reference journals; verify logical scope after sheet reorder, action-time journal text, hidden dependency replay and synthetic save/reopen calculation.
- [ ] Lift structural-operation guards only after named-definition and reference bindings are transformed with exact Excel behavior; cover legacy recovered states without stable source sheet IDs and prefixed writer XML variants.
- [ ] Complete recursive/iterative names, dynamic named-expression coverage, table/array/external/date/default context and worker/canonical cache/menu/save acceptance. Current circular-binding rejection is not completion of those requirements.

## Shared Calculation ribbon controls

- [x] Render the shared desktop Calculation Options, Calculate Now and Calculate Sheet controls in the plugin Formulas tab.
- [x] Reuse the desktop automatic/manual calculation veto, active-sheet calculation request and F9/Shift+F9 behavior in the plugin editor runtime.
- [ ] Bind manual/automatic mode and calculation results to the host-owned workbook root, saved `calcPr`, canonical freshness and history/save/reopen behavior; current plugin controls still operate on the editor projection.

## Data ribbon implementation follow-up

- [x] Reuse the desktop AdvancedFilterDialog with controlled locale, asynchronous apply, retained conditions after rejection and pending close protection.
- [x] Extend canonical set_filter_criteria with exclusive value/custom conditions, preserve other column criteria, and share raw-value condition predicates between streamed desktop filtering, session evaluation and CLI new-filter output.
- [x] Read original/current filter metadata and header labels from the frozen host root, bind dialog apply to that version and integrate pending work with Save flushing.
- [x] Compose filter-hidden rows into current-range and complete-outline reads before explicit visibility overrides.
- [x] Verify schema/predicate/session/desktop/CLI/read regressions, browser AND/OR/retry/stale/pending behavior and native raw-number matching/history/save/reopen/imported-condition preservation.
- [x] Reuse the actual desktop Sort & Filter group and shared large-menu rendering; route Sort variants, Custom Sort, Filter, Clear, Reapply and Advanced through plugin adapters.
- [x] Add distinct canonical clear_filter_criteria and reapply_filter operations, material hydration and formula preflight, desktop/CLI translation, generated native tool schema, shared history and save support.
- [x] Reuse desktop smart filter-region expansion over complete frozen source halos for single-cell/header selections, including merged anchors, bounded reads and stale-state rejection.
- [x] Bind Ctrl/Meta+Shift+L with workbook/modal/input/editing gates and include smart planning/source reads/submission in Save flushing; verify Linux native keyboard behavior.
- [x] Fix shared Dropdown option clicks reopening through a wrapping label, preserving the closed state before subsequent condition selection.
- [x] Verify region/session/CLI/Dropdown regressions, browser existing/new menu flows and native shortcut/AI edit/reapply/clear/remove/imported range/history/save/reopen alongside Custom Sort/Advanced.
- [ ] Complete larger/complex filter variants, formula/result freshness, date/error/protection/hidden/locale/appearance/platform acceptance, native header popups and remaining Data groups; remove temporary value-list controls after popup parity.

- [x] Add canonical row/column outline-level, collapsed-flag and summary-placement operations; journal them in the session and replay through the existing original-based OOXML writer.
- [x] Verify outline XML, source-cell/package preservation and reopened native outline metadata with the real XLSX sidecar; reject spans/levels outside Excel limits.
- [x] Read complete source row/column outlines, transform earlier source-coordinate shifts, replay the session journal and expose a revision-checked JSON outline read port without relying on viewport hydration.
- [x] Connect relative Group/Ungroup Rows and Columns using complete host outline reads and the shared desktop level planner, plus summary-placement state echoes and Save flushing.
- [x] Verify nested row/column grouping, summary flags, eight native Undo/Redo steps, saved XML and independent clean-session reopen in real VS Code.
- [x] Add canonical row/column visibility journals, current-range/outline replay, grid projection and original-based XLSX serialization for detail folding.
- [x] Connect Hide/Show Detail and Clear Outline using desktop nested-group plans, summary collapsed flags and one authoritative history entry per action.
- [x] Decouple the actual desktop outline gutter from its Electron/lazy-state adapter; mount the shared canvases, preserve header sizing and bind exact-group/depth clicks to frozen host transactions.
- [x] Verify nested grid visibility and real shared-gutter group/depth mouse clicks in the browser, including simultaneous row/column outlines and removal of the reserved strip after clear.
- [x] Fix embedded body minimum-width and ribbon-height overflow; verify shared gutter/depth clicks in a 591×365 browser viewport and native row/column depth-button history/save/reopen.
- [ ] Finish native exact-group glyph/lifecycle/geometry acceptance, imported-outline variants, filter/protection/frozen-pane/zoom/appearance/locale cases and installed Windows behavior.
- [ ] Implement outline/visibility coordinate transforms before lifting the structural-edit guards.

- [x] Extract the bounded Goal Seek secant algorithm into the shared gateway, retaining desktop restore-on-failure behavior and testing convergence, finite-value checks, evaluation bounds and cancellation propagation.
- [x] Add host-only Goal Seek over a private calculation-source copy, using the real Node formula engine for scalar trials, root/cancellation checks and bounded solve lifetime; verify cross-sheet convergence and unchanged source data.
- [x] Build frozen host Goal Seek evaluation and a single final canonical input-cell commit; connect the actual desktop dialog without publishing intermediate guesses.
- [x] Add the strict Goal Seek bridge request and host controller with permission/merge preflight, captured revision, private-source loading, one final canonical write, nonconvergence/no-op handling and disposal cancellation.
- [x] Connect the shared Goal Seek dialog to the host controller, track pending solves during Save and verify native history/save/reopen with real source formulas.
- [ ] Complete Goal Seek acceptance for nonlinear/error/no-solution UI cases, permission and lifecycle variants, calculation resource coverage, locale/appearance and installed Windows behavior.

- [x] Extend the shared sort operation and stable comparator with up to three ordered keys; use the same permutation for session value/style/type movement and desktop bulk execution.
- [x] Extract the desktop Custom Sort dialog and connect its three levels, header flag, labels, async retry and frozen selection to the plugin; prove history/save/reopen.
- [ ] Finish Custom Sort acceptance for all three-key variants, formulas, hidden/filtered and protected selections, appearance/locale and installed Windows behavior.
- [x] Extract and mount the actual shared SortDialog with controlled locale, three keys, header toggle, async disabling and error retention; browser-test key selection and retry as one canonical revision.

- [x] Add plugin Remove Duplicates action using the shared desktop row comparison algorithm, selection header option and one frozen host transaction.
- [x] Extract the actual desktop RemoveDuplicatesDialog, preserving the default header flag and 21-locale labels while supporting async apply, pending dismissal protection and retry with retained options.
- [x] Freeze duplicate-removal selection/root/revision on open, include source reads and submission in Save flushing, reject stale work and keep no-op removal out of history.
- [x] Verify browser retry/pending/cancel/no-op/stale/merge guards and native one-step removal/history/save/independent reopen with stationary/moved formulas, typed payloads, styles and unrelated source content.
- [x] Add plugin Flash Fill action using the shared desktop template inference/application algorithm and frozen, chunked host reads with a canonical transaction.
- [x] Add plugin Text to Columns delimiter actions matching the desktop tab/comma/semicolon/space split semantics, with canonical cell writes; a follow-up removes incorrect structural insertion based on the used extent.
- [x] Add a plugin Data Validation editor for list, list-reference, numeric/date range, checkbox and custom-formula rules; save through the canonical operation path and project new rules into the active grid.
- [x] Carry `set_data_validation` through the session candidate feature journal and package writer as an append/remove patch, preserving unrelated validation rules.
- [x] When a new data-validation range partially overlaps an existing sqref rectangle, subtract the covered cells and offset each surviving region's relative formulas from its new top-left cell.
- [x] Prefill the Data Validation dialog from the selected live rule for supported list, list-reference, decimal-between, checkbox and custom-formula rules; reject invalid imported values from the editor mapping instead of coercing them.
- [x] Add a generic Excel whole/decimal/date/time/text-length criterion with all eight comparison operators and optional second formula; persist it through the shared session operation and worksheet XML writer.
- [x] Edit and preserve data-validation blank handling, dropdown visibility, input prompt, error alert mode/style, titles and messages through the plugin dialog, Univer projection and XLSX save.
- [x] Resolve validation list sources backed by workbook/scoped defined names, table columns and cross-sheet A1 ranges into live dropdown references; allow the editor to retain a qualified source range.
- [x] Edit message-only (`any` / OOXML `none`) validation rules without converting them into a constrained list; preserve their prompt/error options through save.
- [ ] Complete Data Validation parity for extended OOXML rules and advanced workflows; implement What-if/Goal Seek, Outline and desktop Get Data/Pivot workflows; verify all commands in installed VS Code and saved/reopened workbooks.
- [x] Add the Data → Consolidate dialog with desktop aggregation functions, position/left-label modes, default selection reference and multiple source ranges.
- [x] Reuse the desktop consolidation matrix algorithm, read labels from the frozen host revision, and submit typed labels and live formulas as one canonical transaction with bounds and overlap checks; output currently shares the host 2,000-cell limit.
- [x] Verify Consolidate through real VS Code UI, exact native Undo/Redo, saved formula XML and independent reopen with a new session reading the formula and both original source values.
- [ ] Complete duplicate-removal acceptance for larger/sparse selections, dates inside compared/moved rows, formula error/result freshness, filtered/protected rows, all appearance/locale and installed platforms; the current command is limited to 2000 selected cells and ready scalar values.

Column-header filter popup follow-up (2026-10-28):

- [x] Mount the same Univer filter preset, styles and en-US/zh-CN/zh-TW resources used by the desktop sheet grid; restore source/canonical ranges, absolute-column criteria and cached filtered rows without evaluating a partial viewport.
- [x] Bind popup opening and changes back to By Values to frozen, complete host reads; retain authoritative formula results and reject stale/pending/over-budget data instead of building an incomplete value list.
- [x] Route value, blank, custom-condition and per-column-clear commands into one canonical transaction, with failed-draft retention, retry, duplicate-submit protection and Save flushing.
- [x] Separate filter-owned row visibility from manual/outline hd flags and remove the superseded column-number/value-textarea filter editor.
- [x] Browser-verify failed retry, pending duplicate clicks, Save waiting, clear/unhide, offscreen source values, stale reads, authoritative formula caches and native condition/value mode switching.
- [x] Verify native header checkbox clicks, one canonical revision, exact Undo/Redo, saved filter XML, unrelated package/style preservation and independent clean-session reopen in VS Code 1.109.5.
- [ ] Complete imported-popup editing and installed Linux/Windows, geometry/locale/appearance/protection/table-filter variants.
- [x] Add canonical fill/font color criteria and bounded XLSX persistence, including desktop multicolor round trips.
- [ ] Complete imported/conditional/rich-text color variants and filters exceeding the current 2,000-cell budget before accepting complete desktop filter parity.

Filter-popup sorting follow-up (2026-10-28):

- [x] Mount the desktop Univer sort preset/styles/locales and its actual embedded Ascending/Descending buttons inside the filter panel.
- [x] Bind popup sorts to the captured filter sheet, range, key column and revision; retain the header, keep the popup open on host rejection and make Save wait for the canonical submission.
- [x] Materialize current hidden-row metadata through the host and strict candidate bridge; sort only visible row slots while keeping hidden rows stationary and moving typed values/formats/styles consistently.
- [x] Preserve source-manual hidden rows in the visibility journal when filter edits first take ownership of the file visibility cache, so adding a filter and sorting do not unhide those rows during save.
- [x] Verify numeric Ascending/Descending, active criteria, noncontiguous visible slots, manual hiding, rejected retry and pending Save in the browser; verify source-hidden rows, native Undo/Redo, original-based XLSX/filter/style preservation and independent clean-session reopen in VS Code 1.109.5.
- [ ] Complete formula movement, notes/hyperlinks/validation semantics, merges, all sort types/locales/geometry/protection variants and installed Linux/Windows acceptance before declaring full sort parity.

Color-filter follow-up (2026-10-29):

- [x] Route the actual desktop By Colors swatches into canonical fill/font criteria, including multiple checked colors, no fill, default black, RGB normalization, cross-column AND and style-only Reapply.
- [x] Carry colors through the DSL, strict range/filter schemas, candidate interpreter, desktop restore/save/operation adapter, CLI and regenerated VS Code native tool schemas.
- [x] Serialize single colors with standard colorFilter/DXF references and complete desktop selections in an identified filter-column extension; import both paths in the real native sidecar.
- [x] Correct solid DXF fill foreground/background selection, preserving existing conditional-format/background and table-style regressions.
- [x] Verify unit/CLI/real-sidecar save/reopen and actual browser fill/font/multiple/no-fill/retry/Reapply workflows, plus native VS Code history, source preservation, single-color import, multicolor import/edit and clearing after independent reopens.
- [ ] Verify external Excel interoperability for multicolor extensions, pending-formula panel opening, rich text/conditional/table/gradient/theme variants, malformed imports, all appearance/locales and installed Linux/Windows VSIX behavior before claiming complete color-filter parity.

Data Tools control reuse and Text to Columns follow-up (2026-10-29):

- [x] Extract the actual desktop DataToolsRibbonGroup and embed it in the plugin with its original icons, descriptions and four-option delimiter dropdown; preserve existing dialog adapters.
- [x] Accept sparse blank source cells, retain absolute row/column offsets, and grow used dimensions through canonical writes without inserting columns or shifting unrelated references.
- [x] Track Text to Columns reads and commit as single-flight work included in Save; verify duplicate rejection, failed atomic commit/retry and all four delimiter menu choices in the browser.
- [x] Verify native VS Code shared-menu selection, exact Undo/Redo, unchanged formula references and custom package content, and independent clean-session XLSX reopen.
- [ ] Complete large/rich/formula/protection/merged-output/locale/appearance and installed Linux/Windows acceptance; retain the existing 2,000-cell budget until complete-source work is implemented.

Forecast and Outline control reuse follow-up (2026-10-29):

- [x] Extract the actual desktop Forecast group and Outline menus; mount What-If Analysis → Goal Seek and the original Group/Ungroup command lists in VS Code, retaining labels/icons/descriptions and summary-direction check marks.
- [x] Remove the superseded plugin Goal Seek button, flat outline buttons and summary checkboxes; route the shared menu commands to the existing canonical solver/outline adapters.
- [x] Retain in-flight semantic-work results across cell-editor flushing so Save returns false when an already-queued recoverable transaction is rejected; verify rejection/retry and unchanged summary marks.
- [x] Fit shared MenuSelect popovers to the actual editor viewport with scrolling and position updates, so a small VS Code webview's status bar cannot cover the final menu entries.
- [x] Verify real VS Code shared menu clicks at the original small viewport, nested row/column outlines and summary settings, exact history, Goal Seek, saved XML and independent reopen including restored menu check marks.
- [ ] Complete Subtotal workflow and its actual desktop dialog, broader menu keyboard/navigation/lifecycle/locale/appearance variants and installed Linux/Windows acceptance before accepting the full Outline/Data ribbon.

Subtotal workflow follow-up (2026-10-29):

- [x] Extract shared adjacent-group Subtotal planning for desktop and canonical candidates, preserving Sum/Count/Average (SUBTOTAL 9/3/1), absolute-column references, group labels and grand-total formulas.
- [x] Reuse the actual desktop Subtotal button and dialog in VS Code, with frozen host-read headers, group/value/function selections, localized errors, pending dismissal protection, retained rejected drafts and retry.
- [x] Add a standalone canonical subtotal operation, complete host structural materialization and private insertion/content/outline stages that publish one final root/history entry; retain the general mixed-transaction guard.
- [x] Regenerate native tool schemas and reject the new session-only operation explicitly in legacy desktop AI/standalone CLI/in-memory execution, rather than silently treating it as a different layout edit.
- [x] Verify adjacent/blank/repeated groups, non-A and nonzero-row formulas, three functions, structural/content phase failures, cancellation, permissions/incomplete source guards and browser pending/Save/retry/stale/cancel behavior.
- [x] Verify real VS Code shared-dialog submission as one revision, exact Undo/Redo, shifted cross-sheet references, original styles/custom XML, saved two-level outline and independent reopen with visible gutter and summary-direction menu mark.
- [ ] Finish existing-outline/visibility/filter/merge/name/visual/reference transforms, large-workbook resource coverage, computed-value/cache acceptance, all locale/appearance and installed Linux/Windows behavior before accepting full Subtotal parity.

Consolidate desktop-dialog reuse follow-up (2026-10-29):

- [x] Replace the plugin Consolidate replica with the actual desktop dialog, preserving its Dropdown controls, source list, descriptions, label-mode behavior and all five functions.
- [x] Add controlled locale, async apply, synchronous duplicate-submit lock, pending input/Cancel/Escape protection and failure retry to the shared component; retain the desktop synchronous callback contract.
- [x] Treat sparse omitted labels as blanks; refuse pending/error label formulas or a ready formula with no current result before canonical writes.
- [x] Browser-verify function formulas, reference-list rules, label-mode Max/Min reset, invalid input, sparse labels, pending source/commit Save, rejection/retry, stale/cancel and pending-formula refusal; native-verify actual desktop dropdown choices, source refs, exact history, original package preservation and independent saved-formula reopen.
- [ ] Complete larger ranges, formula/cache/error variants, all locales/appearances/geometry/lifecycle and installed Linux/Windows acceptance before accepting full Consolidate parity.

Get Data shared group follow-up (2026-10-29):

- [x] Extract the desktop Get Data group as shared `GetDataRibbonGroup` and mount it in both hosts.
- [x] Share desktop CSV picking, byte/cell budgets, charset preference and typed matrix building (`csv-import-model.ts`); parser cap failures map to the desktop "too large" message.
- [x] Import CSV in the plugin as one canonical transaction bound to the post-pick selection/revision, with merge, grid-edge, stale and single-flight guards included in Save flushing.
- [x] Verify native shared-button import, cancel, Refresh All no-pivot message, one revision, exact Undo/Redo, saved package preservation and independent reopen.
- [x] Lift the 2,000-cell CSV limit to the desktop 50,000 cells through a host staged transaction that stays one history entry.
- [x] Implement Merge Workbooks: host picker, private source snapshots, deduped sheets with values/styles/merges/row heights/column widths as one staged commit.
- [ ] Merge Workbooks beyond 50,000 total cells, hyperlinks/validation/conditional formats/visuals of source sheets, and installed Linux/Windows acceptance.
- [x] Pivot-aware Refresh All.
- [ ] Investigate intermittent `protected-view-save.test.ts` failure under full parallel workbook-session runs (1 of 4 runs; passes alone).

Working-tree fixes found while resuming (2026-10-29):

- [x] `projectDataValidationRules` readonly clone type error broke typecheck in four packages.
- [x] Data-validation sqref rewriting kept throwing on document-controlled areas outside the A1 grammar; such areas are now preserved verbatim (restores `dv-sqref-replace` regressions) while parseable areas keep overlap splitting and formula offsets.

Native suite baseline findings (2026-10-29, VS Code 1.109.5, full `--label minimum` run):

- [ ] `Product workspace skill commands` (expects packaged skill 0.1.0, finds 0.2.0) and `Product worksheet dimensions` (91 vs 104 px) fail in isolation at checkpoint 29ceb04c as well; pre-existing, not caused by the Get Data work.
- [ ] Column Header Filter, Advanced Filter and the bulk CSV test pass alone but failed in the full run (notification toasts/chat pane covering the webview, Go To left disabled, timeouts); harden suite isolation.

PivotTable create follow-up (2026-10-29):

- [x] Add session `add_pivot` with desktop default placement, shared layout and source labels, overlap/name checks, bounded private write stages and one history entry.
- [x] Persist session pivots as native pivot parts through a final package-writer stage; refuse coordinate/sheet changes that would desync a pending pivot.
- [x] Mount the shared Data PivotTable group and Insert PivotTable button with the actual desktop dialog (async, retained draft on rejection); verify native create/history/save/reopen.
- [x] Refresh / Refresh All recompute for file pivots through the shared pivot engine, with growth, refreshOnLoad and location persistence; session-created pivots refresh after save, as on the desktop.
- [x] Edit PivotTable with the actual desktop PivotTable Fields pane: relayout of session and file pivots, cursor-following pane, deferred Update, history and saved regenerated definitions.
- [x] Slicers, timelines and PivotChart.
- [ ] `Desktop Insert Link UI` is flaky (Go To stays disabled in some runs, also at HEAD fc726c60); passed in the later Data/Insert run.
- [ ] CLI: `mcp.test.ts` schema budget (96,298 > 94,208 bytes at HEAD), sidecar `info`, missing generated fixture and Homebrew install tests fail at HEAD; pre-existing.
- [ ] A host-side commit that lands while a reopened view performs its first reads leaves the view on "The requested state is no longer current"; the view should resync instead of showing the error screen.
- [ ] Selection-in-pivot hints and edit protection should follow grown (refreshed) output ranges, not only the original pivot ranges.
- [x] Go To flake: engine selection-sync focus moved typing from open dialogs/panes into the hidden cell editor. The focus boundary now holds that focus back while a dialog/pane control (including `aside` panes) owns focus, and replays it only when a cell edit actually starts (set-cell-edit-visible), so non-modal reference editing keeps working. Browser suite 132/132.

View tab follow-up (2026-10-29):

- [x] Mount the actual desktop View band: Normal/Page Break Preview, Gridlines/Formula Bar/Headings/Right-to-Left/cross highlight, Zoom menu/100%/Zoom to Selection and Freeze Panes menu.
- [x] Persist zoom, right-to-left, gridlines/headings and freeze through canonical set_page_setup/set_freeze with history; project right-to-left; formula bar, cross highlight and page-break preview are display-only like the desktop (shared overlays).
- [x] Verify browser and native View band commands, exact Undo/Redo, saved sheetView XML and independent reopen.

Review protection foundation (2026-10-29):

- [x] Session `protect_sheet`: Excel-hash passwords with a transaction-derived salt (view/host agree), verified unprotect, edits refused while protected, protection changes allowed on protected sheets, host-only candidate for the 100,000-round hash, saved/removed sheetProtection element.
- [x] Review tab R1: actual desktop band; Workbook Statistics over the canonical root; Translate routes to native chat; actual Protect/Unprotect Sheet dialog with verified passwords; Protect Workbook (`protect_workbook`) and Allow Edit Ranges (`set_protected_ranges`) with history, save and reopen; protection echo from host features.
- [x] Review tab R2 notes: Univer note preset in the plugin; projected file/session notes in the desktop "Author:" form; popup edits commit as `set_note` when the popup closes or before Save/flushing commands; New/Show-Hide/Show All/Delete/Previous/Next.
- [x] Review tab R2 comments: canonical `set_thread`; shared thread store and desktop Comments pane, marker and hover card; post/reply/resolve/delete/edit as canonical edits; Convert to Notes keeps the author; threaded parts saved and reopened.
- [ ] Note pictures pane (note-pictures-pane-toggle) and VS Code author identity for new comments (currently the desktop fallback author).
- [x] Insert > Checkbox: canonical checkbox validation on the selection.
- [x] Insert > Icons and Equation: actual desktop dialogs; rendered PNGs become canonical pictures (data URLs decoded directly under the webview CSP).
- [x] `set_data_validation` was never in the session operation whitelist, so plugin validation edits (editor and Checkbox) were rejected by the session; added and verified natively.
- [ ] Insert > Screenshot: no VS Code screen-capture API; reports an explicit host limitation.
- [x] Tables: canonical `add_table` and new `edit_table` (rename, style, style options, header/totals rows with the desktop's cell writes, resize, convert to range) for session and file tables; Univer table preset projects file and session tables with Excel palettes; Insert Table / Format as Table / Create & Resize Table dialogs / Table Design contextual tab; saved table parts and reopen verified natively.
- [x] Insert > Sparklines: canonical `add_sparkline` in sessions (cells derived from the current sheet name at projection/save), shared desktop renderer, native menu/history/save/reopen.
- [x] Insert > PivotChart: one transaction (add_chart + series edit of the predicted chart id) with the desktop's pivot series; native verified on a reopened file pivot.
- [x] Insert > Slicer / Timeline: canonical `filter_pivot` (the desktop's applyPivotSlicer + growth rewrite; shrunken output cleared), actual desktop field pickers and floating panels as session UI; native filter/undo/save/remove verified.
- [x] Desktop window shortcut table (global-shortcuts.ts) in the plugin: Ctrl+T Create Table, Ctrl+1, Ctrl+K, Shift+F2, Ctrl+Shift+number formats/borders, Ctrl+; date/time, Shift+F11 Insert Sheet, PageUp/PageDown paging, Alt+Shift+arrows outline, Ctrl+Shift+U formula bar; captured so VS Code's own chords do not fire.
- [x] Home gaps found by a dynamic per-control audit: Format Painter (once/locked/Esc, formats painted on release through paste_cells), Insert/Delete Cells (whole rows/columns direct, shift dialog via relocate_range), Fill → Series (desktop planner, canonical writes), Sort & Filter / Flash Fill routed to the shared Data handlers.
- [x] Conditional Formatting: Univer CF preset in the plugin; file rules projected with the desktop builder (cf-rule-build.ts); rule-manager edits captured as declarative `set_conditional_formats` (save-ability checked at edit time, coordinates pinned while pending); saved as native rules and read back after reopen.
- [x] `Workbook native shortcuts UI` failure was a notification toast covering the Go To button; the test now clears notifications.
- [x] Univer keyboard layer from the desktop (excel-shortcuts.ts, excel-jump-nav.ts, editor-word-jump.ts): tabs, Ctrl+Home/End, Home, Ctrl/Shift+Space, data-edge jumps; Ctrl+9/0 visibility and Ctrl+Shift+= / Ctrl+- route to session edits and the shared Insert/Delete Cells dialog.
- [x] Desktop cell / header context menu (cell-context-menu.ts) enabled in the plugin; Univer built-ins (clear, insert/delete rows/columns/cells, freeze, hide/unhide, AutoFit) replaced by the ribbon's canonical paths; dialog entries reach the shared ribbon dispatcher.
- [ ] Context menu: Filter by selected cell value/color (needs smart-toggle-filter in the view), Text to Number, table row/column entries, range protection entries and clipboard entries use Univer defaults; verify natively.
- [ ] File CF rules on sheets with pending structural edits are not projected (source coordinates); cfvo formula/reference thresholds are not pre-resolved as on the desktop.
- [ ] Allow-edit ranges do not yet permit edits inside a protected sheet (desktop guard allows them).
- [ ] Browser suite under full load: `function confirmation preserves an unfinished grid draft` and `Sort & Filter group toggles smart source regions` failed once each after the table preset was added (130/132); both pass alone. Watch for a focus-timing regression.
