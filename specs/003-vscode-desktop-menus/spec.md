# Desktop menu parity in VS Code

Status: in progress. This feature is not complete when tabs merely look alike.

## Requested outcome

Reproduce the GenOffice desktop document menus in the existing VS Code extension,
for both XLSX and DOCX. The desktop application at the feature branch base is the
reference for tabs, command groups, icons, order, dropdown contents, dialogs,
keyboard interaction and selection-dependent states. Manual editing and AI
editing must continue to share one authoritative document and undo history.

VS Code native chat remains the AI interface, including skills and Python. An AI
entry in a shared ribbon must route to native chat with its document context.
Desktop AI panels, model providers and credential configuration do not belong in
the webview. File, clipboard, image, printing, window and other privileged actions
need explicit host adapters that perform the requested action.

## Reference surfaces

| Document | Standard tabs                                                                | Contextual tabs                                                           |
| -------- | ---------------------------------------------------------------------------- | ------------------------------------------------------------------------- |
| Sheets   | Home, Insert, Page Layout, Formulas, Data, Review, View                      | Chart Design, Table Design, Shape Format, Picture Format                  |
| Docs     | File, Home, Insert, Draw, Design, Layout, References, Mailings, Review, View | Table Design, Table Layout, Picture Format, Shape Format, Header & Footer |

Reference implementations:

- Sheets: `apps/sheets/src/renderer/ExcelShell.tsx`, the extracted
  `SheetsRibbon.tsx`, `ribbon-actions.ts`, and their dialogs and command helpers.
- Docs: `apps/docs/src/renderer/components/Ribbon.tsx`, its `ribbon-*-tab.tsx`
  components, dialogs, and editor transaction helpers.

Quick access commands, ribbon collapse/density, menus within menus, contextual
menus, selection formatting echoes and associated panes are part of the scope.
Host-specific window handling may use VS Code equivalents, with equivalent user
outcomes; a visible command with an empty callback is incomplete.

## Acceptance requirements

1. Enumerate the controls and command variants from the desktop sources. Track
   their host adapter, persistence path and verification evidence; include
   dynamically generated menu options and all contextual selection states.
2. Reuse desktop components and styles where possible. Keep both existing
   appearance choices; the default Office appearance must match the desktop
   document UI. Theme changes must never alter authored document colors.
3. Every enabled menu action must work with the current selection. Dialogs must
   retain desktop defaults, units, validation, previews, confirmation, cancel and
   focus behavior. Disabled states must follow document/selection constraints,
   rather than conceal an unimplemented adapter.
4. Each mutation must enter the existing host session, participate in shared
   undo/redo and survive save/reopen. Extend session operations, projection and
   serialization together when desktop features exceed the current DSL. Preserve
   original OOXML and unsupported content that was not edited.
5. Cover rich DOCX features currently protected by the first editor slice,
   including merged tables, pictures, shapes, headers/footers, sections, fields,
   styles, numbering, review and references. The current protection policy is an
   implementation starting point, not a reduction of the parity requirement.
6. Cover XLSX formulas, page layout/printing, data operations, validation,
   conditional formatting, tables, charts, images/shapes, pivots, protection and
   review alongside all Home and View operations present in the desktop UI.
7. Translate shared controls using the existing 21 locale dictionaries. Supply
   a controlled locale from the embedded host without requiring an Electron
   preload. New translation keys must exist in every locale.
8. Keep browser dependency guards and strict webview CSP. Explicitly bridge
   privileged capabilities; do not emulate a desktop preload object in the
   webview or import a desktop chat runtime to render a ribbon.
9. Verify actual rendered controls and interactions, dialog results, authoritative
   session changes, shared history and saved documents. Run Linux and Windows
   production VSIX checks for both document types, including virtual-drive
   junction paths. No completion claim based only on typechecks or tab counts.

## Current gap

The plugin still displays its initial simplified ribbons. Extracting a reusable
desktop ribbon makes the real UI available to the embedded host, but does not by
itself implement plugin parity. Production adapters, complete DOCX capabilities,
shared styles, full menu coverage and packaged runtime verification remain open.
