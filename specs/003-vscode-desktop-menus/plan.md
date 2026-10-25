# Implementation approach

1. Separate browser presentation from desktop application orchestration. Extract
   the exact Sheets command band, retaining the desktop shell as a consumer.
   Expose office chrome independently of the AI UI barrel. Provide controlled
   locale scopes. Keep command state and callbacks explicit.
2. Inventory each standard/contextual tab, dropdown option, dialog and pane in
   Sheets and Docs. Map controls to desktop command helpers and available session
   operations. Treat missing operations as work to implement, not as reasons to
   remove controls from the goal.
3. Adapt the shared Sheets command band to `WorkbookRuntime`. Extend validated
   operations and projection where necessary, and reuse dialog models. All
   mutation paths must acknowledge host acceptance before being considered saved.
4. Extract privileged ports from the Docs ribbon and its transitive helpers.
   Reuse the same ribbon and dialogs with the DOCX runtime. Extend the policy,
   conversion, session validation, native operations and serializer together for
   rich features; remove initial slice restrictions only with save/history proof.
5. Route native chat, file actions, clipboard, image pickers, export/print and
   window actions through typed host ports. Preserve URI-based resource handling
   and the Windows junction fallback. Keep desktop behavior intact.
6. Share the appropriate desktop chrome styles with both embedded appearances.
   Verify compact/collapsed layouts, narrow panes, contextual transitions,
   popover layering, focus, shortcuts and locale changes in actual browsers.
7. Verify representative actions in each command family end to end; keep a
   complete coverage ledger for all controls. Build and run installed VSIX
   packages on Linux and Windows before final acceptance.

## First implementation evidence

- `SheetsRibbon.tsx` now contains the original desktop command band; the desktop
  shell consumes it. No second menu definition was introduced for VS Code.
- `@genoffice/ui/chrome` supplies browser-only office controls without importing
  the AI UI barrel.
- `paste-special-model.ts` isolates options and availability from Univer's
  clipboard implementation; both the dialog and command band consume the model.
- Docs and Sheets expose `LocaleScope` for controlled embedded rendering.
- The Docs ribbon, font/paragraph dialogs, color palette, insertion/reference
  tabs and wheel zoom now import the browser-only chrome entry. Both desktop
  ribbon dependency graphs pass the embedded browser boundary.
- Embedded component tests exercise all 11 Sheets tab bodies, formatting/menu
  callback parameters, view-state echo, language changes and Escape dismissal.
  A separate browser bundle test checks the transitive dependency boundary.

These are preparation for host integration. They do not prove mutations or save
behavior from the new shared ribbon inside the production extension.

## DOCX dialog integration

The embedded DOCX editor now opens the same Font and Paragraph components used
by the desktop ribbon. It imports the desktop Docs stylesheet and shared dropdown
styles, with native host theme tokens mapped for the VS Code appearance. The
document page retains its authored colors. Dialog labels use the controlled
desktop locale scope.

The document policy explicitly validates the added font effects, advanced
character settings, paragraph spacing/indents and pagination flags. Existing
source mark values remain preservable when outside editable dialog bounds;
newly edited values must pass validation. Original mark values are indexed once
per source block during validation, avoiding repeated subtree scans.

Authoritative resynchronization now restores the ProseMirror selection and emits
a transaction notification so selection-dependent controls update before they
receive focus. This fixes font input remounting after history/resync.

These dialogs are currently launched from the existing simplified toolbar. The
complete desktop ribbon, other dialogs and all XLSX adapters remain open work.

## Native file and chat menu ports

Both embedded editors now expose a shared five-entry file/chat action group.
The browser sends only the typed intent (`open`, `save`, `saveAs`, or `chat` with
document/selection scope). A strict discriminated schema and bridge method
allowlist reject extra URIs, command identifiers and unsupported payloads.

The extension captures the resource/session and panel when registering the
handler. It synchronizes pending edits, checks the captured session and active
panel again, and prevents overlapping menu actions. Open uses the native picker
and dispatches supported files to the appropriate GenOffice custom editor.
Save/Save As use VS Code's native commands and existing custom-document lifecycle.

Chat actions capture a native tool binding before moving focus. Workbook
bindings reuse an extracted selection-capture helper without touching the
clipboard. DOCX bindings use the existing context implementation. The prompt
contains resource/binding/revision metadata, not document contents. Generic chat
opens a partial request in native Agent mode; nonempty ribbon presets can submit
their explicit instruction through the same port. Model configuration, skills,
terminal/Python capabilities and tool consent stay with VS Code.

These entry points prepare the callbacks for the full shared ribbons. Other
desktop file/media/print/window capabilities and every preset's actual ribbon
connection still need implementation and verification.
