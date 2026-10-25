# DOCX reuse research

## Existing schema and save bridge

Decision: reuse `editor/extensions.ts` and `editor/convert.ts` through a thin adapter.
`blocksToPmDoc`, `pmDocOptions`, `pmDocToSavePlan` and `docx-engine.saveDocx` cover
rich paragraphs, tables and original OOXML. Replacing this with HTML conversion
would lose source properties. Full Docs App is coupled to window.desktop and AI.

Two read-only research agents traced this graph. A real esbuild Node/CJS smoke
with only three UI barrel imports replaced produced 18 nodes/15 marks, validated
an empty document and made a save plan without DOM. This is evidence for
feasibility, not proof of rich document fidelity; nonempty fixtures gate coding.

## Three pure UI boundaries

`protected-render.ts` needs wordart-presets; `shape-draw.ts` needs shape-gallery;
`font-list.ts` needs font-list. Add explicit ui subpath exports and replace these
imports. Do not suppress CSS loaders or permit the full UI/AI barrels. Audit the
pure agent-core/redact-range subpath separately; retain all other AI import bans.

## Source identity and serialization

`docxIndex` indexes the ORIGINAL document body; `Block.id` is not durable identity.
The desktop converter intentionally supports split twins. New logical IDs must
be distinct from source anchors. New split nodes clear bookmark/comment/SDT and
revision source metadata; they must not blindly inherit original XML references.
Original ParsedDocFull stays unchanged across save/history. Edited save rebuilds
ZIP and may update core properties; compare untouched entry contents.

`Node.check()` does not validate most attribute values. Host validates raw/source
attributes, protection and anchors in addition to schema shape. Unrecognized
protected nodes can otherwise disappear and repeated indexes can clone content.

## Synchronization

Use a host authority plus ProseMirror collab for unconfirmed local steps, not a
second authority in Webview. Capture initial and appended transaction steps.
[Official collab implementation](https://raw.githubusercontent.com/ProseMirror/prosemirror-collab/master/src/collab.ts)
confirms that sendableSteps reports the authority version and receiveTransaction
confirms local steps while rebasing pending ones. Package metadata confirms
version 1.3.1 with prosemirror-state dependency; it is not included in @tiptap/pm.

Disable editor UndoRedo and flush pending/composition before native history.
Do not accept arbitrary model-provided steps; model operations are typed commands.

## Lists, styles and protection

Real lists need SaveOptions.numbering.newDefs `{numId,kind}` and nodes with that
numId. Missing Heading styles need explicit StyleUpsert and matching styleId;
saveDocx does not synthesize missing Heading1 automatically. Pending numbering
and styles belong in history state.

Document protection is enforced in the desktop App, not automatically by parse
or save. Recreate that guard in host. Complex table NodeViews retain DOM drafts;
make them read-only in the first increment. Redaction marks carry real hidden
text: use explicit safe projections for context/read/export and reject AI/import
whole-block changes that contain hidden spans until preservation is implemented.

## Load and theme boundaries

Two separate HTML entries isolate global styles and lazy-load editor code.
A separate lazy host bundle prevents parsing/initializing the Docs graph for
XLSX-only use. Reuse doc-style-css and listNumbering storage through explicit Map
entries DTOs. Word-like layout uses document widths and fonts; exact Word
pagination is outside the initial fidelity promise.
