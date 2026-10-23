# Foundation / G1 — 2026-10-21

## Implemented boundary

The Node session is the canonical authority. A host materializer reads the original source plus the current overlay and verifies index/protection metadata. Stable state/sheet IDs for a proposed transaction are allocated by the host. The browser executes the proposal in a disposable same-version Univer projection and returns serializable complete roots. Before persistence, the host independently executes the shared pure-data semantics against the same material and rejects a different result. Node code imports no Univer/React/Electron/VS Code runtime.

Candidate changesets retain full prior/next typed overlays, ordered structural logs and feature maps. Original package bytes remain independently retained. Reusable gateway operations run one at a time to preserve batch order; failure after any step only affects candidate data. Formula changes are marked pending; no recalculation success is fabricated. Supported candidate operations currently cover scalar/range values, formula, clear, basic formats, sheet add/delete/rename, restricted row/column changes, and session charts. Fill/copy and native chart feature editing are not yet enabled and belong to US1 tasks. Protected or unhandled structural feature maps fail before commit; there is no silent feature drop.

Actual UTF-8 sizes of both directions are checked against 64 MiB and retained history against 256 MiB. Chunk files and history manifests are written before the canonical commit manifest. A single atomic `current` manifest publishes state and its committed receipt together. Request IDs are deduplicated before revision checks. Post-commit view failure does not undo the committed fact or cause duplicate edit events on retry.

The shared JSON bridge validates session/version/method/fields, 1 MiB control messages, 128 KiB raw chunks (below the 256 KiB limit), bounded 64 MiB transfers and SHA-256 reassembly. No handler receives an incomplete or corrupt transfer. Only typed data handlers can be registered; source reads validate logical sheet/state/range and accept no paths or executable code. Closing the bridge rejects pending requests and disposes handlers. The Webview entry installs the candidate handler only after the provider supplies its session metadata; no document UI is claimed at this stage.

## Actual evidence

Environment: this Linux x64 server, Node 22.21.0, Rust 1.96.0, Univer 0.25.1, Chromium 151.0.7922.34, VS Code 1.109.5.

- Session Vitest suite: 16 passed. Includes real sidecar protocol extraction, snapshot isolation/retention, current reads, structural coordinate mapping, private-store interruption, real-byte history budgets, serial commits, lost-response idempotency, stale revision, precise pre-root publication failure, cancellation before/after commit, and projection failure with durable committed receipt.
- Browser-package unit tests: 5 passed. Serializable history refuses closures/accessors/runtime instances; bridge tests include a >1 MiB Unicode round trip, foreign session/extra fields, corrupt digest before handler execution, and disposal.
- Playwright: 4 passed. Actual Univer candidate failure/disposal, structure/chart inverse roots, cross-sheet reference shift+rename, literal error text and protection; plus a real Node session/Rust sidecar/Chromium candidate round trip with independent host replay. Projection failure leaves one committed edit; resync rebuilds from canonical state. A later batch fails at its second operation and leaves the first committed root intact. The Node/browser link in that last test substitutes for the VS Code transport; it is not a claim that the product CustomEditor UI exists.
- Separate stable VS Code host proof: 2 passed (see `g0-design-proof.md`): contributed native LM tool, preparation/invoke boundary, save/undo/redo callbacks in a real 1.109.5 host.
- Existing sidecar/streaming-save/structure regression: 105 passed using the compatibility Electron imports.
- Targeted Rust error-versus-literal-text regression: 1 passed.
- Three new workspaces: TypeScript and production builds passed; browser dependency boundary plugin passed the bundled graph. Bundle size warning remains for later performance work.

## Gate and limitations

G0/G1 permit US1 implementation. No user story is complete. Product save/history UI, backup recovery, chat bindings/tools, Python exchange, skill installation, platform VSIX and performance acceptance remain unbuilt. Windows/macOS/SSH and real model chat have not been run. These results are deterministic foundation evidence, not capacity or publication approval.
