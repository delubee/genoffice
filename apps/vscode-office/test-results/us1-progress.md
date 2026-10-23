# US1 progress — 2026-10-22

The initial progress notes below are historical. Basic US1 verification is now recorded in [us1-manual.md](us1-manual.md), including remaining product and platform gates.

On this Linux x64 host, the actual product CustomEditor has passed opening its packaged Webview, editing via the browser candidate with independent host replay, native dirty changes and successful save. The standalone writer test reopens typed identifiers/empty strings/numbers/formats and preserves an unknown OOXML entry. Browser tests now exercise the real visible Univer runtime: intercepted edits leave the old displayed value until the host commit; committed values are then projected; IME composition blocks flush.

This testing exposed two issues that were fixed: the initial 40-column page exceeded a small 20-column projection, and a fresh Univer unit could precede its scroll controller. Reads are now clamped to the actual projection and an initial missing scroll controller uses its known origin. Initial Webview readiness is resolved only after initial range installation succeeds. Projection listeners are disposed and pending reads are invalidated by an install epoch, including resync to the same content state.

Another preservation correction keeps hydrated original cells out of the committed overlay unless actually changed. `candidate-overlay.test.ts` verifies this independently of the browser. The persistent sheet map now retains removed original IDs and allocated new-sheet identities so staged saves can replay the complete structural log from the original package on every save.

The follow-up implementation added formatting/copy/fill, session chart display/editing, save fault and cancellation coverage, damaged-file rejection and theme tests. Viewports now use scroll events rather than requesting an unregistered scroll manager; source reads are clipped to immutable original dimensions. Product undo/redo is not implemented yet; the stable API undo test uses the separate API-probe fixture. Full V1 history acceptance remains in US5.
