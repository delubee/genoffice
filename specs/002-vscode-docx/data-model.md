# DOCX data model

- **Session**: schemaVersion=1; sessionId UUID; documentUri explicit file URI;
  immutable original bytes/hash/ParsedDocFull; current content root; monotonic
  revision; stateId; separate collabVersion counting accepted steps; writable,
  dirty and conflict flags. Saving records a saved-root fingerprint, not a new
  original base.
- **Content root**: validated PM JSON; host-generated logical IDs; pending list
  definitions and style upserts; next numbering ID floor. Logical ID differs from
  original docxIndex. Maximum serialized root 32 MiB, history 256 MiB.
- **Source policy**: original anchor metadata, trusted raw attribute values,
  protected object signatures and block writability. Source metadata is not a
  writable API. Advanced/protected blocks survive every accepted transaction.
- **Transaction**: transactionId unique within session; origin manual/agent/import;
  expectedRevision; manual baseCollabVersion and clientId; bounded steps or typed
  operations; receipt with revision/stateId/collabVersion. Same ID+input replays
  receipt; same ID+different input rejects. No partial commit.
- **History**: before/after roots and serialized budget; redo invalidated on new
  edit after undo. Native history restores roots and advances revision; old tool
  bindings cannot regain freshness by coincidental matching text.
- **Selection binding**: bindingId, session/document identity, acknowledged
  revision, logical block IDs and validated PM selection positions. Re-read after
  a version change; new session invalidates old binding. Context is host-derived
  safe text, never a browser-provided quoted string.
- **Presentation DTO**: PM JSON plus style and numbering entries, layout values,
  safe media resources and capability flags; no Node handles or filesystem paths.
- **Exchange**: snapshot identity/version/target, ≤1000 safe blocks and ≤2 MiB,
  result hash and preview plan; trusted workspace-local result paths; stale or
  changed results reject before commit. XML/PM raw attributes are not exchange data.
- **Backup**: versioned private manifest with a serialized current OOXML snapshot,
  its hash and the original disk baseline/fingerprint; canonical storage
  confinement and bounded sizes. Restore creates a new unsaved session identity
  and a fresh native history. The process-local VS Code undo stack is not carried
  across a restart; undo history remains intact across ordinary saves.

Lifecycle: closed → opening → ready → editing/pending → committed → saving →
ready; conflict blocks overwrite and permits explicit Save As/reload. Disposing
invalidates bindings. Backup/restore re-enters opening with a new session ID.
