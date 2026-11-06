# DOCX document and projection contract

All messages carry schemaVersion=1 and session identity; state-sensitive requests
carry revision/stateId. Only provider-validated document paths reach the session.

- `docx.init`: canonical content root, descriptor and presentation DTO.
- `docx.submitSteps`: transactionId, expectedRevision, baseCollabVersion,
  clientId and ≤1000 JSON steps / ≤1 MiB. Host applies with the same schema,
  validates attributes/protection, normalizes fresh IDs and returns accepted
  step stream with client IDs. Only manual Webview channel accepts steps.
- `docx.committed`: from/to collaboration versions, accepted steps/client IDs,
  descriptor and updated presentation overlays. Duplicate events do not reapply.
- `docx.flush`: waits for IME and all unconfirmed/host-pending work; false/busy
  is explicit. Needed before read, tools, save, backup and history.
- `docx.selection`: flushed revision/stateId and PM positions, validated by host.
- `docx.history`: freeze after flush, restore canonical root through native history,
  reset collaboration baseline, unfreeze in finally. No second local undo stack.
- `docx.resync`: same session resynchronization with current root/version; stale
  session messages reject. Invalid manual edits report rejection and recover the
  canonical projection without claiming their application.

Persistence always derives from the immutable original base. Reuse baseline
hash/identity checks and staged publication, including the Windows document-path
compatibility helper. Exchange/recovery confinement still uses strict realpath.
