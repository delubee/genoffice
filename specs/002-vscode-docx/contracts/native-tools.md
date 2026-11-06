# Native DOCX tools and programming contract

Keep existing genoffice_workbook_* tools unchanged. New genoffice_docx_* tools
use explicit bindings and versioned requests; context without target lists open
DOCX documents and never guesses focus.

- **context**: explicit documentUri or bindingId, optional selection binding;
  returns identity, revision, safe selection, supported operations and limits.
- **read**: binding, expectedRevision and block IDs/range; ≤100 blocks and 256 KiB.
  Returns safe text/type/logical IDs and editability, not raw PM JSON/XML/media.
- **apply**: binding, expectedRevision, unique transactionId, summary and ≤100
  supported high-level operations. Replace text, insert/delete paragraphs,
  paragraph/font formatting and basic list operations use host-generated steps.
  Validate all targets first and commit atomically. Protected/hidden content,
  unsupported structure and stale revisions fail without partial effects.
- **operationStatus**: session + transactionId returns the factual receipt;
  retries never invent a new ID to guess whether a mutation happened.
- **export**: a current safe structured snapshot for workspace Python; no direct
  overwrite of an open DOCX. Only native terminal/program execution is used.
- **prepareImport/applyImport**: bounded structured result validation and preview,
  then hash/version-bound transaction through the same history. Changed files,
  stale versions, path escapes and untrusted XML reject.

Reads and writes flush current editor state. A tool call's confirmation describes
actual target and changes; contents are data, not instructions. A tool call does
not implicitly save the file. Real-model and skill-discovery tests are distinct
from API-level tool tests and must be reported independently.
