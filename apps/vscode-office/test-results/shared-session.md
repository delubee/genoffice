# Shared session extraction — 2026-10-21

## Implemented

- T012: moved the pure Node sidecar client to `@genoffice/workbook-session/sidecar/client`; the Electron import remains a compatibility re-export. Request IDs, session-scoped cancellation, protocol version guards, timeouts and crash notifications are retained.
- T013: moved workbook-only Zod protocols to `@genoffice/xlsx-gateway/shared/workbook-protocol` and shared transport limits to `workbook-limits`. Desktop screenshot, AI configuration and print/host interfaces remain in the Electron module. Existing public workbook names are re-exported.
- T014/T015: copied immutable original package, streaming SHA-256, realpath/device/inode identity, reference-counted source cleanup, stable logical sheet IDs and base-relative typed overlays. An opened source cannot change if its original path is modified later.
- T016: bounded immutable read views combine original reads, structural coordinate mapping and typed overlay. They distinguish blank, empty string, literal identifiers, numeric dates/date systems, formula cache/pending state and display values. The source refuses incomplete index coverage instead of presenting missing cells as blanks. Unsupported source styles fail validation. No source style rewrite occurs on disk.
- T017: content-addressed 256 KiB chunks, digest checks, 256 KiB manifests with complete references, atomic rename after file fsync and reference-aware garbage collection. Fault injection before root rename preserves the prior root.

## Actual validation

Linux x64 / Node 22.21.0 only. The existing Rust sidecar release build succeeded. New session suite: 11 tests passed, including real sidecar open/read via the extracted client, original-file mutation after snapshot, retention, typed reads and interrupted manifest publication. Existing Electron-import sidecar/streaming-save/structure suites: 105 tests passed. Gateway and Sheets TypeScript checks passed. Three new workspaces built successfully.

The Rust reader needed an additional optional `isError` field to preserve error-versus-literal-string identity; older normal-cell payloads remain unchanged. Its targeted regression is recorded separately after execution. Candidate transactions, document UI, saving, tools, Python exchange and recovery are still incomplete. These infrastructure results do not establish any user-story completion or large-workbook capacity.
