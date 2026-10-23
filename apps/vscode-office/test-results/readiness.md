# Readiness audit — 2026-10-22

**Preview only; release acceptance incomplete.** Implemented code, local deterministic acceptance and unavailable-environment gates are distinguished below. No release or full-spec completion is claimed.

## Functional requirements

| Requirement | Implementation / local evidence | Remaining scope |
| --- | --- | --- |
| FR-001 | Shared custom editor; `us1-manual.md`, browser/native tests | Target OS repetition |
| FR-002 | Save/Save As/dirty/failure; native and save-fault tests | Target OS UI matrix |
| FR-003 | Manual path independent of model/Python; failure tests | Broader offline manual acceptance |
| FR-004 | Eight native LM tools; no custom chat/model accounts | Real model conversation |
| FR-005 | Explicit context and frozen selection binding | Real agent consumption |
| FR-006 | Current overlay reads and flush/busy handling | Verified locally |
| FR-007 | Explicit session binding; focus/same-name tests | Verified locally |
| FR-008 | Ordinary DSL and chart/structure UI; native and browser tests | Native objects limited by capabilities |
| FR-009 | Read-only query path; actual receipts | Verified locally |
| FR-010 | Packaged skill/install/update/status; `us6-skills.md` | Real skill discovery/conversation |
| FR-011 | Immutable current typed snapshot; 100k native proof | Verified locally |
| FR-012 | Visible stdlib Python and workspace results | Native terminal/model approval unverified |
| FR-013 | Private prepared import; new Summary and manual continuation | Verified locally |
| FR-014 | Registry/path/hash/type/shape/revision/TTL checks | Verified locally |
| FR-015 | Typed roundtrips, both date systems, serial60 and formulas | Pending caches reported explicitly |
| FR-016 | All imports use canonical transactions/history | Verified locally |
| FR-017 | Revision/structure staleness and chain guards | Verified locally |
| FR-018 | Disk conflict/alias ownership, Save As/revert paths | Target OS/SSH repetition |
| FR-019 | Close/reopen/revert epoch invalidation | Verified locally |
| FR-020 | Atomic candidates, failure/cancellation tests | Verified locally |
| FR-021 | One undo per commit, retained history budgets | Verified locally |
| FR-022 | Save retains timeline; undo restores dirty state | Verified locally |
| FR-023 | Full recovery copies, hot-exit/SIGKILL and explicit recover command | Human close matrix / target OS repetition |
| FR-024 | Python failure/cancel and partial import receipts | Real terminal workflow unverified |
| FR-025 | Shared guards for read-only/protected/unsupported writes | Advanced object editing intentionally limited |
| FR-026 | Staged base save; protected/chart/unknown-part preservation tests | Broader external workbook corpus |
| FR-027 | Explicit limits, reject-before-change budgets, separate capacity proof | No untested maximum-capacity claim |
| FR-028 | Native schema discovery, errors and idempotent receipt retries | Real agent discovery |

## Success criteria

| Criterion | Status |
| --- | --- |
| SC-001 | Local 20/20 <=10 s; specified 8-GiB/SSD reference environment unverified |
| SC-002 | Preset local save/reopen/value/formula/format tests passed |
| SC-003 | Local current-state/target/focus/stale-write tests passed; real model unverified |
| SC-004 | Actual 100k native Python summary/chart/save/reopen with independent totals passed |
| SC-005 | Local 20/20 <=2 s; reference environment unverified |
| SC-006 | Preset atomic failure, artifact corruption, program failure and cancellation tests passed |
| SC-007 | Local mixed native history, save and actual recovery proofs passed; OS/UI matrix pending |
| SC-008 | Real logged-in Chat read/edit/Python/skill conversations unverified |

## Implementation gates

| Gate | Status |
| --- | --- |
| G0 | Local stable API / isolated candidate / serialization proof passed |
| G1 | Shared Node/browser boundary, gateway tests and actual Electron builds passed |
| G2 | Local document/history/save/conflict/recovery evidence passed |
| G3 | Product tool/skill code and native invocation/install passed; real Chat acceptance pending |
| G4 | Local Python typed exchange, 100k input and 100k complete output passed |
| G5 | Linux VSIX and minimum/current stable local checks passed; Windows/macOS/SSH, real Chat and reference performance pending |

Unverified platform, remote, model and performance-reference tasks remain unchecked in tasks.md. This audit does not convert unavailable evidence into a pass. No `.specify/extensions.yml` is present, so no before/after implementation hooks apply.
