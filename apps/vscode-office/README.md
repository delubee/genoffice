# GenOffice Workbook for VS Code

Preview extension for manual and agent editing of `.xlsx` workbooks. Documents use a custom editor; chat, model choice, skills and terminal execution use VS Code's native agent features. VS Code 1.109.5 or later is required. No model credentials are managed by this extension.

## Install and open

Install the VSIX matching the **extension host**: Linux x64, Windows x64 or macOS arm64. In Remote SSH, install it on the remote host. Use **Extensions → Install from VSIX**, trust the workspace, and open an XLSX with **Open With → GenOffice Workbook** or **GenOffice: Open Workbook**. Unsupported host architectures, virtual workspaces and browser-only VS Code are rejected.

Linux x64 is locally verified. Windows 11 x64, macOS 14 arm64 and actual Remote SSH acceptance remain unverified. See [implementation status](IMPLEMENTATION-STATUS.md) in the source repository and `test-results/` for exact evidence. This preview is not a release candidate.

## Manual and AI edits

The editor supports values, formulas, basic formats, typed copy/fill, worksheet/row/column operations and session charts. All edits share one versioned document and native undo/redo. Save and Save As write XLSX; applying an AI edit does not implicitly save it. Original native charts and advanced/protected structures are preserved or rejected when unsupported.

1. Select a range and run **GenOffice: Bind Workbook Selection**.
2. Paste the copied binding reference into native Chat and enable GenOffice tools in the agent tool picker. An agent that supports registered tools is required.
3. Read current context/capabilities before editing. After a structure change, rebind the selection. A stale revision requires a fresh read and reconsideration.
4. Use **GenOffice: Workbook Skill Status** to inspect local registration; it cannot verify tool enablement for a particular conversation.

Eight tools: `genoffice_workbook_context`, `_read`, `_aggregate`, `_apply`, `_operationStatus`, `_export`, `_prepareImport`, `_applyImport` (each suffix uses the `genoffice_workbook` prefix). Explicit URI or binding determines the target; changing editor focus never retargets a running task.

## Optional skill and Python

Run **GenOffice: Install or Update Workbook Skill** and choose a workspace. It installs `.github/skills/genoffice-workbook/`, with visible references and `scripts/analyze.py`. Existing files are reviewed; updates require explicit choice and preserve a backup. Enable Agent Skills through the offered VS Code settings link if needed. Tools and manual editing work without the skill or Python.

For complex analysis, export current unsaved typed data into `.genoffice/work/<task>/`. Review/copy the provided Python script there, then let the native agent terminal run the visible command, subject to its usual permissions:

```sh
python3 analysis.py snapshot.manifest.json
```

Use the **actual manifest path returned by export**; filenames above are examples. Python must already exist on the extension host; dependency installation is never automatic. The example uses only Python's standard library. It emits a result manifest; prepareImport validates and privately stages it, then applyImport commits each next batch. Keep the original binding and revision chain. Do not overwrite an open XLSX from Python. Review the proposed target and independently verify business totals after import.

Cancellation preserves already committed batches and receipts. Undo reverts each batch. Other edits between batches invalidate the plan; export and recompute. Query operationStatus after an uncertain response instead of submitting a second write. Scripts and public results are not silently deleted.

## Limits and measured capacity

| Operation | Limit |
| --- | --- |
| Current read | 2,000 cells; 200 with formats |
| Aggregate | 1,000,000 cells; 50 frequent values |
| Ordinary expanded edit | 2,000 cells; 1,000 operations |
| Supported range operation | 200,000 cells, subject to budgets |
| Typed exchange artifact | 1,000,000 cells AND 64 MiB |
| NDJSON record / manifest | 256 KiB each |
| Import commit | At most 1,000 cells |
| Import plan lifetime | 30 minutes |
| Forward + inverse transaction | 64 MiB |
| Retained unique history storage | 256 MiB; edits reject before eviction |

100,000-cell input analysis and complete 100,000-cell output were actually verified locally. Full output took about 11 minutes for 100 batches in a Node/Rust correctness test; this is not a UI latency claim. Spreadsheet address dimensions and configured limits are not tested capacity claims. Formula values with unverified caches are reported pending. Original dates preserve workbook date system and serials, including serial 60.

## Build from source

From the repository root (Node 22.12+, npm, Rust):

```sh
npm ci
npm run native:build -w apps/vscode-office
npm run vscode:build
npm run vscode:typecheck
npm run vscode:test
npm run notices
npm run package -w apps/vscode-office
```

Each platform packages its own sidecar. VSIX files and SHA-256 sums go to `apps/vscode-office/artifacts/`. No Cargo, Electron app or download is required at extension runtime. No command publishes to Marketplace. See [troubleshooting](docs/troubleshooting.md).
