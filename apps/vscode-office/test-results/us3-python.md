# US3 — Python exchange evidence (2026-10-22)

## Native VS Code path

`python-native-proof.json` records the actual VS Code 1.109.5 product tools and Univer view on Linux x64. A deterministic 100,000-cell, two-sheet workbook starts with Data!D2=100 (North); an unsaved edit changes it to 120. Export sees 120 without saving the XLSX. The standard-library Python script produces a Summary whose North total increases by exactly 20; independent expected totals for other regions are unchanged. Prepare changes no content. Import, chart creation, further manual editing, undo/redo, save and reopen pass. Exporting the reopened input produces exactly the same typed data hash as the exported unsaved input; an unknown package part is preserved.

An earlier isolated run took 85,087 ms. Final full-suite runs took 179,186 ms on VS Code 1.109.5 and 66,841 ms on 1.143.0; versioned `python-native-minimum-proof.json` / `python-native-stable-proof.json` retain them. These durations cover the entire scenario, including export, Python, import, UI interactions and save/reopen. This is not an SC-001/005 latency measurement. Python was launched by the test harness, not by a real model's terminal tool. No model approval or conversation is claimed.

## Complete 100,000-cell output

Command: `npm run test:capacity -w apps/vscode-office`.

`python-full-output-proof.json` records a separate actual Node session + Rust source/save/reopen run: 100,000 output cells, 100 batches of at most 1,000, all receipts recorded, 101 retained history entries, 85,856,435 retained serialized history bytes. Public result data was altered after prepare; the immutable private copy remained correct. Current and saved/reopened typed data hashes match the input. After three committed batches, cancellation leaves exactly 3,000 cells committed; three undos restore the prior manual state. A fresh export/recompute is used before completing the later plan.

Duration: 669,997 ms (about 11 minutes). This establishes capacity and correctness, not interactive throughput. Earlier runs exposed quadratic snapshot work and a 600-second test timeout; the candidate builder and owned in-memory adapter were corrected. Product cell, artifact and history limits were not increased. The final capacity test has a 20-minute harness timeout.

## Types and preservation

`tests/python-types.test.ts` passes for 1900 and 1904 workbooks, serial 60 and fractions, percent, boolean, 00123, blank versus empty text, literal equals text, errors and formulas. Imported formula caches remain pending rather than being presented as authoritative. Editing another sheet preserves original protected/merged/chart/unknown parts byte for byte; protected writes fail closed.
