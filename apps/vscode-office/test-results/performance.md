# Performance evidence — 2026-10-22

Command: `npm run benchmark -w apps/vscode-office` (display required on headless Linux).

Actual product VS Code 1.109.5 Webview, deterministic 10,000 x 10 two-sheet XLSX, SHA-256 `ae1fb1a0c89752af8f122ab54d3ca57b9304c912bd93f2be1feb96c2b63d3da1`. Each open uses a distinct document/session. First sample is a cold workbook in an already activated extension host; later samples use the warm host. This does not measure a fresh VS Code process for each open. Dependency preparation is excluded.

SC-001 ends after installed first viewport, successful editable-view flush and two animation frames. SC-005 ends after the canonical commit, complete installed grid values independently compared with all 1,000 expected numbers, and two animation frames. It includes host validation/history/projection work, has no model or external process, and does not subtract any incidental formula calculation. `benchmark-raw.json` retains every sample.

| Local measurement | Samples under threshold | Median ms | Maximum ms |
| --- | --- | --- | --- |
| SC-001 <=10,000 ms | 20/20 | 1278.13 | 1737.67 |
| SC-005 <=2,000 ms | 20/20 | 649.67 | 886.46 |

Both **local** thresholds satisfy 19/20. The machine reports 4 logical CPUs (AMD EPYC 9354P) and 16,524,685,312 bytes of RAM; SSD class has not been independently established. It is **not** the required 4-vCPU / 8-GiB / SSD reference environment. Reference SC-001/005 acceptance remains **unverified**, and T093 remains unchecked. No threshold was changed.

The separate 100k output correctness run took 669,997 ms across 100 batches; that run has no Webview and is not substituted for either UI gate. See `us3-python.md`.
