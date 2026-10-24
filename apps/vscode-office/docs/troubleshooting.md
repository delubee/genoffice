# Troubleshooting

- **No editable workbook:** trust the workspace, choose the custom editor explicitly and install the platform VSIX on the extension host. Web/virtual resources and unsupported architectures are rejected. The extension output channel reports platform and VS Code version.
- **No tool in Chat:** choose a native agent supporting LM tools and enable the eight GenOffice tools in its picker. Local registration does not prove this conversation has enabled them. Real model conversations have not been accepted in this environment.
- **Skill not discovered:** run the install command in the desired workspace; inspect the generated SKILL.md and native `chat.useAgentSkills` setting. The extension never changes global model credentials. Keep Existing leaves user modifications intact; explicit updates preserve a complete backup under `.genoffice/skill-backups/`.
- **requires_target / stale binding:** bind an explicit open workbook and selection. Structure changes require rebinding; closing/reopening creates a new session. Same filenames in different directories are distinct documents.
- **editor_busy:** finish an in-progress cell edit or IME composition. The extension does not read old disk data as a substitute for unsaved content.
- **revision_conflict / stale import:** another edit or undo intervened. Read again; for Python results export and recompute. Refreshing only the revision does not authorize stale results.
- **Uncertain timeout:** query operationStatus with the original session/transaction or plan. A commit may have succeeded; do not invent a new transaction to retry it.
- **Python unavailable / dependency failure:** the native terminal and interpreter must run on the workspace host. Review the visible script and use an existing interpreter. The bundled example needs no third-party libraries. The extension does not install interpreters, libraries or bypass terminal permissions.
- **Import rejected:** check the registered snapshot, result hash, complete rectangle, types, target, exact next cursor and 30-minute expiry. Formulas require explicit formula mode. Cancellation leaves previous committed batches and their undo history; report partial counts.
- **External file conflict:** saving stops when the disk file changed or an alias is owned elsewhere. Preserve the unsaved work with Save As, or explicitly revert/reopen after reviewing the external file. Never force an old source onto it.
- **Save failed:** the document remains dirty. Check filesystem permissions/free space and choose another path. A failure after publication may report `fileWritten`; inspect that fact before retrying.
- **After a crash:** reopen the workspace. If the native tab does not return (observed on a first-ever crashed VS Code window), run **GenOffice: Recover Unsaved Workbook**. A verified full recovery copy restores a new session; the old undo timeline is not restored. Backups do not silently overwrite the original XLSX.
- **Protected/native objects:** original charts, merges and protected sheets may be preserved without being editable. Read capabilities; unsupported transformations fail closed. Session charts have a separate editor pane.
- **Slow large output:** the measured 100k output correctness run took about 11 minutes; it does not establish interactive speed. Do not disable history budgets or confuse maximum grid coordinates with verified capacity.

## Remote and platform status

The extension is declared `extensionKind: workspace`; source files, sidecar and exchange paths resolve in that extension host. An SSH URI from the desktop is not a host filesystem URI. Actual SSH Ubuntu 22.04 x64, Windows 11 x64 and macOS 14 arm64 verification require those environments; local boundary tests are not substitutes. The workflow uses explicitly labeled self-hosted Windows 11/macOS 14 arm64 runners and Ubuntu 22.04. They must be provisioned (including a display and Python for test harnesses) before those jobs can run.
