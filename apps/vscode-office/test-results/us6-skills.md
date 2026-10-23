# US6 — workspace skill evidence (2026-10-22)

Packaged skill: `genoffice-workbook`, metadata version 0.1.0. The Skill Creator structural validator passes. The full visible standard-library Python example is bundled with the skill; operation and file-exchange references match the eight registered v1 tools.

Pure installer tests cover explicit root selection boundaries, unchanged and user-modified installation, keep/cancel without mutation, staged update with complete backup, digest checks and missing-tool guidance. Outside symlinks are rejected. Backups live outside the native agent skill scan directory. No global model setting is modified.

A real VS Code 1.109.5 native extension-host test passed: **GenOffice: Install or Update Workbook Skill** installs into the explicit workspace `.github/skills/genoffice-workbook`; **GenOffice: Workbook Skill Status** reports version, all eight locally registered tools, and later user modification. `skill-native-proof.json` is the machine-readable evidence. The test only creates/removes its own isolated installation.

Native VS Code bundled implementation confirms `.github/skills` discovery and the `chat.useAgentSkills` setting. The extension offers an explicit settings link and never enables it silently. Stable APIs do not expose per-conversation tool selection; status reports this limitation.

Real logged-in agent discovery, read/edit/Python conversations and rejection of a native tool-confirmation dialog are **unverified**. Native tool harness calls and installer tests do not substitute for these checks.
