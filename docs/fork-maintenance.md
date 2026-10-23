# Maintaining this fork

## Remotes and branches

| Name                    | Role                                                             |
| ----------------------- | ---------------------------------------------------------------- |
| `origin`                | Our fork: `git@github.com:delubee/genoffice.git`                 |
| `upstream`              | Original project: `https://github.com/genspark-ai/genoffice.git` |
| `main`                  | Clean upstream mirror; tracks `upstream/main`                    |
| `dev/genoffice`         | Integration branch for our own features                          |
| `feature/vscode-office` | Current VS Code editor work                                      |
| `feature/*`             | Subsequent isolated features                                     |

The local default push remote is `origin`. The local `upstream` push URL is
`DISABLED`, preventing accidental pushes to the original project. Remote and
branch configuration lives in `.git/config` and is not distributed by a commit.

The initial fork baseline is `285540a4`. Existing uncommitted VS Code changes
were carried to `feature/vscode-office` without changing their contents. The
feature and integration branches start at that baseline. Uncommitted files are
worktree contents, not saved branch history; switching branches is not a backup.
Review and commit them on the feature branch before switching or integrating.

During setup on 2026-10-23, the clean local `main` was fast-forwarded to
`e416f9ba8312` (six upstream commits after the original baseline). Those commits
have not been merged into the feature or integration branch. Review and commit
the current feature before bringing that upstream update into its working code.

## Set up another clone

```bash
git remote add -t main upstream https://github.com/genspark-ai/genoffice.git
git remote set-url --push upstream DISABLED
git config remote.pushDefault origin
git config push.default simple
git fetch --no-tags upstream main
git branch --set-upstream-to=upstream/main main
```

If `upstream` already exists, inspect its URL instead of adding it again. A
shared clone can fetch `dev/genoffice` and feature branches once they have been
explicitly pushed to the fork. Local branches created during setup are not
automatically published.

## Review and save our current work

```bash
git status --short
git diff --stat
git diff
```

Inspect untracked source files as well: `git diff` does not show their contents.
Stage explicit, reviewed paths rather than blindly adding the whole worktree.
Use `git diff --cached` before each commit. Do not stage generated packages,
workspace documents, recovery data, credentials or local agent tooling.

Suggested review groups for the existing work:

1. Fork maintenance, ignore rules and repository documentation.
2. Shared gateway/session extraction and compatibility changes.
3. Browser spreadsheet editor and VS Code document lifecycle/history.
4. Native tools, typed Python exchange and packaged document skill.
5. Platform packaging, workflow, fixtures and verification records.

Some groups depend on earlier code. Each commit intended as an independently
usable checkpoint must include its required package configuration and pass the
relevant gates. Do not create artificial independent commits that cannot build.

## Integrate a feature

After reviewing, committing and testing the feature, use a clean worktree:

```bash
git switch dev/genoffice
git merge --no-ff feature/vscode-office
```

Inspect the merged result and run the relevant gates before explicitly pushing:

```bash
git push -u origin dev/genoffice
git push -u origin feature/vscode-office
```

Publishing these branches is a separate user action; setup does not push them.

## Bring in upstream updates

Start with a clean worktree and no merge/rebase in progress. Commit the current
feature first; do not move unsaved feature work onto `main`.

```bash
git fetch --no-tags upstream main
git switch main
git merge --ff-only upstream/main
git switch dev/genoffice
git merge main
```

If the fast-forward fails, investigate the divergence instead of resetting or
force-pushing `main`. If the integration merge conflicts, resolve and review
each conflict, stage the resolved files and finish with `git merge --continue`.
Use `git merge --abort` to return to the pre-merge state when needed.

Run the regression gates below, then explicitly push the fork's updated
branches. Shared integration branches use merges to retain published history.
Rebase is suitable for private, unpublished feature commits; avoid rewriting
shared branch history.

## Identify our code and integration risks

```bash
npm run fork:status
git log --oneline upstream/main..dev/genoffice
git diff --stat upstream/main...dev/genoffice
git diff upstream/main...dev/genoffice
```

`fork:status` is read-only and uses the last fetched upstream reference. For
the current unmerged feature, replace `dev/genoffice` with
`feature/vscode-office`. Uncommitted work is not included in branch comparisons.

| Area                              | Our changes / upstream integration risk                                         |
| --------------------------------- | ------------------------------------------------------------------------------- |
| `apps/vscode-office/`             | Custom editor, native tools, exchange adapter, packaging and skill resources    |
| `packages/workbook-session/`      | Node session authority, source/save/history/recovery and typed exchange         |
| `packages/sheets-editor/`         | Shared browser editor and projection/candidate ports                            |
| `packages/xlsx-gateway/`          | Shared protocol/candidate changes and existing save/adapter compatibility       |
| `apps/sheets/native/xlsx-engine/` | Typed cell/error preservation; rerun native tests after upstream parser changes |
| `apps/sheets/`                    | Compatibility re-exports and actual Electron build configuration                |
| Root manifests, lockfile, CI      | Workspace/dependency/build gate integration                                     |

Keep future fork features in dedicated directories and make focused upstream
file changes. This table is an inventory; Git commits remain the source of
truth for ownership and history.

## Minimum checks after integration

```bash
npm run vscode:typecheck
npm run vscode:test
npm run test -w @genoffice/xlsx-gateway
npm run test -w @genoffice/sheets
npm run test:e2e -w @genoffice/sheets-editor
npm run vscode:build
npm run build -w @genoffice/shell
```

Run native VS Code integration and installed-VSIX tests on affected platforms.
Headless GUI tests need a display server. Full performance/real-agent/platform
acceptance is documented separately under `apps/vscode-office/test-results/`;
an upstream merge does not establish those gates automatically.

## Local skills versus product resources

`.agents/skills/`, `.codex/skills/` and `.github/skills/` at the repository root
are ignored as local development tooling. Existing upstream-tracked files under
`skills/genoffice/` retain their upstream history. The extension's
`apps/vscode-office/skills/genoffice-workbook/` is a product resource required for
skill installation and reproducible VSIX builds, and remains eligible for Git.
Ignore rules do not untrack files that are already committed.
