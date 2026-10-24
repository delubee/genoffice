import { execFileSync } from 'node:child_process'
import { fileURLToPath } from 'node:url'

const cwd = fileURLToPath(new URL('../', import.meta.url))
const git = (...args) => execFileSync('git', args, { cwd, encoding: 'utf8' }).trim()
const exists = (ref) => {
  try {
    git('rev-parse', '--verify', '--quiet', ref)
    return true
  } catch {
    return false
  }
}

console.log(`Current branch: ${git('branch', '--show-current') || '(detached HEAD)'}`)
console.log(`Worktree: ${git('status', '--porcelain').length ? 'uncommitted changes' : 'clean'}`)
for (const remote of ['origin', 'upstream']) {
  try {
    console.log(`${remote} fetch: ${git('remote', 'get-url', remote)}`)
    console.log(`${remote} push: ${git('remote', 'get-url', '--push', remote)}`)
  } catch {
    console.log(`${remote}: not configured`)
  }
}
console.log('Local commits compared with the last fetched upstream/main:')
for (const branch of ['main', 'dev/genoffice', 'feature/vscode-office']) {
  if (!exists(`refs/heads/${branch}`)) {
    console.log(`  ${branch}: absent`)
    continue
  }
  if (!exists('refs/remotes/upstream/main')) {
    console.log(`  ${branch}: fetch upstream first`)
    continue
  }
  const [upstreamOnly, branchOnly] = git(
    'rev-list',
    '--left-right',
    '--count',
    `upstream/main...${branch}`,
  ).split(/\s+/)
  console.log(`  ${branch}: ${branchOnly} branch-only, ${upstreamOnly} upstream-only`)
}
console.log('Uncommitted work is not included in commit counts. No fetch or mutation performed.')
