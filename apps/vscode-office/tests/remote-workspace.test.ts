import { expect, it } from 'vitest'
import { selectWorkspace, assertHostFile } from '../src/extension/resource-policy'

const local = { scheme: 'file', toString: () => 'file:///workspace' }
const other = { scheme: 'file', toString: () => 'file:///other' }
it('requires trust and explicit workspace identity when roots are ambiguous', () => {
  expect(() => selectWorkspace({ trusted: false, roots: [local] })).toThrow(/Trust/)
  expect(() => selectWorkspace({ trusted: true, roots: [local, other] })).toThrow(/explicit/)
  expect(
    selectWorkspace({ trusted: true, roots: [local, other], requested: local.toString() }),
  ).toBe(local)
  expect(() =>
    selectWorkspace({ trusted: true, roots: [local], requested: other.toString() }),
  ).toThrow()
})
it('accepts host filesystem resources and rejects desktop SSH/virtual or foreign authority results', () => {
  const ssh = {
    scheme: 'vscode-remote',
    toString: () => 'vscode-remote://ssh-remote+host/workspace',
  }
  expect(() => selectWorkspace({ trusted: true, roots: [ssh] })).toThrow(/extension-host/)
  expect(() =>
    assertHostFile({ scheme: 'vscode-remote', authority: 'ssh-remote+host' }, ''),
  ).toThrow()
  expect(() => assertHostFile({ scheme: 'file', authority: 'other-host' }, '')).toThrow()
  expect(() => assertHostFile({ scheme: 'file', authority: '' }, '')).not.toThrow()
})
