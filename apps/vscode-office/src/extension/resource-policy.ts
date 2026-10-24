import { WorkbookError } from '@genoffice/workbook-session/domain'

export function selectWorkspace<T extends { scheme: string; toString(): string }>(input: {
  trusted: boolean
  roots: readonly T[]
  requested?: string
  containing?: T
}): T {
  if (!input.trusted)
    throw new WorkbookError('untrusted_workspace', 'Trust this workspace before exchanging files.')
  const selected = input.requested
    ? input.roots.find((root) => root.toString() === input.requested)
    : (input.containing ?? (input.roots.length === 1 ? input.roots[0] : undefined))
  if (!selected)
    throw new WorkbookError(
      'requires_workspace',
      'Choose an explicit workspace root for this snapshot.',
      { workspaceUris: input.roots.map((root) => root.toString()) },
    )
  if (selected.scheme !== 'file')
    throw new WorkbookError(
      'unsupported_resource',
      'The exchange workspace must map to an extension-host filesystem directory.',
    )
  return selected
}

export function assertHostFile(uri: { scheme: string; authority: string }, authority: string) {
  if (uri.scheme !== 'file' || uri.authority !== authority)
    throw new WorkbookError(
      'path_not_allowed',
      'Use the result file URI on this workspace host and authority.',
    )
}
