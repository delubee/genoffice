import { parsePivotDefinition } from '@genoffice/xlsx-gateway/gateway/xlsx-pivot'
import { workbookPivotDefinitionSchema } from '@genoffice/xlsx-gateway/shared/workbook-protocol'
import { constants } from 'node:fs'
import { copyFile, stat, chmod } from 'node:fs/promises'
import { join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { assertNotCancelled, WorkbookError, type OriginalBase } from './domain'
import type { WorkbookSource } from './ports'
import { hashFile, privateDirectory, RetainedResource } from './resources'
import { resolveDocumentPath } from './document-path'

export async function openWorkbookSnapshot(input: {
  uri: string
  storageDirectory: string
  source: WorkbookSource
  signal?: AbortSignal
}) {
  assertNotCancelled(input.signal)
  const url = new URL(input.uri)
  if (url.protocol !== 'file:')
    throw new WorkbookError(
      'unsupported_resource',
      'Only a file on the extension host is supported.',
    )
  const path = fileURLToPath(url),
    canonicalPath = await resolveDocumentPath(path),
    physical = await stat(canonicalPath)
  if (!physical.isFile())
    throw new WorkbookError('unsupported_resource', 'The workbook must be a regular file.')
  const directory = await privateDirectory(input.storageDirectory, 'base-')
  let sidecarId: string | undefined
  const owner = new RetainedResource(directory, async () => {
    if (sidecarId) await input.source.close(sidecarId)
  })
  try {
    const snapshotPath = join(directory, 'original.xlsx')
    await copyFile(canonicalPath, snapshotPath, constants.COPYFILE_EXCL)
    await chmod(snapshotPath, 0o400)
    assertNotCancelled(input.signal)
    const sha256 = await hashFile(snapshotPath)
    const opened = await input.source.openSnapshot(snapshotPath)
    sidecarId = opened.sessionId
    const manifest = await input.source.archiveManifest(snapshotPath)
    assertNotCancelled(input.signal)
    const metadata = await withPivotDefinitions(input.source, snapshotPath, opened.metadata)
    assertNotCancelled(input.signal)
    const original: OriginalBase = Object.freeze({
      path: snapshotPath,
      sha256,
      sourceUri: input.uri,
      sidecarSessionId: sidecarId,
    })
    return {
      original,
      metadata:
        metadata && typeof metadata === 'object'
          ? { ...metadata, readOnly: (physical.mode & 0o222) === 0 }
          : metadata,
      archiveManifest: manifest,
      canonicalPath,
      physicalId: `${physical.dev}:${physical.ino}`,
      retain: () => owner.retain(),
      dispose: () => owner.dispose(),
    }
  } catch (error) {
    await owner.dispose()
    throw error
  }
}

/**
 * File PivotTables: parsed definitions ride in each sheet's original metadata
 * so refresh runs on the canonical root without renderer archive access.
 * A part that cannot be read or parsed records its reason and refresh refuses it.
 */
async function withPivotDefinitions(source: WorkbookSource, path: string, metadata: unknown) {
  if (!source.readArchiveText || !metadata || typeof metadata !== 'object') return metadata
  const sheets = (metadata as { sheets?: unknown }).sheets
  if (!Array.isArray(sheets)) return metadata
  return {
    ...metadata,
    sheets: await Promise.all(
      sheets.map(async (sheet: { pivotTables?: { path: string; cachePath: string | null }[] }) => {
        const pivots = sheet.pivotTables ?? []
        if (!pivots.length) return sheet
        const pivotDefinitions = await Promise.all(
          pivots.slice(0, 100).map(async ({ path: part, cachePath }) => {
            if (!cachePath) return { path: part, error: 'missing_cache' }
            try {
              const [pivotXml, cacheXml] = await Promise.all([
                source.readArchiveText!(path, part),
                source.readArchiveText!(path, cachePath),
              ])
              return {
                path: part,
                cachePath,
                definition: workbookPivotDefinitionSchema.parse(
                  parsePivotDefinition(pivotXml, cacheXml),
                ),
              }
            } catch (error) {
              return { path: part, error: error instanceof Error ? error.message : String(error) }
            }
          }),
        )
        return { ...sheet, pivotDefinitions }
      }),
    ),
  }
}
