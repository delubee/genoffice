import type {
  CommitReceipt,
  DocumentDescriptor,
  PreparedChangeset,
  TransactionInput,
  WorkbookState,
} from './domain'
import type { ImportBatch } from './exchange/schemas'

export interface WorkbookSource {
  openSnapshot(path: string): Promise<{ sessionId: string; metadata: unknown }>
  readBaseRange(sessionId: string, sheetId: string, range: string): Promise<unknown>
  readFeatures(sessionId: string, sheetId: string): Promise<unknown>
  archiveManifest(path: string): Promise<unknown>
  close(sessionId: string): Promise<void>
}

export interface CandidateInterpreter {
  prepare(
    state: Readonly<WorkbookState>,
    intent: TransactionInput,
    signal?: AbortSignal,
  ): Promise<PreparedChangeset>
  prepareImport?(
    state: WorkbookState,
    intent: TransactionInput,
    batch: ImportBatch,
    signal?: AbortSignal,
  ): Promise<PreparedChangeset>
  dispose(): void | Promise<void>
}

export interface ResourceStore {
  writeChunk(data: Uint8Array): Promise<string>
  readChunk(reference: string): Promise<Uint8Array>
  publishManifest(name: string, data: unknown): Promise<void>
  readManifest(name: string): Promise<unknown>
  /** Metadata only. Content consumers must verify referenced chunks themselves. */
  readManifestMetadata?(name: string): Promise<unknown>
  delete(reference: string): Promise<void>
  collect?(retainedManifestNames: readonly string[]): Promise<void>
}

export interface PackageWriter {
  stage(
    originalPath: string,
    current: Readonly<WorkbookState>,
    targetPath: string,
    signal?: AbortSignal,
  ): Promise<{ stagedPath: string; hash: string }>
  publish(stagedPath: string, targetPath: string, expectedHash: string | null): Promise<void>
}

export interface HostLifecycle {
  committed(receipt: CommitReceipt, descriptor: DocumentDescriptor): void
  projectionChanged(stateId: string, revision: number): void
  historyChanged?(descriptor: DocumentDescriptor): void
  conflict(documentUri: string, reason: string): void
}
