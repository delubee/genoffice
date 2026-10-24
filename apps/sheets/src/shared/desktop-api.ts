import { z } from 'zod'
import { normalizeCustomProfile } from '@genoffice/ai-provider/browser'

import {
  MAX_CREATE_DOCUMENT_CONTENT_CHARS,
  MAX_CREATE_DOCUMENT_TITLE_CHARS,
  MAX_CSV_EXPORT_CHARS,
  MAX_PDF_TEMPLATE_CHARS,
} from './ipc-channels'
import type {
  AiChatRequest,
  AiChatResponse,
  AiSettings,
  AiStreamChunk,
  AiStreamRequest,
  GenSparkAccountStatus,
} from '@genoffice/ai-provider'
import type { AiPanelPrefs } from '@genoffice/ui'

// edit schemas shared with the xlsx gateway package; re-exported so IPC consumers keep one import site
export {
  CHART_CATEGORY_WIRE_MAX,
  CHART_TEXT_WIRE_MAX,
  MAX_PATCH_ENTRY_BYTES,
  editableBorderStyleSchema,
  workbookChartEditSchema,
  workbookStyleEditSchema,
  workbookVisualEditSchema,
} from '@genoffice/xlsx-gateway/shared/edit-schemas'
export type {
  WorkbookChartEdit,
  WorkbookRichRun,
  WorkbookStyleEdit,
  WorkbookVisualEdit,
} from '@genoffice/xlsx-gateway/shared/edit-schemas'

export * from '@genoffice/xlsx-gateway/shared/workbook-protocol'
import type { WorkbookFile, WorkbookSaveRequest, WorkbookSaveEditsBegin, WorkbookSaveEditsChunk, WorkbookSaveEditsAbort, WorkbookSaveResult, WorkbookRangeRequest, WorkbookFormulaCellsRequest, WorkbookFormulaCellsResult, WorkbookFindCellsRequest, WorkbookFindCellsResult, WorkbookRowOutlineResult, WorkbookRangeResult, WorkbookRecalcRequest, WorkbookRecalcResult, WorkbookMediaRequest, WorkbookMediaResult, WorkbookPivotRequest, WorkbookPivotDefinition } from '@genoffice/xlsx-gateway/shared/workbook-protocol'

export const localImageRequestSchema = z
  .object({
    path: z.string().min(1).max(1024),
  })
  .strict()

export const localImageResultSchema = z
  .object({
    mediaType: z.enum(['image/png', 'image/jpeg', 'image/gif']),
    /// ~20MB decoded, mirroring the picker-based insert limit.
    base64: z.string().min(1).max(28_000_000),
  })
  .strict()

/// Insert → Screenshot: enumerate capturable windows/screens for the picker
/// grid, then grab the chosen source at full resolution. 'denied' means the
/// OS blocks capture (macOS Screen Recording permission) — the picker shows
/// guidance instead of an empty grid.
export const screenSourcesResultSchema = z
  .object({
    status: z.enum(['ok', 'denied']),
    sources: z.array(
      z
        .object({
          id: z.string().min(1),
          name: z.string(),
          kind: z.enum(['screen', 'window']),
          /// Small preview as a data URL; empty when the OS returned none.
          thumbnail: z.string(),
        })
        .strict(),
    ),
  })
  .strict()

export const screenCaptureRequestSchema = z
  .object({
    id: z.string().min(1).max(256),
  })
  .strict()

export const screenCaptureResultSchema = z
  .object({
    mediaType: z.literal('image/png'),
    base64: z.string().min(1).max(28_000_000),
    width: z.number().int().positive(),
    height: z.number().int().positive(),
  })
  .strict()

export type LocalImageRequest = z.infer<typeof localImageRequestSchema>
export type LocalImageResult = z.infer<typeof localImageResultSchema>
export type ScreenSourcesResult = z.infer<typeof screenSourcesResultSchema>
export type ScreenCaptureRequest = z.infer<typeof screenCaptureRequestSchema>
export type ScreenCaptureResult = z.infer<typeof screenCaptureResultSchema>

// ---- AI settings + chat/stream: canonical types live in @genoffice/ai-provider,
// shared with apps/docs. Validated here like every other renderer→main request in
// this file; the validated shape is cast to AiSettings at the main-process call
// site, which has every known provider key once merged through
// resolveAiSettings/defaultAiSettings. ----

const aiCustomCompatSchema = z
  .object({
    omitTemperature: z.boolean().optional(),
    echoReasoning: z.boolean().optional(),
    vision: z.boolean().optional(),
    useMaxCompletionTokens: z.boolean().optional(),
  })
  .strict()

const aiCustomModelProfileSchema = z
  .object({
    id: z.string().trim().min(1),
    alias: z.string().optional(),
    name: z.string().optional(),
    ...aiCustomCompatSchema.shape,
    maxTokens: z.number().int().positive().optional(),
    contextWindow: z.number().int().positive().optional(),
  })
  .strict()

// Keep IPC fields explicit, then use the canonical normalizer for header
// legality and profile size bounds, just like the other apps' settings gate.
const aiCustomProfileSchema = z
  .object({
    api: z.enum(['openai-compatible', 'anthropic', 'gemini']).optional(),
    apiKeyEnv: z.string().optional(),
    headers: z.record(z.string(), z.string()).optional(),
    idleTimeoutMs: z.number().int().positive().optional(),
    compat: aiCustomCompatSchema.optional(),
    bodyExtras: z.record(z.string(), z.unknown()).optional(),
    modelProfiles: z.array(aiCustomModelProfileSchema).optional(),
    probedProfiles: z.array(aiCustomModelProfileSchema).optional(),
  })
  .strict()

const aiProviderConfigSchema = z
  .object({
    apiKey: z.string(),
    model: z.string(),
    baseUrl: z.string().optional(),
    cliPath: z.string().optional(),
    profile: aiCustomProfileSchema.transform(normalizeCustomProfile).optional(),
  })
  .strict()

const aiMediaProviderConfigSchema = z
  .object({
    apiKey: z.string(),
    baseUrl: z.string().optional(),
    imageModel: z.string(),
    analysisModel: z.string(),
  })
  .strict()

const aiCustomEndpointSchema = z
  .object({
    id: z.string().min(1),
    name: z.string(),
    baseUrl: z.string(),
    apiKey: z.string(),
    model: z.string(),
    models: z.array(z.string()).optional(),
    ...aiCustomProfileSchema.shape,
  })
  .strict()
  .transform(({ id, name, baseUrl, apiKey, model, models, ...profile }) => ({
    id,
    name,
    baseUrl,
    apiKey,
    model,
    ...(models !== undefined ? { models } : {}),
    ...normalizeCustomProfile(profile),
  }))

export const aiSettingsInputSchema = z
  .object({
    provider: z.string().min(1),
    providers: z.record(z.string(), aiProviderConfigSchema),
    customEndpoints: z.array(aiCustomEndpointSchema).optional(),
    customEndpoint: z.string().optional(),
    gskToolsEnabled: z.boolean().optional(),
    media: z
      .object({
        imageProvider: z.string().min(1).optional(),
        analysisProvider: z.string().min(1).optional(),
        videoAnalysisProvider: z.string().min(1).optional(),
        // pre-catalog single choice, still accepted on read
        provider: z.string().min(1).optional(),
        providers: z.record(z.string(), aiMediaProviderConfigSchema),
      })
      .strict()
      .optional(),
    search: z
      .object({
        provider: z.string().min(1),
        providers: z.record(z.string(), z.object({ apiKey: z.string() }).strict()),
      })
      .strict()
      .optional(),
    // bounds are enforced by clampMaxOutputTokens on read; the schema only
    // rejects nonsense (this object is .strict(), so an omitted key here would
    // make the whole settings save fail)
    maxOutputTokens: z.number().int().positive().optional(),
  })
  .strict()

const agentToolResultSchema = z
  .object({
    id: z.string(),
    name: z.string(),
    output: z.string(),
    isError: z.boolean().optional(),
  })
  .strict()

const agentToolCallSchema = z
  .object({
    id: z.string(),
    name: z.string(),
    input: z.record(z.string(), z.unknown()),
    signature: z.string().optional(),
  })
  .strict()

/// Inline vision input on a user turn (image attachments, base64 without data: prefix).
const agentImageSchema = z
  .object({
    base64: z.string().min(1),
    mime: z.string().min(1).max(64),
  })
  .strict()

const agentMessageSchema = z.union([
  z
    .object({
      role: z.literal('user'),
      text: z.string(),
      images: z.array(agentImageSchema).max(20).optional(),
    })
    .strict(),
  z
    .object({
      role: z.literal('assistant'),
      text: z.string(),
      toolCalls: z.array(agentToolCallSchema).optional(),
      // captured model thinking, echoed back for interleaved-thinking models
      reasoning: z.string().optional(),
    })
    .strict(),
  z.object({ role: z.literal('tool'), results: z.array(agentToolResultSchema) }).strict(),
])

const agentToolDefSchema = z
  .object({
    name: z.string(),
    description: z.string(),
    inputSchema: z.record(z.string(), z.unknown()),
  })
  .strict()

const MAX_AI_MESSAGES = 500
const MAX_AI_TOOLS = 50

export const aiChatRequestSchema = z
  .object({
    settings: aiSettingsInputSchema,
    system: z.string(),
    user: z.string(),
  })
  .strict()

export const aiStreamRequestSchema = z
  .object({
    requestId: z.string().min(1),
    sessionId: z.string().uuid().optional(),
    settings: aiSettingsInputSchema,
    system: z.string(),
    messages: z.array(agentMessageSchema).max(MAX_AI_MESSAGES),
    tools: z.array(agentToolDefSchema).max(MAX_AI_TOOLS).optional(),
    maxTokens: z.number().int().positive().optional(),
  })
  .strict()

export type AiSettingsInput = z.infer<typeof aiSettingsInputSchema>
export type AiChatRequestInput = z.infer<typeof aiChatRequestSchema>
export type AiStreamRequestInput = z.infer<typeof aiStreamRequestSchema>

const pdfPageVariantSchema = z
  .object({
    headerTemplate: z.string().min(1).max(MAX_PDF_TEMPLATE_CHARS).optional(),
    footerTemplate: z.string().min(1).max(MAX_PDF_TEMPLATE_CHARS).optional(),
  })
  .strict()

/// A rendered print job: the renderer lays the sheet out as HTML, the main
/// process turns it into a PDF via a hidden window.
export const workbookExportPdfRequestSchema = z
  .object({
    fileName: z.string().min(1).max(255),
    html: z.string().min(1).max(20_000_000),
    landscape: z.boolean(),
    pageSize: z.union([
      z.enum(['A3', 'A4', 'A5', 'Legal', 'Letter', 'Tabloid']),
      /// Custom size in inches.
      z
        .object({
          width: z.number().positive().max(100),
          height: z.number().positive().max(100),
        })
        .strict(),
    ]),
    /// Inches.
    margins: z
      .object({
        top: z.number().min(0).max(3),
        bottom: z.number().min(0).max(3),
        left: z.number().min(0).max(3),
        right: z.number().min(0).max(3),
      })
      .strict(),
    scale: z.number().min(0.1).max(2),
    /// Chromium print header/footer templates (rendered in the margin
    /// boxes; `pageNumber`/`totalPages` spans resolve per page). These are
    /// the odd-page (default) templates.
    headerTemplate: z.string().min(1).max(MAX_PDF_TEMPLATE_CHARS).optional(),
    footerTemplate: z.string().min(1).max(MAX_PDF_TEMPLATE_CHARS).optional(),
    /// Excel's differentFirst / differentOddEven variants. Presence means
    /// the variant is active (an empty object prints page 1 / even pages
    /// with no header or footer); the main process renders them in extra
    /// printToPDF passes and stitches the pages together.
    firstPage: pdfPageVariantSchema.optional(),
    evenPages: pdfPageVariantSchema.optional(),
    /// Headless export mode only (--headless-export): write here instead of
    /// opening the save dialog. Ignored by a normal GUI session.
    outPath: z.string().min(1).max(4096).optional(),
    /// Print only: silent print on this printer ('' = the system default);
    /// absent opens the system print dialog.
    deviceName: z.string().max(255).optional(),
    copies: z.number().int().min(1).max(999).optional(),
    collate: z.boolean().optional(),
  })
  .strict()

export const printerInfoSchema = z
  .object({
    name: z.string().max(255),
    displayName: z.string().max(255),
  })
  .strict()
export type PrinterInfo = z.infer<typeof printerInfoSchema>

export const workbookExportPdfResultSchema = z.union([
  z.object({ canceled: z.literal(true) }).strict(),
  z.object({ canceled: z.literal(false), path: z.string().min(1) }).strict(),
])

export type WorkbookExportPdfRequest = z.infer<typeof workbookExportPdfRequestSchema>
export type WorkbookExportPdfResult = z.infer<typeof workbookExportPdfResultSchema>

/// Print of the same laid-out HTML through the system print dialog. `ok: false`
/// without an error is the user closing the dialog.
export const workbookPrintResultSchema = z.union([
  z.object({ ok: z.literal(true) }).strict(),
  z.object({ ok: z.literal(false), error: z.string().optional() }).strict(),
])
export type WorkbookPrintResult = z.infer<typeof workbookPrintResultSchema>

/// CSV export of the active sheet: the renderer serializes display values,
/// the main process runs the loss warning + save dialog and writes the bytes.
export const workbookExportCsvRequestSchema = z
  .object({
    fileName: z.string().min(1).max(255),
    content: z.string().max(MAX_CSV_EXPORT_CHARS),
    /// The exported sheet carries formulas — CSV keeps their values only, so
    /// the main process offers Save As .xlsx first (Excel's warning flow).
    hasFormulas: z.boolean(),
    /// Set when the workbook has more sheets than the exported one.
    activeSheetName: z.string().max(255).optional(),
    /// Write straight to this path (from the Save As dialog's CSV pick)
    /// instead of opening another save dialog.
    targetPath: z.string().min(1).optional(),
  })
  .strict()

export const workbookExportCsvResultSchema = z.union([
  z
    .object({
      canceled: z.literal(true),
      /// The user chose "Save as .xlsx" in the formula-loss warning; the
      /// renderer routes into the regular Save As flow instead.
      saveAsXlsxInstead: z.boolean().optional(),
    })
    .strict(),
  z.object({ canceled: z.literal(false), path: z.string().min(1) }).strict(),
])

export type WorkbookExportCsvRequest = z.infer<typeof workbookExportCsvRequestSchema>
export type WorkbookExportCsvResult = z.infer<typeof workbookExportCsvResultSchema>

/// AI create_document: a new standalone file written into the default save
/// folder (no dialog) under a unique sanitized name and opened in a new tab.
/// xlsx/csv carry one worksheet's serialized CSV in content (the renderer
/// reads the grid, mirroring the manual CSV export); docx/pdf/md carry
/// AI-authored HTML/Markdown, routed by the shell into the docs-owned
/// creation flow (#960).
export const workbookCreateDocumentRequestSchema = z
  .object({
    type: z.enum(['xlsx', 'csv', 'docx', 'pdf', 'md', 'html']),
    title: z.string().min(1).max(MAX_CREATE_DOCUMENT_TITLE_CHARS),
    content: z.string().min(1).max(MAX_CSV_EXPORT_CHARS),
    /// xlsx only: the worksheet name inside the created workbook (the source
    /// sheet's name, so it already satisfies Excel's naming rules).
    sheetName: z.string().min(1).max(31).optional(),
  })
  .strict()
  .superRefine((request, ctx) => {
    if (
      request.type !== 'xlsx' &&
      request.type !== 'csv' &&
      request.content.length > MAX_CREATE_DOCUMENT_CONTENT_CHARS
    ) {
      ctx.addIssue({
        code: z.ZodIssueCode.custom,
        path: ['content'],
        message: `content must not exceed ${MAX_CREATE_DOCUMENT_CONTENT_CHARS} characters for docx/pdf/md`,
      })
    }
  })

/// Loose result shape (ok: boolean, not a literal union) so the shell can
/// wire the docs-owned creator in directly — its CreateDocumentResult is
/// structurally identical.
export const workbookCreateDocumentResultSchema = z
  .object({
    ok: z.boolean(),
    path: z.string().min(1).optional(),
    error: z.string().optional(),
  })
  .strict()

export type WorkbookCreateDocumentRequest = z.infer<typeof workbookCreateDocumentRequestSchema>
export type WorkbookCreateDocumentResult = z.infer<typeof workbookCreateDocumentResultSchema>

// ---- Chat attachments (local files fed to the agent via tools; same structure
// as apps/docs and apps/slides) ----

/** Image attachment extensions: no text extraction; read as base64 on send and
 * passed to the model as a multimodal image with the user message */
export const ATTACHMENT_IMAGE_EXTS = new Set(['png', 'jpg', 'jpeg', 'gif', 'webp'])

export interface AttachmentMeta {
  /** Absolute local path; the file never leaves the machine */
  path: string
  name: string
  /** Lowercase extension, no dot */
  ext: string
  sizeBytes: number
}

export interface AttachmentAddResult {
  accepted: AttachmentMeta[]
  /** Per-file rejection reason (too large/unsupported type/unreadable) */
  rejected: string[]
}

export interface AttachmentReadResult {
  ok: boolean
  error?: string
  name?: string
  /** Total character count of the extracted text */
  totalChars?: number
  /** The requested chunk */
  text?: string
  offset?: number
}

/** Raw bytes of an image attachment (multimodal input) */
export interface AttachmentImageResult {
  ok: boolean
  /** raw base64 (no data: prefix) */
  base64?: string
  mime?: string
  error?: string
}

export type UiTheme = 'light' | 'dark' | 'system'

/**
 * Document page theme preference (genoffice#1811): what the editors' canvas/paper does
 * relative to the UI theme. 'follow' keeps the previous single-theme behavior.
 */
export type DocTheme = 'follow' | 'light' | 'dark'

/** shell-wide AutoSave default; updatedAt is 0 until the user has ever set it */
export interface AutoSaveDefault {
  on: boolean
  updatedAt: number
}

/** Autosave-recovery prompt raised by main during workbook open (strings pre-localized). */
export interface RecoveryPromptPayload {
  title: string
  body: string
  restoreLabel: string
  discardLabel: string
  /** base name of the workbook being opened */
  fileName: string
  /** mtime of the recovery copy (epoch ms) — when the unsaved work was last autosaved */
  savedAtMs: number
}

/** MCP visible-grid bridge message (shell → renderer, correlated by requestId). */
export interface McpCommandMessage {
  requestId: string
  command: 'apply_ops' | 'read_sheet' | 'save_sheet'
  payload: unknown
}

/** MCP visible-grid bridge reply (renderer → shell). */
export interface McpCommandResult {
  requestId: string
  ok: boolean
  result?: unknown
  error?: string
}

export interface DesktopApi {
  /** open any shell-routed document path (an AI answer's filenav citation) */
  openSourcePath(path: string): Promise<void>
  /** current UI language (persisted by the shell in app-settings.json) */
  getLanguage(): Promise<'zh' | 'en' | 'ja' | 'ko' | 'fr' | 'de' | 'es' | 'th' | 'id' | 'ru' | 'ar'>
  /** language switched from the shell home page */
  onLanguageChanged(
    handler: (
      lang: 'zh' | 'en' | 'ja' | 'ko' | 'fr' | 'de' | 'es' | 'th' | 'id' | 'ru' | 'ar',
    ) => void,
  ): () => void
  /** current UI theme preference (persisted by the shell in app-settings.json) */
  getTheme(): Promise<UiTheme>
  /** theme switched from the shell home page */
  onThemeChanged(handler: (theme: UiTheme) => void): () => void
  /** current document page theme preference (genoffice#1811, persisted by the shell in app-settings.json) */
  getDocumentTheme(): Promise<DocTheme>
  /** document page theme switched from the shell home page */
  onDocumentThemeChanged(handler: (theme: DocTheme) => void): () => void
  /** shell-wide AutoSave default (see useAutoSavePref) */
  getAutoSaveDefault(): Promise<AutoSaveDefault>
  onAutoSaveDefaultChanged(handler: (value: AutoSaveDefault) => void): () => void
  /** AI panel text size + chat-input spellcheck (Settings → General in the shell) */
  getAiPanelPrefs(): Promise<AiPanelPrefs>
  setAiPanelPrefs(patch: Partial<AiPanelPrefs>): Promise<AiPanelPrefs>
  onAiPanelPrefsChanged(handler: (prefs: AiPanelPrefs) => void): () => void
  /**
   * the user pressed the shell chrome (tab strip) or started dragging the
   * window — no DOM event or blur reaches this view, so the shell relays the
   * press for dismissing open popovers
   */
  onChromePressed(handler: () => void): () => void
  selectWorkbook(): Promise<WorkbookFile | null>
  /** multi-select xlsx picker; each file opens a secondary sidecar session for merge reads */
  selectWorkbooksForMerge(): Promise<WorkbookFile[] | null>
  /** open explicit spreadsheet paths (chat attachments) as merge-source sessions */
  openWorkbooksForMerge(paths: string[]): Promise<WorkbookFile[] | null>
  /// Re-open a known path through the normal open path (the pipeline
  /// selectWorkbook runs) to adopt a live session after a sidecar crash.
  /// Not the merge-source open: a recovered workbook is a normal session, and
  /// a csv/xls is imported as the workbook itself, not as a merge input.
  reopenWorkbook(path: string): Promise<WorkbookFile | null>
  /// The sidecar process died: every session id this renderer holds is gone.
  /// A positive crash signal — the main process raises it only for an actual
  /// process death, never for a session it closed or swapped on purpose.
  onSidecarCrashed(callback: () => void): () => void
  readWorkbookRange(request: WorkbookRangeRequest): Promise<WorkbookRangeResult>
  readWorkbookFormulas(request: WorkbookFormulaCellsRequest): Promise<WorkbookFormulaCellsResult>
  findWorkbookCells(request: WorkbookFindCellsRequest): Promise<WorkbookFindCellsResult>
  readWorkbookRowOutline(request: WorkbookFormulaCellsRequest): Promise<WorkbookRowOutlineResult>
  recalcWorkbook(request: WorkbookRecalcRequest): Promise<WorkbookRecalcResult>
  readWorkbookMedia(request: WorkbookMediaRequest): Promise<WorkbookMediaResult>
  readPivotDefinition(request: WorkbookPivotRequest): Promise<WorkbookPivotDefinition>
  readLocalImage(request: LocalImageRequest): Promise<LocalImageResult>
  captureScreenSources(): Promise<ScreenSourcesResult>
  /// null when the source vanished between listing and capture.
  captureScreenSource(request: ScreenCaptureRequest): Promise<ScreenCaptureResult | null>
  saveWorkbookEdits(request: WorkbookSaveRequest): Promise<WorkbookSaveResult>
  /// Chunked upload of a large save's cell edits (above MAX_SAVE_EDITS):
  /// begin a transfer, send ordered slices, then reference the transfer via
  /// editsTransferId in the following save/recovery request.
  beginSaveEditsTransfer(request: WorkbookSaveEditsBegin): Promise<void>
  sendSaveEditsChunk(request: WorkbookSaveEditsChunk): Promise<void>
  /// Frees an unconsumed transfer after a failed upload or save; silent no-op
  /// if the transfer was already consumed or expired.
  abortSaveEditsTransfer(request: WorkbookSaveEditsAbort): Promise<void>
  /// Crash-recovery copy of the pending edits, written under userData.
  /// Best-effort: never prompts, never touches the opened file.
  writeWorkbookRecovery(request: WorkbookSaveRequest): Promise<{ ok: boolean }>
  /// Rename a still-untitled workbook after AI-generated content (no-op unless
  /// the file still carries the shell's auto-created untitled name).
  autoRenameWorkbook(
    sessionId: string,
    baseName: string,
  ): Promise<{ renamed: boolean; name?: string }>
  exportPdf(request: WorkbookExportPdfRequest): Promise<WorkbookExportPdfResult>
  printWorkbook(request: WorkbookExportPdfRequest): Promise<WorkbookPrintResult>
  listPrinters(): Promise<PrinterInfo[]>
  exportCsv(request: WorkbookExportCsvRequest): Promise<WorkbookExportCsvResult>
  /// First Save of a CSV session: native "keep this format?" dialog.
  confirmCsvSave(): Promise<'csv' | 'xlsx' | 'cancel'>
  /// AI create_document: write a new standalone file into the default save
  /// folder (no dialog) and open it in a new tab.
  createDocument(request: WorkbookCreateDocumentRequest): Promise<WorkbookCreateDocumentResult>
  closeWorkbook(sessionId: string): Promise<void>
  openExternal(url: string): Promise<void>
  /// Application-menu File commands (Open/Save/Save As); returns unsubscribe.
  onMenuAction(callback: (action: MenuAction) => void): () => void
  /// The open workbook was renamed on disk (renamed in the shell Home list);
  /// emits the new file name.
  onWorkbookRenamed(callback: (newName: string) => void): () => void
  /// Mirrors the pending-edit badge to the main process for the close guard.
  notifyPendingEdits(count: number): void
  /// Main asks the renderer to save before closing; reply via reportCloseSaveResult.
  onCloseSaveRequest(callback: () => void): () => void
  reportCloseSaveResult(ok: boolean): void
  /// Main found a newer autosaved recovery copy while opening a workbook; the
  /// renderer shows the styled restore prompt and replies via replyRecoveryPrompt.
  /// Strings arrive pre-localized from the main process.
  onRecoveryPrompt(callback: (prompt: RecoveryPromptPayload) => void): () => void
  replyRecoveryPrompt(restore: boolean): void
  /// Returns true once when this tab was opened via "New Spreadsheet" from the
  /// shell home.
  consumeNewBlankWorkbook(): Promise<boolean>
  /// MCP visible-grid bridge (see renderer/mcp-bridge.ts): the shell pushes one
  /// command at a time; the renderer executes it and reports the correlated
  /// result. onMcpCommand returns unsubscribe.
  onMcpCommand(callback: (message: McpCommandMessage) => void): () => void
  reportMcpResult(result: McpCommandResult): void
  signalMcpReady(): void
  /// Is a shell-queued workbook path still waiting to be opened? (The shell's
  /// 'open' nudge loop can time out on slow cold starts; the renderer pulls.)
  hasQueuedWorkbook(): Promise<boolean>
  /// Headless export mode (--headless-export): the PDF path this hidden
  /// renderer must export to; null in a normal session.
  consumeHeadlessExport(): Promise<string | null>
  /// Headless export mode: report the export outcome so the main process can quit.
  headlessExportDone(result: { ok: boolean; error?: string }): void
  /// Author name for threaded comments (Settings → General, falling back to the OS account).
  getUserDisplayName(): Promise<string>
  onUserDisplayNameChanged(handler: (name: string) => void): () => void
  getAiSettings(): Promise<AiSettings>
  setAiSettings(settings: AiSettings): Promise<void>
  /// ai-settings.json was rewritten by any renderer; re-read it
  onAiSettingsChanged(handler: () => void): () => void
  /// shell only: switch to Home and open Settings › AI Model (rejects in standalone)
  openAiModelSettings(): Promise<void>
  aiChat(request: AiChatRequest): Promise<AiChatResponse>
  /// start a streaming AI call; deltas arrive via onAiStream with the same requestId
  aiStream(request: AiStreamRequest): Promise<void>
  aiStreamCancel(requestId: string): Promise<void>
  /// Genspark account status (gsk login state); withEmail also returns the email
  /// (needs a network request, slower)
  aiGskStatus(withEmail?: boolean): Promise<GenSparkAccountStatus>
  /// Opens the browser to sign in to Genspark (fire-and-forget; aiGskStatus
  /// becomes signed-in on completion)
  aiGskLogin(): Promise<void>
  /// Web search (main-process Serper/DuckDuckGo, shared with docs/slides)
  webSearch(query: string, maxResults?: number): Promise<WebSearchResult>
  /// Image search (same shared main-process channel as docs/slides)
  imageSearch(query: string, maxResults?: number): Promise<ImageSearchResponse>
  /// AI image generation via the Genspark account (sheets-owned channel)
  generateImage(op: { prompt: string; aspectRatio?: string }): Promise<GenerateImageResult>
  /// Downloads an image URL in the main process (SSRF-guarded); null on failure
  fetchImage(url: string): Promise<{ base64: string; mime: string } | null>
  onAiStream(handler: (chunk: AiStreamChunk) => void): () => void
  /// Chat attachments: multi-select file dialog (returns null on cancel)
  pickAttachments(): Promise<AttachmentAddResult | null>
  /// Validates dropped paths and returns attachment metadata
  addAttachmentPaths(paths: string[]): Promise<AttachmentAddResult>
  /// Persists a clipboard-pasted image (no local path) to a temp file and adds it
  /// as an attachment
  addPastedImage(data: ArrayBuffer, ext: string): Promise<AttachmentAddResult>
  /// Reads one chunk of an attachment's extracted text
  readAttachment(path: string, offset: number, maxChars: number): Promise<AttachmentReadResult>
  /// Reads an image attachment as base64 for multimodal input (≤5MB)
  readAttachmentImage(path: string): Promise<AttachmentImageResult>
  /// Absolute path of a File dropped onto the window (Electron webUtils)
  getPathForFile(file: File): string
}

export type MenuAction =
  'open' | 'save' | 'save-as' | 'print' | 'export-pdf' | 'export-csv' | 'undo' | 'redo'

export interface WebSearchResult {
  results: Array<{ title: string; url: string; snippet: string }>
  answer?: string
  method: string
  /** failure reason when method === 'error' */
  error?: string
}

export interface ImageSearchResponse {
  images: Array<{
    title: string
    imageUrl: string
    sourceUrl: string
    source: string
    width?: number
    height?: number
  }>
  method: string
  /** failure reason when method === 'error' */
  error?: string
}

export interface GenerateImageResult {
  url?: string
  error?: string
}
