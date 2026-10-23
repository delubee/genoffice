import * as vscode from 'vscode'

interface ProbeDocument extends vscode.CustomDocument {
  value: number
}
export function activate(context: vscode.ExtensionContext) {
  const events = new vscode.EventEmitter<vscode.CustomDocumentEditEvent<ProbeDocument>>()
  let document: ProbeDocument | undefined
  let saves = 0,
    undos = 0,
    redos = 0,
    edits = 0,
    preparations = 0,
    invokes = 0
  const write = async (doc: ProbeDocument, target = doc.uri) => {
    await vscode.workspace.fs.writeFile(target, new TextEncoder().encode(String(doc.value)))
  }
  const provider: vscode.CustomEditorProvider<ProbeDocument> = {
    onDidChangeCustomDocument: events.event,
    async openCustomDocument(uri) {
      const value = Number(new TextDecoder().decode(await vscode.workspace.fs.readFile(uri)))
      document = {
        uri,
        value,
        dispose() {
          document = undefined
        },
      }
      return document
    },
    async resolveCustomEditor(_doc, panel) {
      panel.webview.html = '<html><body>Stable API probe</body></html>'
    },
    async saveCustomDocument(doc) {
      await write(doc)
      saves++
    },
    async saveCustomDocumentAs(doc, destination) {
      await write(doc, destination)
    },
    async revertCustomDocument(doc) {
      doc.value = Number(new TextDecoder().decode(await vscode.workspace.fs.readFile(doc.uri)))
    },
    async backupCustomDocument(doc, { destination }) {
      await write(doc, destination)
      return {
        id: destination.toString(),
        delete() {
          void vscode.workspace.fs.delete(destination)
        },
      }
    },
  }
  function edit(value: number) {
    if (!document) throw new Error('No test document open')
    const doc = document,
      before = doc.value
    doc.value = value
    edits++
    events.fire({
      document: doc,
      label: 'Probe edit',
      undo() {
        doc.value = before
        undos++
      },
      redo() {
        doc.value = value
        redos++
      },
    })
  }
  const tool: vscode.LanguageModelTool<{ expected: number; value: number }> = {
    prepareInvocation({ input }) {
      preparations++
      return { invocationMessage: `Probe ${document?.uri.toString()} revision ${input.expected}` }
    },
    invoke({ input }) {
      invokes++
      if (document?.value !== input.expected)
        throw new Error('Stale preparation: invoke revalidated current state')
      edit(input.value)
      return new vscode.LanguageModelToolResult([new vscode.LanguageModelTextPart('committed')])
    },
  }
  context.subscriptions.push(
    events,
    vscode.window.registerCustomEditorProvider('genoffice-test.probe', provider),
    vscode.lm.registerTool('genoffice_test_probe', tool),
  )
  return {
    edit,
    prepare: (input: { expected: number; value: number }) =>
      tool.prepareInvocation!({ input }, new vscode.CancellationTokenSource().token),
    inspect: () => ({ value: document?.value, saves, undos, redos, edits, preparations, invokes }),
  }
}
