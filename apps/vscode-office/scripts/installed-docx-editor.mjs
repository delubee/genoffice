/** Operate only the isolated installed-package fixture's actual webview. */
export async function editInstalledDocx(browser, text) {
  const root = await browser.newBrowserCDPSession()
  const finder = `const find = (doc) => {
    const editor = doc.querySelector('.tiptap[contenteditable="true"]');
    if (editor) return editor;
    for (const frame of doc.querySelectorAll('iframe')) {
      try { const found = frame.contentDocument && find(frame.contentDocument); if (found) return found; } catch {}
    }
    return null;
  };`
  try {
    const { targetInfos } = await root.send('Target.getTargets')
    for (const target of targetInfos) {
      if (target.type !== 'iframe' || !target.url.startsWith('vscode-webview://')) continue
      const { sessionId } = await root.send('Target.attachToTarget', {
        targetId: target.targetId,
        flatten: false,
      })
      let sequence = 0
      const pending = new Map()
      const receive = (event) => {
        if (event.sessionId !== sessionId) return
        const response = JSON.parse(event.message),
          item = pending.get(response.id)
        if (!item) return
        clearTimeout(item.timer)
        pending.delete(response.id)
        if (response.error) item.reject(new Error(response.error.message))
        else item.resolve(response.result)
      }
      root.on('Target.receivedMessageFromTarget', receive)
      const send = (method, params = {}) =>
        new Promise((resolve, reject) => {
          const id = ++sequence
          const timer = setTimeout(() => {
            pending.delete(id)
            reject(new Error(`Fixture editor command timed out: ${method}`))
          }, 5000)
          pending.set(id, { resolve, reject, timer })
          void root
            .send('Target.sendMessageToTarget', {
              sessionId,
              message: JSON.stringify({ id, method, params }),
            })
            .catch((error) => {
              clearTimeout(timer)
              pending.delete(id)
              reject(error)
            })
        })
      try {
        const found = await send('Runtime.evaluate', {
          expression: `(() => { ${finder} return !!find(document) })()`,
          returnByValue: true,
        })
        if (!found.result?.value) continue
        await send('Runtime.evaluate', {
          expression: `(() => { ${finder}
          const editor = find(document); editor.focus();
          const range = editor.ownerDocument.createRange(); range.selectNodeContents(editor);
          const selection = editor.ownerDocument.defaultView.getSelection(); selection.removeAllRanges(); selection.addRange(range);
        })()`,
          userGesture: true,
        })
        await send('Input.insertText', { text })
        const edited = await send('Runtime.evaluate', {
          expression: `(() => { ${finder} return find(document)?.textContent === ${JSON.stringify(text)} })()`,
          returnByValue: true,
        })
        if (edited.result?.value) return true
      } finally {
        root.off('Target.receivedMessageFromTarget', receive)
        for (const item of pending.values()) {
          clearTimeout(item.timer)
          item.reject(new Error('Fixture editor detached'))
        }
        await root.send('Target.detachFromTarget', { sessionId }).catch(() => {})
      }
    }
    return false
  } finally {
    await root.detach()
  }
}
