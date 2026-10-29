import * as assert from 'node:assert/strict'
import * as vscode from 'vscode'
import { suite, test } from 'mocha'
import { chromium, type Frame } from '@playwright/test'
import JSZip from 'jszip'
import { manualWorkbook } from '../fixtures/manual-workbook'
import {
  planCellEditsToXlsx,
  createBufferEntrySource,
  assembleWithJsZip,
} from '@genoffice/xlsx-gateway/gateway/xlsx-gateway'
import type { WorkbookDocument } from '../../src/extension/workbook-document'

const pixel =
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII='
suite('Desktop Source Visuals UI', () => {
  test('edits native shapes and pictures through shared contextual ribbons and preserves media/history on save/reopen', async function () {
    const port = process.env.GENOFFICE_TEST_CDP_PORT
    if (!port) return this.skip()
    this.timeout(120_000)
    const api = (await vscode.extensions
      .getExtension('genoffice.genoffice-workbook')!
      .activate()) as { getDocument(uri: vscode.Uri): WorkbookDocument | undefined }
    const uri = vscode.Uri.joinPath(
      vscode.workspace.workspaceFolders![0]!.uri,
      'source-visuals-ui.xlsx',
    )
    const anchor = (fromColumn: number, fromRow: number, toColumn: number, toRow: number) => ({
      fromColumn,
      fromRow,
      toColumn,
      toRow,
      fromColumnOffset: 0,
      fromRowOffset: 0,
      toColumnOffset: 0,
      toRowOffset: 0,
    })
    const base = await manualWorkbook()
    const additions = [
      ...[
        [3, 3, 5, 6],
        [7, 3, 10, 7],
        [4, 9, 6, 12],
      ].map((box, i) => ({
        sheetName: 'Data',
        anchor: anchor(box[0]!, box[1]!, box[2]!, box[3]!),
        shape: {
          shapeType: 'rect',
          fillColor: '#3399FF',
          lineColor: '#222222',
          text: `Shape ${i}`,
          altText: `Shape ${i}`,
        },
      })),
      {
        sheetName: 'Data',
        anchor: anchor(9, 9, 12, 13),
        image: { mediaType: 'image/png' as const, base64: pixel, altText: 'Picture 0' },
      },
    ]
    const plan = await planCellEditsToXlsx(
      await createBufferEntrySource(base),
      [],
      [],
      [],
      undefined,
      [],
      [],
      [],
      [],
      [],
      null,
      additions,
    )
    const original = await assembleWithJsZip(base, plan)
    await vscode.workspace.fs.writeFile(uri, original.buffer)
    const browser = await chromium.connectOverCDP(`http://127.0.0.1:${port}`)
    const getFrame = async (): Promise<Frame> => {
      const started = Date.now()
      while (Date.now() - started < 20_000) {
        for (const context of browser.contexts())
          for (const page of context.pages()) {
            if (!(await page.title().then((title) => title.includes('source-visuals-ui.xlsx'))))
              continue
            for (const frame of page.frames())
              if (
                await frame
                  .locator('.workbook-desktop-header')
                  .count()
                  .catch(() => 0)
              )
                return frame
          }
        await new Promise((resolve) => setTimeout(resolve, 100))
      }
      throw new Error('Source visual workbook did not load')
    }
    try {
      await vscode.commands.executeCommand('vscode.openWith', uri, 'genoffice.workbook')
      let document = api.getDocument(uri)!
      await document.whenViewReady()
      let frame = await getFrame()
      const metadata = structuredClone(
        (document.engine.state.features!.originalMetadata as any).visuals,
      )
      const shapes = metadata.filter((v: any) => v.kind === 'shape'),
        picture = metadata.find((v: any) => v.kind === 'image')
      assert.equal(shapes.length, 3)
      assert.ok(picture)
      const object = (id: string) => frame.locator(`[data-print-visual="${id}"] .shape-editable`)
      await object(shapes[0].id).waitFor()
      const image = frame.locator(`[data-print-visual="${picture.id}"] img`)
      await image.waitFor()
      const mediaStart = Date.now()
      while (
        !(await image.evaluate((img: HTMLImageElement) => img.complete && img.naturalWidth > 0)) &&
        Date.now() - mediaStart < 10_000
      )
        await new Promise((resolve) => setTimeout(resolve, 50))
      assert.equal(await image.evaluate((img: HTMLImageElement) => img.naturalWidth), 1)
      await object(shapes[0].id).click({ position: { x: 15, y: 15 } })
      await frame.getByRole('tab', { name: 'Shape Format', exact: true }).click()
      let ribbon = frame.locator('#ribbon-panel-shapeFormat')
      const beforePaint = structuredClone(document.engine.state)
      await ribbon.getByRole('button', { name: 'Shape fill options', exact: true }).click()
      await frame
        .getByRole('grid', { name: 'Standard Colors', exact: true })
        .getByRole('gridcell', { name: 'Red', exact: true })
        .click()
      await document.bridge!.request('view.flush', {})
      const afterPaint = structuredClone(document.engine.state)
      assert.equal(
        (afterPaint.sheets[0]!.features!.visualEdits as any)[shapes[0].id].fillColor.toUpperCase(),
        '#FF0000',
      )
      await vscode.commands.executeCommand('undo')
      assert.deepEqual(document.engine.state, beforePaint)
      await vscode.commands.executeCommand('redo')
      assert.deepEqual(document.engine.state, afterPaint)
      await ribbon.getByRole('button', { name: 'Rotate', exact: true }).click()
      await ribbon.getByRole('option', { name: 'Rotate Right 90°', exact: true }).click()
      await document.bridge!.request('view.flush', {})
      await ribbon.getByRole('button', { name: 'Size and Properties…', exact: true }).click()
      const size = frame.getByRole('dialog', { name: 'Size and Properties', exact: true })
      await size.getByRole('spinbutton', { name: /Width/ }).fill('3.2')
      await size.getByRole('spinbutton', { name: /Height/ }).fill('2.1')
      await size.getByRole('spinbutton', { name: /Rotation/ }).fill('35')
      const beforeSize = structuredClone(document.engine.state),
        sizeRevision = document.engine.descriptor.revision
      await size.getByRole('button', { name: 'OK', exact: true }).click()
      await size.waitFor({ state: 'hidden' })
      await document.bridge!.request('view.flush', {})
      assert.equal(document.engine.descriptor.revision, sizeRevision + 1)
      const afterSize = structuredClone(document.engine.state)
      assert.deepEqual(
        (afterSize.sheets[0]!.features!.visualEdits as any)[shapes[0].id].frameSize,
        { width: 1152000, height: 756000 },
      )
      await vscode.commands.executeCommand('undo')
      assert.deepEqual(document.engine.state, beforeSize)
      await vscode.commands.executeCommand('redo')
      assert.deepEqual(document.engine.state, afterSize)
      await object(picture.id).click({ position: { x: 15, y: 15 } })
      await frame.getByRole('tab', { name: 'Picture Format', exact: true }).click()
      const pictureRibbon = frame.locator('#ribbon-panel-pictureFormat')
      await pictureRibbon.getByRole('button', { name: 'Rotate', exact: true }).click()
      await pictureRibbon.getByRole('option', { name: 'Flip Vertical', exact: true }).click()
      await document.bridge!.request('view.flush', {})
      await object(shapes[1].id).click({ position: { x: 15, y: 15 } })
      await object(shapes[2].id).click({ position: { x: 15, y: 15 }, modifiers: ['Shift'] })
      await frame.getByRole('tab', { name: 'Shape Format', exact: true }).click()
      ribbon = frame.locator('#ribbon-panel-shapeFormat')
      const beforeAlign = structuredClone(document.engine.state),
        alignRevision = document.engine.descriptor.revision
      await ribbon.getByRole('button', { name: 'Align', exact: true }).click()
      await ribbon.getByRole('option', { name: 'Align Top', exact: true }).click()
      await document.bridge!.request('view.flush', {})
      assert.equal(document.engine.descriptor.revision, alignRevision + 1)
      const afterAlign = structuredClone(document.engine.state)
      await vscode.commands.executeCommand('undo')
      assert.deepEqual(document.engine.state, beforeAlign)
      await vscode.commands.executeCommand('redo')
      assert.deepEqual(document.engine.state, afterAlign)
      await vscode.commands.executeCommand('workbench.action.files.save')
      const saved = await JSZip.loadAsync(await vscode.workspace.fs.readFile(uri)),
        oldZip = await JSZip.loadAsync(original.buffer)
      for (const path of Object.keys(oldZip.files).filter((path) => path.startsWith('xl/media/')))
        assert.deepEqual(
          await saved.file(path)!.async('uint8array'),
          await oldZip.file(path)!.async('uint8array'),
        )
      assert.equal(
        await saved.file('customXml/manual-sentinel.xml')!.async('string'),
        '<preserve id="manual-fixture">unchanged</preserve>',
      )
      const session = document.sessionId
      await vscode.commands.executeCommand('workbench.action.closeActiveEditor')
      const closed = Date.now()
      while (api.getDocument(uri) && Date.now() - closed < 10_000)
        await new Promise((resolve) => setTimeout(resolve, 50))
      assert.equal(api.getDocument(uri), undefined)
      await vscode.commands.executeCommand('vscode.openWith', uri, 'genoffice.workbook')
      document = api.getDocument(uri)!
      await document.whenViewReady()
      assert.notEqual(document.sessionId, session)
      const reopened = (document.engine.state.features!.originalMetadata as any).visuals
      const savedShape = reopened.find((v: any) => v.altText === 'Shape 0'),
        savedPicture = reopened.find((v: any) => v.altText === 'Picture 0')
      assert.equal(savedShape.fillColor.toUpperCase(), '#FF0000')
      assert.equal(savedShape.rotation, 35)
      assert.equal(savedShape.frameWidth, 1152000)
      assert.equal(savedShape.frameHeight, 756000)
      assert.equal(savedPicture.flipV, true)
      frame = await getFrame()
      await frame.locator(`[data-print-visual="${savedPicture.id}"] img`).waitFor()
      const beforeInsert = structuredClone(document.engine.state)
      await frame.getByRole('tab', { name: 'Insert', exact: true }).click()
      await frame
        .locator('.workbook-desktop-header')
        .getByRole('button', { name: 'Shapes', exact: true })
        .click()
      await frame
        .locator('.rb-shape-gallery')
        .getByRole('button', { name: 'Rectangle', exact: true })
        .click()
      const grid = frame.locator('.workbook-grid')
      const armed = Date.now()
      while (
        !(await grid.evaluate((el) => el.classList.contains('sheet-shape-drawing'))) &&
        Date.now() - armed < 5000
      )
        await new Promise((resolve) => setTimeout(resolve, 25))
      assert.equal(await grid.evaluate((el) => el.classList.contains('sheet-shape-drawing')), true)
      const surface = await grid.locator('canvas').evaluateAll(
        (nodes) =>
          nodes
            .map((node) => {
              const r = node.getBoundingClientRect()
              return { x: r.x, y: r.y, width: r.width, height: r.height }
            })
            .sort((a, b) => b.width * b.height - a.width * a.height)[0]!,
      )
      const ownerPage = frame.page()
      await ownerPage.mouse.move(surface.x + 550, surface.y + 200)
      await ownerPage.mouse.down()
      await ownerPage.mouse.move(surface.x + 680, surface.y + 285, { steps: 6 })
      await ownerPage.mouse.up()
      const addedFrame = frame.locator('[data-print-visual^="added-shape-"] .shape-editable')
      await addedFrame.waitFor()
      const inserted = structuredClone(document.engine.state)
      const insertedShape = (inserted.sheets[0]!.features!.visuals as any[]).find(
        (v) => v.kind === 'shape',
      )
      assert.ok(insertedShape)
      await vscode.commands.executeCommand('undo')
      assert.deepEqual(document.engine.state, beforeInsert)
      await vscode.commands.executeCommand('redo')
      assert.deepEqual(document.engine.state, inserted)
      await addedFrame.dblclick({ position: { x: 25, y: 25 } })
      const shapeText = addedFrame.locator('.shape-text-editor')
      await shapeText.fill('Native inserted shape')
      assert.equal(await shapeText.evaluate((el) => window.document.activeElement === el), true)
      await vscode.commands.executeCommand('workbench.action.files.save')
      assert.equal(
        (document.engine.state.sheets[0]!.features!.visuals as any[]).find(
          (v) => v.id === insertedShape.id,
        ).text,
        'Native inserted shape',
      )
      assert.equal(document.engine.descriptor.dirty, false)
      const insertedSession = document.sessionId
      await vscode.commands.executeCommand('workbench.action.closeActiveEditor')
      const insertClosed = Date.now()
      while (api.getDocument(uri) && Date.now() - insertClosed < 10_000)
        await new Promise((resolve) => setTimeout(resolve, 50))
      assert.equal(api.getDocument(uri), undefined)
      await vscode.commands.executeCommand('vscode.openWith', uri, 'genoffice.workbook')
      document = api.getDocument(uri)!
      await document.whenViewReady()
      assert.notEqual(document.sessionId, insertedSession)
      const newMetadata = (document.engine.state.features!.originalMetadata as any).visuals
      assert.equal(newMetadata.filter((v: any) => v.kind === 'shape').length, 4)
      const savedInserted = newMetadata.find((v: any) => v.text === 'Native inserted shape')
      assert.ok(savedInserted)
      for (const key of [
        'fromRow',
        'toRow',
        'fromColumn',
        'toColumn',
        'fromRowOffset',
        'toRowOffset',
        'fromColumnOffset',
        'toColumnOffset',
      ])
        assert.equal(savedInserted.anchor[key], insertedShape.anchor[key])
      frame = await getFrame()
      const beforePicture = structuredClone(document.engine.state)
      await frame.getByRole('tab', { name: 'Insert', exact: true }).click()
      const chooserPromise = frame.page().waitForEvent('filechooser')
      await frame
        .locator('.workbook-desktop-header')
        .getByRole('button', { name: 'Pictures', exact: true })
        .click()
      await (
        await chooserPromise
      ).setFiles({
        name: 'Native pixel.png',
        mimeType: 'image/png',
        buffer: Buffer.from(pixel, 'base64'),
      })
      const insertedImage = frame.locator('[data-print-visual^="added-image-"] img')
      await insertedImage.waitFor()
      await document.bridge!.request('view.flush', {})
      const afterPicture = structuredClone(document.engine.state)
      const newPicture = (afterPicture.sheets[0]!.features!.visuals as any[]).find(
        (v) => v.kind === 'image',
      )
      assert.ok(newPicture)
      assert.equal(await insertedImage.evaluate((el: HTMLImageElement) => el.naturalWidth), 1)
      await vscode.commands.executeCommand('undo')
      assert.deepEqual(document.engine.state, beforePicture)
      await vscode.commands.executeCommand('redo')
      assert.deepEqual(document.engine.state, afterPicture)
      await vscode.commands.executeCommand('workbench.action.files.save')
      assert.equal(document.engine.descriptor.dirty, false)
      const pictureZip = await JSZip.loadAsync(await vscode.workspace.fs.readFile(uri))
      const mediaBytes = await Promise.all(
        Object.keys(pictureZip.files)
          .filter((name) => name.startsWith('xl/media/') && !pictureZip.files[name]!.dir)
          .map((name) => pictureZip.file(name)!.async('base64')),
      )
      assert.ok(mediaBytes.includes(pixel))
      const pictureSession = document.sessionId
      await vscode.commands.executeCommand('workbench.action.closeActiveEditor')
      const pictureClosed = Date.now()
      while (api.getDocument(uri) && Date.now() - pictureClosed < 10_000)
        await new Promise((resolve) => setTimeout(resolve, 50))
      assert.equal(api.getDocument(uri), undefined)
      await vscode.commands.executeCommand('vscode.openWith', uri, 'genoffice.workbook')
      document = api.getDocument(uri)!
      await document.whenViewReady()
      assert.notEqual(document.sessionId, pictureSession)
      const pictures = (
        (document.engine.state.features!.originalMetadata as any).visuals as any[]
      ).filter((v) => v.kind === 'image')
      assert.equal(pictures.length, 2)
      assert.ok(
        pictures.some((v) =>
          Object.keys(newPicture.anchor).every((key) => v.anchor[key] === newPicture.anchor[key]),
        ),
      )
    } finally {
      if (api.getDocument(uri))
        await vscode.commands.executeCommand('workbench.action.files.save', uri)
      await vscode.commands.executeCommand('workbench.action.closeAllEditors')
      await vscode.workspace.fs.delete(uri)
      await browser.close()
    }
  })
})
