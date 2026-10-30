import { it, expect } from 'vitest'
import JSZip from 'jszip'
import { buildDocx } from '../../docx-engine/tests/helpers/build-docx'
import { DocxSession } from '../src/session'

const apply = (session: DocxSession, settings: unknown) =>
  session.applySettings({
    transactionId: crypto.randomUUID(),
    expectedRevision: session.descriptor().revision,
    settings,
  })

it('applies Layout and Design settings as history entries and saves them to the package', async () => {
  const session = await DocxSession.open(
    await buildDocx({ bodyXml: '<w:p><w:r><w:t>Settings</w:t></w:r></w:p>' }),
  )
  const section = session.presentation().settings!.section!
  expect(section.pageWidth).toBeGreaterThan(0)
  const content = session.content()
  const narrow = {
    ...section,
    marginTop: 720,
    marginRight: 720,
    marginBottom: 720,
    marginLeft: 720,
  }
  const commit = apply(session, { section: narrow })
  expect(commit.steps).toEqual([])
  expect(commit.presentation!.settings!.section!.marginLeft).toBe(720)
  apply(session, { pageColor: 'FFF2CC', watermark: 'DRAFT', mirrorMargins: true })
  expect(session.content()).toEqual(content)
  session.history('undo')
  expect(session.presentation().settings!.pageColor ?? null).toBe(null)
  session.history('redo')
  const zip = await JSZip.loadAsync(await session.save())
  const documentXml = await zip.file('word/document.xml')!.async('string')
  expect(documentXml).toMatch(/<w:pgMar\b[^>]*w:left="720"/)
  expect(documentXml).toMatch(/<w:background\b[^>]*w:color="FFF2CC"/)
  const settingsXml = await zip.file('word/settings.xml')!.async('string')
  expect(settingsXml).toContain('<w:mirrorMargins')
  const headers = Object.keys(zip.files).filter((name) => /^word\/header\d+\.xml$/.test(name))
  const headerXml = await Promise.all(headers.map((name) => zip.file(name)!.async('string')))
  expect(headerXml.some((xml) => xml.includes('DRAFT'))).toBe(true)
  const reopened = await DocxSession.open(await session.save())
  expect(reopened.presentation().settings).toMatchObject({
    pageColor: 'FFF2CC',
    watermark: 'DRAFT',
    mirrorMargins: true,
  })
  expect(reopened.presentation().settings!.section!.marginLeft).toBe(720)
})

it('refuses settings that could carry markup or exceed page bounds', async () => {
  const session = await DocxSession.open(
    await buildDocx({ bodyXml: '<w:p><w:r><w:t>Settings</w:t></w:r></w:p>' }),
  )
  const section = session.presentation().settings!.section!
  const revision = session.descriptor().revision
  for (const settings of [
    { section: { ...section, vAlign: '"/><w:evil/>' } },
    { section: { ...section, injected: 1 } },
    { section: { ...section, pageWidth: 0 } },
    { pageColor: 'red' },
    { watermark: '<w:evil/>' },
    { unknown: true },
  ])
    expect(() => apply(session, settings), JSON.stringify(settings).slice(0, 80)).toThrow()
  expect(session.descriptor().revision).toBe(revision)
})

it('saves header and footer edits as regenerated parts and refuses forged run XML', async () => {
  const session = await DocxSession.open(
    await buildDocx({ bodyXml: '<w:p><w:r><w:t>Body</w:t></w:r></w:p>' }),
  )
  apply(session, {
    header: { text: 'Quarterly report' },
    footer: { text: 'Page #', pageNumber: true },
  })
  const zip = await JSZip.loadAsync(await session.save())
  const parts = await Promise.all(
    Object.keys(zip.files)
      .filter((name) => /^word\/(header|footer)\d+\.xml$/.test(name))
      .map(async (name) => [name, await zip.file(name)!.async('string')] as const),
  )
  expect(
    parts.some(([name, xml]) => name.includes('header') && xml.includes('Quarterly report')),
  ).toBe(true)
  expect(parts.some(([name, xml]) => name.includes('footer') && /PAGE/.test(xml))).toBe(true)
  const reopened = await DocxSession.open(await session.save())
  expect(reopened.presentation().settings!.header!.text).toContain('Quarterly report')
  const revision = session.descriptor().revision
  expect(() =>
    apply(session, {
      header: {
        text: 'x',
        paras: [{ runs: [{ text: 'x', rawRPr: '<w:rPr><w:evil/></w:rPr>' }] }],
      },
    }),
  ).toThrow('Untrusted')
  expect(() =>
    apply(session, { header: { text: 'x', paras: [{ runs: [], align: '<b>' }] } }),
  ).toThrow()
  expect(session.descriptor().revision).toBe(revision)
})

it('saves Design theme fonts and colors and refuses markup in typefaces', async () => {
  const session = await DocxSession.open(
    await buildDocx({ bodyXml: '<w:p><w:r><w:t>Theme</w:t></w:r></w:p>' }),
  )
  apply(session, {
    themeFonts: { major: 'Georgia', minor: 'Verdana' },
    themeColors: { name: 'Custom', accent1: '1F4E79' },
  })
  expect(session.presentation().css).toContain('Verdana')
  const zip = await JSZip.loadAsync(await session.save())
  const theme = Object.keys(zip.files).find((name) => /^word\/theme\/theme\d+\.xml$/.test(name))!
  const xml = await zip.file(theme)!.async('string')
  expect(xml).toContain('typeface="Georgia"')
  expect(xml).toContain('1F4E79')
  expect(() => apply(session, { themeFonts: { major: 'A"/><evil/>', minor: 'B' } })).toThrow()
})
