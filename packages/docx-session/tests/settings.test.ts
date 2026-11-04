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

it('saves Different First Page and Different Odd & Even Pages', async () => {
  const session = await DocxSession.open(
    await buildDocx({ bodyXml: '<w:p><w:r><w:t>Pages</w:t></w:r></w:p>' }),
  )
  expect(session.presentation().settings).toMatchObject({
    titlePg: false,
    evenAndOddHeaders: false,
  })
  apply(session, { titlePg: true, evenAndOddHeaders: true })
  const zip = await JSZip.loadAsync(await session.save())
  expect(await zip.file('word/document.xml')!.async('string')).toContain('<w:titlePg')
  expect(await zip.file('word/settings.xml')!.async('string')).toContain('<w:evenAndOddHeaders')
  const reopened = await DocxSession.open(await session.save())
  expect(reopened.presentation().settings).toMatchObject({ titlePg: true, evenAndOddHeaders: true })
})

it('saves the Page Number Format of the final section', async () => {
  const session = await DocxSession.open(
    await buildDocx({ bodyXml: '<w:p><w:r><w:t>Pages</w:t></w:r></w:p>' }),
  )
  apply(session, { pageNumbering: { fmt: 'lowerRoman', start: 3 } })
  const zip = await JSZip.loadAsync(await session.save())
  expect(await zip.file('word/document.xml')!.async('string')).toMatch(
    /<w:pgNumType\b[^>]*w:fmt="lowerRoman"[^>]*w:start="3"|<w:pgNumType\b[^>]*w:start="3"[^>]*w:fmt="lowerRoman"/,
  )
  const reopened = await DocxSession.open(await session.save())
  expect(reopened.presentation().settings!.pageNumbering).toEqual({ fmt: 'lowerRoman', start: 3 })
  expect(() => apply(session, { pageNumbering: { fmt: '"/><w:evil/>' } })).toThrow()
})

it('saves first-page and even-page header and footer parts', async () => {
  const session = await DocxSession.open(
    await buildDocx({ bodyXml: '<w:p><w:r><w:t>Pages</w:t></w:r></w:p>' }),
  )
  apply(session, {
    titlePg: true,
    evenAndOddHeaders: true,
    headerFirst: { text: 'Cover header' },
    footerEven: { text: 'Even footer' },
  })
  const zip = await JSZip.loadAsync(await session.save())
  const parts = await Promise.all(
    Object.keys(zip.files)
      .filter((name) => /^word\/(header|footer)\d+\.xml$/.test(name))
      .map((name) => zip.file(name)!.async('string')),
  )
  expect(parts.some((xml) => xml.includes('Cover header'))).toBe(true)
  expect(parts.some((xml) => xml.includes('Even footer'))).toBe(true)
  expect(await zip.file('word/document.xml')!.async('string')).toMatch(
    /<w:headerReference\b[^>]*w:type="first"/,
  )
  const reopened = await DocxSession.open(await session.save())
  expect(reopened.presentation().settings).toMatchObject({
    headerFirst: { text: 'Cover header' },
    footerEven: { text: 'Even footer' },
  })
})

it('edits the page setup of a non-final section and saves it into that section', async () => {
  const pgMar =
    '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/>'
  const session = await DocxSession.open(
    await buildDocx({
      bodyXml:
        '<w:p><w:r><w:t>One</w:t></w:r></w:p>' +
        `<w:p><w:pPr><w:sectPr><w:pgSz w:w="12240" w:h="15840"/>${pgMar}</w:sectPr></w:pPr></w:p>` +
        '<w:p><w:r><w:t>Two</w:t></w:r></w:p>',
    }),
  )
  const sections = session.presentation().settings!.sections!
  expect(sections).toHaveLength(2)
  expect(sections.map((entry) => entry.final)).toEqual([false, true])
  const first = sections[0]!
  apply(session, {
    sectionEdits: {
      [String(first.blockIndex)]: {
        section: {
          ...first.section,
          orientation: 'landscape',
          pageWidth: 15840,
          pageHeight: 12240,
        },
        startType: 'continuous',
        pageNumbering: { fmt: 'lowerRoman', start: 3 },
        titlePg: true,
      },
    },
  })
  expect(session.presentation().settings!.sections![0]).toMatchObject({
    section: { orientation: 'landscape', pageWidth: 15840 },
    startType: 'continuous',
    titlePg: true,
    pageNumbering: { fmt: 'lowerRoman', start: 3 },
  })
  const reopened = await DocxSession.open(await session.save())
  const saved = reopened.presentation().settings!.sections!
  expect(saved).toHaveLength(2)
  expect(saved[0]).toMatchObject({
    section: { orientation: 'landscape', pageWidth: 15840, pageHeight: 12240 },
    startType: 'continuous',
    titlePg: true,
    pageNumbering: { fmt: 'lowerRoman', start: 3 },
  })
  // The final section keeps the file's own geometry.
  expect(saved[1]!.section.pageWidth).toBe(sections[1]!.section.pageWidth)
  expect(() => apply(session, { sectionEdits: { '999': { titlePg: true } } })).toThrow(
    'Unknown section',
  )
  expect(() =>
    apply(session, { sectionEdits: { [String(sections[1]!.blockIndex)]: { titlePg: true } } }),
  ).toThrow('Unknown section')
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

it('saves Protect Document settings and reopens the document read-only', async () => {
  const session = await DocxSession.open(
    await buildDocx({ bodyXml: '<w:p><w:r><w:t>Final</w:t></w:r></w:p>' }),
  )
  apply(session, { protection: { edit: 'readOnly', enforced: true }, removePersonalInfo: true })
  const saved = await session.save()
  const zip = await JSZip.loadAsync(saved)
  const settingsXml = await zip.file('word/settings.xml')!.async('string')
  expect(settingsXml).toMatch(/<w:documentProtection\b[^>]*w:edit="readOnly"/)
  expect(settingsXml).toContain('<w:removePersonalInformation')
  const reopened = await DocxSession.open(saved)
  expect(reopened.readonly).toBe(true)
  expect(() =>
    apply(session, { protection: { edit: 'readOnly', enforced: true, hash: '<x/>' } }),
  ).toThrow()
})

it('saves References sources for citations and bibliographies', async () => {
  const session = await DocxSession.open(
    await buildDocx({ bodyXml: '<w:p><w:r><w:t>Cited</w:t></w:r></w:p>' }),
  )
  const source = {
    tag: 'Wang2024',
    type: 'Book',
    author: 'Wang Li',
    title: 'Field Notes',
    year: '2024',
  }
  apply(session, { sources: [source] })
  const reopened = await DocxSession.open(await session.save())
  expect(reopened.presentation().settings!.sources).toEqual([expect.objectContaining(source)])
  expect(() => apply(session, { sources: [{ ...source, tag: '<x/>' }] })).toThrow()
})

it('saves drawn ink as floating pictures from the view raster and refuses unknown anchors', async () => {
  const session = await DocxSession.open(
    await buildDocx({ bodyXml: '<w:p><w:r><w:t>Sketch</w:t></w:r></w:p>' }),
  )
  const png =
    'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg=='
  const ink = {
    id: 'ink1-a',
    anchorIndex: 0,
    stroke: {
      tool: 'pen',
      color: 'C00000',
      width: 2,
      points: [
        { x: 1, y: 1 },
        { x: 20, y: 10 },
      ],
    },
    raster: { base64: png, widthPx: 21, heightPx: 11, offsetXPx: 0, offsetYPx: 0 },
  }
  expect(() => apply(session, { inks: [{ ...ink, anchorIndex: 99 }] })).toThrow(
    'Unknown ink anchor',
  )
  apply(session, { inks: [ink] })
  const zip = await JSZip.loadAsync(await session.save())
  expect(Object.keys(zip.files).some((name) => /^word\/media\/.+\.png$/.test(name))).toBe(true)
  const reopened = await DocxSession.open(await session.save())
  expect(reopened.presentation().settings!.inks).toHaveLength(1)
})
