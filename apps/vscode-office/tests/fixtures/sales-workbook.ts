import JSZip from 'jszip'

export const SALES_SEED = 20261022
export const SALES_EXPECTED = {
  seed: SALES_SEED,
  rows: 10000,
  columns: 10,
  cellCount: 100000,
  dataRows: 9999,
  manualEdit: { sheet: 'Data', address: 'D2', original: 100, edited: 120, region: 'North' },
  // Independently calculated with integer cents; not read back from the XLSX.
  originalCents: { North: 12509014, South: 12502500, East: 12505000, West: 12496551 },
  editedCents: { North: 12511014, South: 12502500, East: 12505000, West: 12496551 },
  regionRows: { North: 2500, South: 2500, East: 2500, West: 2499 },
}
const ns = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
const rel = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
const escape = (value: string) =>
  value.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;')
const text = (address: string, value: string, style = 0) =>
  `<c r="${address}" t="inlineStr" s="${style}"><is><t xml:space="preserve">${escape(value)}</t></is></c>`
const number = (address: string, value: number, style = 0) =>
  `<c r="${address}" s="${style}"><v>${value}</v></c>`
const worksheet = (dimension: string, rows: string) =>
  `<?xml version="1.0" encoding="UTF-8"?><worksheet xmlns="${ns}" xmlns:r="${rel}"><dimension ref="${dimension}"/><sheetViews><sheetView workbookViewId="0"/></sheetViews><sheetFormatPr defaultRowHeight="15"/><sheetData>${rows}</sheetData></worksheet>`

async function pack(sheets: { name: string; xml: string }[], date1904 = false) {
  const zip = new JSZip()
  zip.file(
    '[Content_Types].xml',
    `<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>${sheets.map((_, index) => `<Override PartName="/xl/worksheets/sheet${index + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`).join('')}</Types>`,
  )
  zip.file(
    '_rels/.rels',
    `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${rel}/officeDocument" Target="xl/workbook.xml"/></Relationships>`,
  )
  zip.file(
    'xl/workbook.xml',
    `<workbook xmlns="${ns}" xmlns:r="${rel}"><workbookPr date1904="${date1904 ? '1' : '0'}"/><bookViews><workbookView/></bookViews><sheets>${sheets.map((sheet, index) => `<sheet name="${escape(sheet.name)}" sheetId="${index + 1}" r:id="rId${index + 1}"/>`).join('')}</sheets><calcPr calcId="191029" fullCalcOnLoad="1"/></workbook>`,
  )
  zip.file(
    'xl/_rels/workbook.xml.rels',
    `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${sheets.map((_, index) => `<Relationship Id="rId${index + 1}" Type="${rel}/worksheet" Target="worksheets/sheet${index + 1}.xml"/>`).join('')}<Relationship Id="styles" Type="${rel}/styles" Target="styles.xml"/></Relationships>`,
  )
  zip.file(
    'xl/styles.xml',
    `<styleSheet xmlns="${ns}"><numFmts count="1"><numFmt numFmtId="164" formatCode="yyyy-mm-dd hh:mm"/></numFmts><fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="11"/><name val="Calibri"/></font></fonts><fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FFCCEEFF"/><bgColor indexed="64"/></patternFill></fill></fills><borders count="1"><border/></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="5"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0"/><xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/><xf numFmtId="9" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/><xf numFmtId="4" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>`,
  )
  sheets.forEach((sheet, index) => zip.file(`xl/worksheets/sheet${index + 1}.xml`, sheet.xml))
  zip.file(
    'customXml/genoffice-sentinel.xml',
    '<preserve fixture="genoffice-v1">unknown part must survive</preserve>',
  )
  zip.forEach((_name, entry) => {
    entry.date = new Date('2000-01-01T00:00:00Z')
  })
  return zip.generateAsync({
    type: 'nodebuffer',
    compression: 'DEFLATE',
    compressionOptions: { level: 6 },
  })
}

export async function salesWorkbook() {
  const headers = [
    'ID',
    'Region',
    'Product',
    'Sales',
    'Quantity',
    'Date',
    'Discount',
    'Active',
    'Notes',
    'Status',
  ]
  const rows = [
    `<row r="1">${headers.map((header, column) => text(`${String.fromCharCode(65 + column)}1`, header, 1)).join('')}</row>`,
  ]
  const regions = ['North', 'South', 'East', 'West']
  for (let index = 0; index < 9999; index++) {
    const row = index + 2,
      amount = index === 0 ? 100 : (((index * 37 + (SALES_SEED % 997)) % 10000) + 1) / 100
    rows.push(
      `<row r="${row}">${text(`A${row}`, index === 0 ? '00123' : String(index + 1).padStart(8, '0'))}${text(`B${row}`, regions[index % 4]!)}${text(`C${row}`, `Widget ${index % 7}`)}${number(`D${row}`, amount, 4)}${number(`E${row}`, (index % 5) + 1)}${number(`F${row}`, 45000 + (index % 365) + 0.25, 2)}${number(`G${row}`, 0.05 * (index % 4), 3)}<c r="H${row}" t="b"><v>${index % 2 === 0 ? '1' : '0'}</v></c>${text(`I${row}`, index % 10 ? `Item ${index}` : '')}${text(`J${row}`, index % 11 ? 'open' : '=literal')}</row>`,
    )
  }
  return pack([
    { name: 'Data', xml: worksheet('A1:J10000', rows.join('')) },
    {
      name: 'Config',
      xml: worksheet(
        'A1:B3',
        `<row r="1">${text('A1', 'Seed', 1)}${number('B1', SALES_SEED)}</row><row r="2">${text('A2', 'Discount')}${number('B2', 0.1, 3)}</row><row r="3">${text('A3', 'Double first sale')}<c r="B3"><f>Data!D2*2</f><v>200</v></c></row>`,
      ),
    },
  ])
}

export async function typesWorkbook(date1904 = false) {
  const headers = [
    'Text ID',
    'Empty text',
    'Blank',
    'Serial 60',
    'Time',
    'Percent',
    'Boolean',
    'Literal formula',
    'Error',
    'Ready formula',
    'Pending formula',
    'Error formula',
  ]
  const row = `${text('A2', '00123')}${text('B2', '')}<c r="C2"/>${number('D2', 60, 2)}${number('E2', 60.5, 2)}${number('F2', 0.25, 3)}<c r="G2" t="b"><v>1</v></c>${text('H2', '=1+2')}<c r="I2" t="e"><v>#REF!</v></c><c r="J2"><f>1+2</f><v>3</v></c><c r="K2"><f>NOW()</f></c><c r="L2" t="e"><f>1/0</f><v>#DIV/0!</v></c>`
  return pack(
    [
      {
        name: 'Types',
        xml: worksheet(
          'A1:L2',
          `<row r="1">${headers.map((header, column) => text(`${String.fromCharCode(65 + column)}1`, header, 1)).join('')}</row><row r="2">${row}</row>`,
        ),
      },
    ],
    date1904,
  )
}

export async function preservationWorkbook() {
  const source = await pack([
    {
      name: 'Preserve',
      xml: worksheet(
        'A1:D4',
        `<row r="1">${text('A1', 'Protected original chart', 1)}</row><row r="2">${text('A2', 'North')}${number('B2', 10)}</row><row r="3">${text('A3', 'South')}${number('B3', 20)}</row>`,
      ),
    },
    { name: 'Editable', xml: worksheet('A1:A1', `<row r="1">${number('A1', 7)}</row>`) },
  ])
  const zip = await JSZip.loadAsync(source)
  const first = await zip.file('xl/worksheets/sheet1.xml')!.async('string')
  zip.file(
    'xl/worksheets/sheet1.xml',
    first.replace(
      '</worksheet>',
      '<sheetProtection sheet="1"/><protectedRanges><protectedRange name="OriginalRange" sqref="D1:D4"/></protectedRanges><mergeCells count="1"><mergeCell ref="A1:B1"/></mergeCells><drawing r:id="drawing1"/></worksheet>',
    ),
  )
  zip.file(
    'xl/worksheets/_rels/sheet1.xml.rels',
    `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="drawing1" Type="${rel}/drawing" Target="../drawings/drawing1.xml"/></Relationships>`,
  )
  zip.file(
    'xl/drawings/drawing1.xml',
    `<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><xdr:twoCellAnchor><xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>5</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>6</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>18</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to><xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="2" name="Original chart"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="${rel}" r:id="chart1"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor></xdr:wsDr>`,
  )
  zip.file(
    'xl/drawings/_rels/drawing1.xml.rels',
    `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="chart1" Type="${rel}/chart" Target="../charts/chart1.xml"/></Relationships>`,
  )
  zip.file(
    'xl/charts/chart1.xml',
    `<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart><c:plotArea><c:layout/><c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:ser><c:idx val="0"/><c:order val="0"/><c:tx><c:v>Sales</c:v></c:tx><c:cat><c:strRef><c:f>Preserve!$A$2:$A$3</c:f><c:strCache><c:ptCount val="2"/><c:pt idx="0"><c:v>North</c:v></c:pt><c:pt idx="1"><c:v>South</c:v></c:pt></c:strCache></c:strRef></c:cat><c:val><c:numRef><c:f>Preserve!$B$2:$B$3</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="2"/><c:pt idx="0"><c:v>10</c:v></c:pt><c:pt idx="1"><c:v>20</c:v></c:pt></c:numCache></c:numRef></c:val></c:ser><c:axId val="1"/><c:axId val="2"/></c:barChart><c:catAx><c:axId val="1"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="b"/><c:crossAx val="2"/></c:catAx><c:valAx><c:axId val="2"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="l"/><c:crossAx val="1"/></c:valAx></c:plotArea><c:plotVisOnly val="1"/></c:chart></c:chartSpace>`,
  )
  const contentTypes = await zip.file('[Content_Types].xml')!.async('string')
  zip.file(
    '[Content_Types].xml',
    contentTypes.replace(
      '</Types>',
      '<Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/><Override PartName="/xl/charts/chart1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/></Types>',
    ),
  )
  zip.forEach((_name, entry) => {
    entry.date = new Date('2000-01-01T00:00:00Z')
  })
  return zip.generateAsync({
    type: 'nodebuffer',
    compression: 'DEFLATE',
    compressionOptions: { level: 6 },
  })
}
