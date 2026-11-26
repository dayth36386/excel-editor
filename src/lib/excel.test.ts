import ExcelJS from 'exceljs'
import { describe, expect, it } from 'vitest'
import { cellData, colName, compareSort, displayText, editText, formatWithNumFmt, parseInput, viewOf } from './excel'
import { deleteCol, deleteRow, insertCol, insertRow, snapshotCol, snapshotRow } from './sheetOps'
import { parseTSV, toTSVCell } from './tsv'

describe('tsv', () => {
  it('parses clipboard text from Excel, including quoted cells', () => {
    expect(parseTSV('a\tb\r\nc\td\r\n')).toEqual([
      ['a', 'b'],
      ['c', 'd'],
    ])
    expect(parseTSV('"line1\nline2"\t"say ""hi"""')).toEqual([['line1\nline2', 'say "hi"']])
    expect(toTSVCell('a\tb')).toBe('"a\tb"')
  })
})

describe('viewOf (Excel view)', () => {
  it('carries the cell formatting into the grid', () => {
    const wb = new ExcelJS.Workbook()
    const ws = wb.addWorksheet('S')
    const head = ws.getCell('A1')
    head.value = 'Name'
    head.font = { bold: true, color: { argb: 'FFFFFFFF' } }
    head.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF107C41' } }
    expect(viewOf(head)).toEqual({
      text: 'Name',
      style: 'font-weight:600;color:#FFFFFF;background-color:#107C41',
      align: undefined,
    })

    const num = ws.getCell('B2')
    num.value = 1234.5
    num.numFmt = '#,##0.00'
    expect(viewOf(num)).toEqual({ text: '1,234.50', style: undefined, align: 'right' })

    // A filled cell without a font colour keeps Excel's black text (readable in dark mode too).
    const filled = ws.getCell('C3')
    filled.value = 'x'
    filled.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFF00' } }
    expect(viewOf(filled)?.style).toBe('color:#000000;background-color:#FFFF00')

    expect(viewOf(ws.getCell('Z9'))).toBeUndefined()
  })
})

/** Write a workbook to xlsx bytes and read it back, as the app does on save + reopen. */
async function roundTrip(wb: ExcelJS.Workbook) {
  const buf = await wb.xlsx.writeBuffer()
  const back = new ExcelJS.Workbook()
  await back.xlsx.load(buf)
  return back
}

function sampleBook() {
  const wb = new ExcelJS.Workbook()
  const ws = wb.addWorksheet('Data')
  ws.columns = [
    { header: 'Name', width: 20 },
    { header: 'Phone', width: 14 },
    { header: 'Points', width: 10, style: { numFmt: '#,##0' } },
  ]
  ws.getRow(1).font = { bold: true, color: { argb: 'FFFFFFFF' } }
  ws.getRow(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF107C41' } }
  ws.addRow(['Ann', '0812345678', 100])
  ws.addRow(['Bob', '0899999999', 250])
  ws.getCell('C4').value = { formula: 'SUM(C2:C3)', result: 350 }
  return wb
}

describe('colName', () => {
  it('converts column numbers to letters', () => {
    expect([1, 26, 27, 52, 703].map(colName)).toEqual(['A', 'Z', 'AA', 'AZ', 'AAA'])
  })
})

describe('parseInput', () => {
  it('converts numbers but keeps leading-zero codes as text', () => {
    expect(parseInput('42')).toBe(42)
    expect(parseInput('-3.5')).toBe(-3.5)
    expect(parseInput('1,234.50')).toBe(1234.5)
    expect(parseInput('0')).toBe(0)
    expect(parseInput('0.25')).toBe(0.25)
    expect(parseInput('0812345678')).toBe('0812345678')
    expect(parseInput('+66812345678')).toBe('+66812345678')
  })

  it('handles formulas, booleans, blanks and forced text', () => {
    expect(parseInput('=SUM(A1:A3)')).toEqual({ formula: 'SUM(A1:A3)' })
    expect(parseInput('true')).toBe(true)
    expect(parseInput('')).toBeNull()
    expect(parseInput("'123")).toBe('123')
    expect(parseInput('123', null, '@')).toBe('123')
  })

  it('parses ISO dates and rejects impossible ones', () => {
    expect(parseInput('2026-11-24')).toEqual(new Date(Date.UTC(2026, 10, 24)))
    expect(parseInput('2026-11-24 09:30')).toEqual(new Date(Date.UTC(2026, 10, 24, 9, 30)))
    expect(parseInput('2026-02-30')).toBe('2026-02-30')
  })

  it('keeps the link when a hyperlink cell is retyped', () => {
    const prev = { text: 'old', hyperlink: 'https://example.com' }
    expect(parseInput('new', prev)).toEqual({ text: 'new', hyperlink: 'https://example.com' })
  })
})

describe('displayText', () => {
  it('formats values for the grid', () => {
    expect(displayText(0.1 + 0.2)).toBe('0.3')
    expect(displayText(new Date(Date.UTC(2025, 0, 5)))).toBe('2025-01-05')
    expect(displayText({ formula: 'A1*2', result: 8 } as ExcelJS.CellValue)).toBe('8')
    expect(displayText({ richText: [{ text: 'a' }, { text: 'b' }] })).toBe('ab')
    expect(displayText(true)).toBe('TRUE')
  })
})

describe('formatWithNumFmt', () => {
  it('applies common Excel number formats', () => {
    expect(formatWithNumFmt(1234567.891, '#,##0')).toBe('1,234,568')
    expect(formatWithNumFmt(21102.3, '#,##0.00')).toBe('21,102.30')
    expect(formatWithNumFmt(0.256, '0.0%')).toBe('25.6%')
    expect(formatWithNumFmt(-1500, '"฿"#,##0.00')).toBe('-฿1,500.00')
    expect(formatWithNumFmt(42, '00000')).toBe('00042')
    expect(formatWithNumFmt(3.5, 'General')).toBe('3.5')
    expect(formatWithNumFmt(3.5, '0.00E+00')).toBe('3.5')
  })

  it('scales by 1000 for each trailing comma', () => {
    expect(formatWithNumFmt(166000, '#,##0,"k"')).toBe('166k')
    expect(formatWithNumFmt(1234567, '#,##0,"k"')).toBe('1,235k')
    expect(formatWithNumFmt(2500000, '0.0,,"M"')).toBe('2.5M')
  })
})

describe('compareSort', () => {
  it('puts numbers before text and sorts each naturally', () => {
    expect([10, 'ข', 2, 'ก'].sort(compareSort)).toEqual([2, 10, 'ก', 'ข'])
  })
})

describe('editing and saving', () => {
  it('saves an edit and keeps formatting, other cells and formulas', async () => {
    const wb = await roundTrip(sampleBook())
    const ws = wb.getWorksheet('Data')!

    const cell = ws.getCell('C2')
    cell.value = parseInput('999', cell.value, cell.numFmt)
    ws.getCell('A1').value = parseInput('Full name', ws.getCell('A1').value)

    const back = (await roundTrip(wb)).getWorksheet('Data')!
    expect(back.getCell('C2').value).toBe(999)
    expect(back.getCell('C2').numFmt).toBe('#,##0')
    expect(back.getCell('A1').value).toBe('Full name')
    expect(back.getCell('A1').font?.bold).toBe(true)
    expect(back.getCell('A1').fill).toMatchObject({ fgColor: { argb: 'FF107C41' } })
    expect(back.getCell('B2').value).toBe('0812345678')
    expect(editText(back.getCell('C4'))).toBe('=SUM(C2:C3)')
    expect(back.getColumn(1).width).toBe(20)
  })

  it('writes typed dates with a date format so Excel shows a date', async () => {
    const wb = sampleBook()
    const ws = wb.getWorksheet('Data')!
    ws.getCell('D2').value = parseInput('2026-11-24')
    const back = (await roundTrip(wb)).getWorksheet('Data')!
    expect(back.getCell('D2').value).toEqual(new Date(Date.UTC(2026, 10, 24)))
    expect(back.getCell('D2').numFmt).toBeTruthy()
  })

  it('describes cells for the table', () => {
    const ws = sampleBook().getWorksheet('Data')!
    expect(cellData(ws.getCell('A2'))).toEqual({ text: 'Ann', sort: 'ann', kind: 'text', formula: false })
    expect(cellData(ws.getCell('C3'))).toEqual({ text: '250', sort: 250, kind: 'number', formula: false })
    expect(cellData(ws.getCell('C4'))).toMatchObject({ text: '350', sort: 350, kind: 'number', formula: true })
    expect(cellData(ws.getCell('Z99'))).toBeUndefined()
  })
})

describe('rows and columns', () => {
  it('inserts a row styled like its neighbour, and undoes a delete', async () => {
    const wb = sampleBook()
    const ws = wb.getWorksheet('Data')!
    ws.getCell('A2').font = { italic: true }

    insertRow(ws, 2)
    expect(ws.getCell('A2').value).toBeNull()
    expect(ws.getCell('A2').font?.italic).toBe(true)
    expect(ws.getCell('A3').value).toBe('Ann')

    deleteRow(ws, 2)
    expect(ws.getCell('A2').value).toBe('Ann')

    insertRow(ws, 4, undefined, 'above')
    expect(ws.getCell('A4').value).toBeNull()
    expect(ws.getCell('C4').numFmt).toBe('#,##0')
    deleteRow(ws, 4)

    const snap = snapshotRow(ws, 2)
    deleteRow(ws, 2)
    expect(ws.getCell('A2').value).toBe('Bob')
    insertRow(ws, 2, snap)
    expect(ws.getCell('A2').value).toBe('Ann')
    expect(ws.getCell('A2').font?.italic).toBe(true)
    expect(ws.getCell('A3').value).toBe('Bob')

    const back = (await roundTrip(wb)).getWorksheet('Data')!
    expect(back.getRow(2).values).toEqual([, 'Ann', '0812345678', 100])
  })

  it('inserts and deletes columns, moving widths with them', () => {
    const ws = sampleBook().getWorksheet('Data')!
    insertCol(ws, 2)
    expect(ws.getCell('B2').value).toBeNull()
    expect(ws.getCell('C2').value).toBe('0812345678')
    expect(ws.getColumn(3).width).toBe(14)
    expect(ws.getColumn(4).width).toBe(10)

    const snap = snapshotCol(ws, 3)
    deleteCol(ws, 3)
    expect(ws.getCell('C2').value).toBe(100)
    insertCol(ws, 3, snap)
    expect(ws.getCell('C2').value).toBe('0812345678')
    expect(ws.getColumn(3).width).toBe(14)
  })
})
