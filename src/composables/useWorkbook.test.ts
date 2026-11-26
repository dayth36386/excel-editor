import { readFileSync } from 'node:fs'
import ExcelJS from 'exceljs'
import { describe, expect, it } from 'vitest'
import { detectLayout } from '@/lib/layout'
import { normalizeXlsx } from '@/lib/normalizeXlsx'
import { useWorkbook } from './useWorkbook'

const toArrayBuffer = (buf: Uint8Array) =>
  buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength) as ArrayBuffer

// Scoreboard saved by openpyxl: merged title (A1:F1) and legend (A2:F2), header in row 4,
// labs in rows 5–13, a SUM total row 14 with no cached results, a merged footnote in row 16.
const scoreboard = () =>
  toArrayBuffer(readFileSync(new URL('../lib/__fixtures__/openpyxl-comments.xlsx', import.meta.url)))

describe('detectLayout', () => {
  it('finds a header below a title and stops before footnotes', async () => {
    const wb = new ExcelJS.Workbook()
    await wb.xlsx.load(normalizeXlsx(scoreboard()))
    expect(detectLayout(wb.worksheets[0])).toEqual({ headerRow: 4, lastTableRow: 14, lastRow: 16, lastCol: 6 })
  })

  it('uses row 1 for an ordinary sheet', () => {
    const ws = new ExcelJS.Workbook().addWorksheet('S')
    ws.addRow(['Name', 'Points'])
    ws.addRow(['Ann', 10])
    ws.addRow(['Bob', 20])
    expect(detectLayout(ws)).toMatchObject({ headerRow: 1, lastTableRow: 3 })
  })

  it('skips a key/value block above a wider table', () => {
    const ws = new ExcelJS.Workbook().addWorksheet('S')
    ws.addRow(['Report', 'Q3'])
    ws.addRow(['Owner', 'Ann'])
    ws.addRow([])
    ws.addRow(['ID', 'Name', 'Tier', 'Points'])
    ws.addRow([1, 'Ann', 'Gold', 10])
    expect(detectLayout(ws).headerRow).toBe(4)
  })
})

describe('useWorkbook with the openpyxl scoreboard', () => {
  async function open() {
    const book = useWorkbook()
    await book.load(scoreboard(), 'lab-scoreboard.xlsx')
    return book
  }

  it('shows the real header row and only the table rows', async () => {
    const book = await open()
    expect(book.headerRow.value).toBe(4)
    expect(book.columns.value.map((c) => c.title)).toEqual([
      'Lab',
      'ข้อที่ทำได้ (/5)',
      'Prompts',
      'Agent steps',
      'Tokens',
      'นาที',
    ])
    expect(book.rows.value.map((r) => r.id)).toEqual([5, 6, 7, 8, 9, 10, 11, 12, 13, 14])
  })

  it('computes the totals, and updates them after an edit', async () => {
    const book = await open()
    const total = () => book.rows.value.find((r) => r.id === 14)!.cells
    expect(total().c2?.text).toBe('4 / 45')
    expect(total().c5?.text).toBe('166k')
    expect(total().c6?.text).toBe('26s')

    book.setCell(6, 2, '5')
    book.setCell(6, 5, '34000')
    expect(total().c2?.text).toBe('9 / 45')
    expect(total().c5?.text).toBe('200k')
    expect(book.gridView(14, 2)?.text).toBe('9 / 45')

    book.undo()
    expect(total().c5?.text).toBe('166k')
  })
})
