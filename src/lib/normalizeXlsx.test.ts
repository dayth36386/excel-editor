import { readFileSync } from 'node:fs'
import ExcelJS from 'exceljs'
import { describe, expect, it } from 'vitest'
import { normalizeXlsx } from './normalizeXlsx'

// A workbook saved by Python's openpyxl with a cell comment: comments live in
// xl/comments/comment1.xml and are referenced with absolute "/xl/…" targets.
const fixture = () => {
  const buf = readFileSync(new URL('./__fixtures__/openpyxl-comments.xlsx', import.meta.url))
  return buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength) as ArrayBuffer
}

async function load(data: ArrayBuffer) {
  const wb = new ExcelJS.Workbook()
  await wb.xlsx.load(data)
  return wb
}

describe('normalizeXlsx', () => {
  it('reproduces the ExcelJS failure on the raw openpyxl file', async () => {
    await expect(load(fixture())).rejects.toThrow(/comments/)
  })

  it('makes the file loadable, keeping values, formulas and the comment', async () => {
    const ws = (await load(normalizeXlsx(fixture()))).worksheets[0]
    expect(ws.getCell('A5').value).toBe('0 Baseline')
    expect(ws.getCell('B5').value).toBe(4)
    expect(ws.getCell('E5').value).toBe(166000)
    expect(ws.getCell('B14').formula).toBe('SUM(B5:B13)')
    expect(ws.getCell('B5').note).toBeTruthy()
  })

  it('still saves and reopens after normalising', async () => {
    const wb = await load(normalizeXlsx(fixture()))
    wb.worksheets[0].getCell('B6').value = 5
    const again = await load((await wb.xlsx.writeBuffer()) as ArrayBuffer)
    expect(again.worksheets[0].getCell('B6').value).toBe(5)
    expect(again.worksheets[0].getCell('B5').note).toBeTruthy()
  })

  it('returns files written by Excel/ExcelJS untouched', async () => {
    const wb = new ExcelJS.Workbook()
    const ws = wb.addWorksheet('S')
    ws.getCell('A1').value = 'x'
    ws.getCell('A1').note = 'hello'
    const buf = (await wb.xlsx.writeBuffer()) as ArrayBuffer
    expect(normalizeXlsx(buf)).toBe(buf)
  })
})
