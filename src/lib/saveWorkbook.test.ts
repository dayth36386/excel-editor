import { readFileSync } from 'node:fs'
import ExcelJS from 'exceljs'
import { strFromU8, unzipSync } from 'fflate'
import { describe, expect, it } from 'vitest'
import { normalizeXlsx } from './normalizeXlsx'
import { mergeDataValidations, writeWorkbook } from './saveWorkbook'

const colNumber = (letters: string) => [...letters].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0)

/** Every cell covered by the <dataValidation sqref="…"> ranges of a saved sheet, with how often. */
function validationCoverage(xlsx: ArrayBuffer, sheet = 'xl/worksheets/sheet1.xml') {
  const xml = strFromU8(unzipSync(new Uint8Array(xlsx))[sheet])
  const counts = new Map<string, number>()
  for (const [, sqref] of xml.matchAll(/<dataValidation\b[^>]*\bsqref="([^"]+)"/g)) {
    for (const part of sqref.split(' ')) {
      const [a, b = a] = part.split(':')
      const [, c1, r1] = /([A-Z]+)(\d+)/.exec(a)!
      const [, c2, r2] = /([A-Z]+)(\d+)/.exec(b)!
      for (let r = +r1; r <= +r2; r++)
        for (let c = colNumber(c1); c <= colNumber(c2); c++) counts.set(`${r}:${c}`, (counts.get(`${r}:${c}`) ?? 0) + 1)
    }
  }
  return counts
}

const rule = (min: number, max: number) => ({ type: 'whole', operator: 'between', formulae: [min, max] })

describe('mergeDataValidations', () => {
  it('merges per-cell rules into non-overlapping rectangles', () => {
    const score = rule(0, 5)
    const count = rule(0, 999)
    const model: Record<string, Record<string, unknown>> = {}
    for (let r = 5; r <= 13; r++) {
      model[`B${r}`] = score
      for (const c of 'CDEF') model[`${c}${r}`] = { ...count } // equal content, different objects
    }
    expect(Object.keys(mergeDataValidations(model)).sort()).toEqual(['B5:B13', 'C5:F13'])
  })

  it('keeps existing ranges, single cells and drops removed rules', () => {
    const r = rule(1, 2)
    // L-shape A1, A2, B1: grows down first, so A1:A2 plus B1.
    expect(mergeDataValidations({ 'E1:E3': r, A1: r, A2: r, B1: r, C9: rule(3, 4), D1: undefined })).toEqual({
      'E1:E3': r,
      'A1:A2': r,
      B1: r,
      C9: rule(3, 4),
    })
  })
})

describe('writeWorkbook', () => {
  it('saves the openpyxl scoreboard without overlapping validations (Excel "repair" bug)', async () => {
    const buf = readFileSync(new URL('./__fixtures__/openpyxl-comments.xlsx', import.meta.url))
    const wb = new ExcelJS.Workbook()
    await wb.xlsx.load(normalizeXlsx(buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength) as ArrayBuffer))
    wb.worksheets[0].getCell('B6').value = 5

    // What ExcelJS alone writes: B10:B13 and B5:B13 overlap.
    const plain = validationCoverage((await wb.xlsx.writeBuffer()) as ArrayBuffer)
    expect(Math.max(...plain.values())).toBeGreaterThan(1)

    const saved = await writeWorkbook(wb)
    const coverage = validationCoverage(saved)
    expect(Math.max(...coverage.values())).toBe(1)
    expect(coverage.size).toBe(9 * 5) // B5:F13

    // The rules still work after reopening, and the in-memory workbook is unchanged.
    const again = new ExcelJS.Workbook()
    await again.xlsx.load(saved)
    expect(again.worksheets[0].getCell('B9').dataValidation).toMatchObject({ type: 'whole', formulae: [0, 5] })
    expect(again.worksheets[0].getCell('B6').value).toBe(5)
    expect(wb.worksheets[0].getCell('B10').dataValidation).toBeTruthy()
  })
})
