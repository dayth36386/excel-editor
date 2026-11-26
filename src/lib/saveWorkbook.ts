import type { Workbook } from 'exceljs'
import { colName } from './excel'

type Validation = Record<string, unknown>
interface ValidationStore {
  model: Record<string, Validation | undefined>
}

const ADDRESS = /^([A-Z]+)(\d+)$/
const colNumber = (letters: string) => [...letters].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0)
const sameRule = (a: Validation, b: Validation) => a === b || JSON.stringify(a) === JSON.stringify(b)

/**
 * ExcelJS keeps data validations per cell ("B5", "B6", …) and merges them back into
 * ranges when writing, but it sorts addresses as text ("B10" before "B5") and never
 * skips cells it already used, so it writes overlapping ranges like B10:B13 and B5:B13.
 * Excel treats that as a corrupt file and offers to repair it (dropping the validation).
 *
 * This merges the cells into non-overlapping rectangles ourselves; ExcelJS writes keys
 * that already are ranges ("B5:B13") as they are.
 */
export function mergeDataValidations(model: ValidationStore['model']): Record<string, Validation> {
  const out: Record<string, Validation> = {}
  const cells: { r: number; c: number; rule: Validation }[] = []
  for (const [key, rule] of Object.entries(model)) {
    if (!rule) continue
    const m = ADDRESS.exec(key)
    if (m) cells.push({ r: Number(m[2]), c: colNumber(m[1]), rule })
    else out[key] = rule // already a range
  }
  cells.sort((a, b) => a.r - b.r || a.c - b.c)

  const at = new Map(cells.map((x) => [`${x.r}:${x.c}`, x.rule]))
  const used = new Set<string>()
  const free = (r: number, c: number, rule: Validation) => {
    const other = at.get(`${r}:${c}`)
    return other !== undefined && !used.has(`${r}:${c}`) && sameRule(other, rule)
  }

  for (const { r, c, rule } of cells) {
    if (used.has(`${r}:${c}`)) continue
    let h = 1
    while (free(r + h, c, rule)) h++
    let w = 1
    while (Array.from({ length: h }, (_, i) => free(r + i, c + w, rule)).every(Boolean)) w++
    for (let i = 0; i < h; i++) for (let j = 0; j < w; j++) used.add(`${r + i}:${c + j}`)
    const topLeft = `${colName(c)}${r}`
    out[h > 1 || w > 1 ? `${topLeft}:${colName(c + w - 1)}${r + h - 1}` : topLeft] = rule
  }
  return out
}

/**
 * Write a workbook to .xlsx bytes, working around ExcelJS output that Excel would
 * otherwise "repair". The in-memory workbook is left exactly as it was.
 */
export async function writeWorkbook(wb: Workbook): Promise<ArrayBuffer> {
  // Formulas aren't all recalculated in the browser, so ask Excel to do it on open.
  wb.calcProperties = { ...wb.calcProperties, fullCalcOnLoad: true }
  const restore = wb.worksheets.map((ws) => {
    const store = (ws as unknown as { dataValidations: ValidationStore }).dataValidations
    const original = store.model
    store.model = mergeDataValidations(original)
    return () => (store.model = original)
  })
  try {
    return (await wb.xlsx.writeBuffer()) as ArrayBuffer
  } finally {
    restore.forEach((undo) => undo())
  }
}
