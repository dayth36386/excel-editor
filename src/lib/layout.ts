import type { Cell, Worksheet } from 'exceljs'
import { displayText, isMergeSlave } from './excel'

export interface SheetLayout {
  /** Row whose cells become the column titles in the table view. */
  headerRow: number
  /** Last row of the table (after the header), not counting notes and blank rows below it. */
  lastTableRow: number
  /** Last used row and column of the whole sheet. */
  lastRow: number
  lastCol: number
}

interface RowInfo {
  /** Columns holding a value (merged copies excluded). */
  cols: number[]
  texts: number
  /** Only holds a cell merged across several columns: a title, legend or footnote. */
  banner: boolean
}

const HEADER_SEARCH_ROWS = 30

/** Number of columns a merged cell spans (1 when it isn't merged). */
function mergeWidth(ws: Worksheet, cell: Cell): number {
  if (!cell.isMerged) return 1
  let w = 1
  while (ws.findCell(cell.fullAddress.row, cell.fullAddress.col + w)?.master?.address === cell.address) w++
  return w
}

/**
 * Work out where the table is on a sheet. Real-world sheets often start with a title,
 * a note or a blank row, and end with footnotes, so "the first row is the header" is
 * not good enough:
 *
 * - The header is the first row (within the first 30) that holds at least 60% as many
 *   values as the fullest row there, mostly text, and isn't a single merged banner.
 * - The table ends at the last row below it that isn't blank or a merged banner,
 *   so footnotes merged across the page stay out of the table.
 */
export function detectLayout(ws: Worksheet): SheetLayout {
  const rows = new Map<number, RowInfo>()
  let lastRow = 0
  let lastCol = 0

  ws.eachRow({ includeEmpty: false }, (row, r) => {
    const info: RowInfo = { cols: [], texts: 0, banner: false }
    let widest = 1
    row.eachCell({ includeEmpty: false }, (cell, c) => {
      if (isMergeSlave(cell) || displayText(cell.value) === '') return
      info.cols.push(c)
      if (
        typeof cell.value === 'string' ||
        (cell.value && typeof cell.value === 'object' && 'richText' in cell.value)
      ) {
        info.texts++
      }
      widest = Math.max(widest, mergeWidth(ws, cell))
      lastCol = Math.max(lastCol, c)
    })
    if (!info.cols.length) return
    info.banner = info.cols.length === 1 && widest > 1
    rows.set(r, info)
    lastRow = Math.max(lastRow, r)
  })

  // eachRow visits rows in order, so the first key is the first row with a value.
  const firstValueRow = rows.keys().next().value ?? 1
  const firstRows = [...rows.entries()].filter(([r]) => r < firstValueRow + HEADER_SEARCH_ROWS)
  const fullest = Math.max(0, ...firstRows.map(([, i]) => (i.banner ? 0 : i.cols.length)))
  const isHeader = (i: RowInfo) =>
    !i.banner && i.cols.length >= Math.max(2, Math.ceil(fullest * 0.6)) && i.texts >= i.cols.length / 2
  const headerRow =
    firstRows.find(([, i]) => isHeader(i))?.[0] ?? firstRows.find(([, i]) => !i.banner)?.[0] ?? firstRows[0]?.[0] ?? 1

  let lastTableRow = headerRow
  for (const [r, i] of rows) if (r > headerRow && !i.banner) lastTableRow = Math.max(lastTableRow, r)

  return { headerRow, lastTableRow, lastRow, lastCol }
}
