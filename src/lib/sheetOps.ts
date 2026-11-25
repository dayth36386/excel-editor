import type { CellValue, Style, Worksheet } from 'exceljs'
import { isMergeSlave } from './excel'

/** Values and styles of one row or column, kept so a delete can be undone. */
export interface LineSnapshot {
  cells: { i: number; value: CellValue; style: Partial<Style> }[]
  /** Row height or column width. */
  size?: number
}

// ExcelJS shares style objects between cells; copy before reusing one.
const cloneStyle = (s: Partial<Style>): Partial<Style> => structuredClone(s ?? {})

export function snapshotRow(ws: Worksheet, at: number): LineSnapshot {
  const row = ws.getRow(at)
  const cells: LineSnapshot['cells'] = []
  row.eachCell({ includeEmpty: true }, (cell, c) => {
    cells.push({ i: c, value: isMergeSlave(cell) ? null : cell.value, style: cloneStyle(cell.style) })
  })
  return { cells, size: row.height }
}

export function snapshotCol(ws: Worksheet, at: number): LineSnapshot {
  const col = ws.getColumn(at)
  const cells: LineSnapshot['cells'] = []
  col.eachCell({ includeEmpty: true }, (cell, r) => {
    cells.push({ i: r, value: isMergeSlave(cell) ? null : cell.value, style: cloneStyle(cell.style) })
  })
  return { cells, size: col.width }
}

/** Which neighbouring row a new empty row copies its formatting from. */
export type StyleFrom = 'above' | 'below'

/**
 * Insert an empty row at `at`. Without a snapshot the new row copies the
 * formatting of the row above or below it, so it looks like the rest of the
 * table; with one, the snapshot is put back (undo of a delete).
 */
export function insertRow(ws: Worksheet, at: number, snap?: LineSnapshot, from: StyleFrom = 'below') {
  if (!snap) {
    ws.insertRow(at, [], from === 'above' ? 'i+' : 'o+')
    return
  }
  ws.spliceRows(at, 0, [])
  const row = ws.getRow(at)
  for (const { i, value, style } of snap.cells) {
    const cell = row.getCell(i)
    cell.style = cloneStyle(style)
    cell.value = value
  }
  if (snap.size) row.height = snap.size
}

export function deleteRow(ws: Worksheet, at: number) {
  ws.spliceRows(at, 1)
}

/** Insert an empty column at `at`; same formatting rules as insertRow. */
export function insertCol(ws: Worksheet, at: number, snap?: LineSnapshot) {
  ws.spliceColumns(at, 0, [])
  if (snap) {
    for (const { i, value, style } of snap.cells) {
      const cell = ws.getCell(i, at)
      cell.style = cloneStyle(style)
      cell.value = value
    }
    if (snap.size) ws.getColumn(at).width = snap.size
    return
  }
  const right = ws.getColumn(at + 1)
  right.eachCell({ includeEmpty: true }, (cell, r) => {
    ws.getCell(r, at).style = cloneStyle(cell.style)
  })
  if (right.width) ws.getColumn(at).width = right.width
}

export function deleteCol(ws: Worksheet, at: number) {
  ws.spliceColumns(at, 1)
}
