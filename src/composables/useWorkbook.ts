import { computed, markRaw, ref, shallowRef } from 'vue'
import type { Cell, CellValue, Workbook, Worksheet } from 'exceljs'
import {
  cellData,
  colName,
  dateToSerial,
  displayText,
  editText,
  isMergeSlave,
  parseInput,
  viewOf,
  type CellData,
  type CellKind,
  type CellView,
  type FormulaResult,
} from '@/lib/excel'
import { evaluateFormula, type Scalar } from '@/lib/formula'
import { detectLayout } from '@/lib/layout'
import * as ops from '@/lib/sheetOps'
import type { LineSnapshot, StyleFrom } from '@/lib/sheetOps'
import { canSaveAs, downloadBlob, ensureWritable, pickSaveTarget, writeToHandle, XLSX_MIME } from '@/lib/fileAccess'

export type SaveResult = 'written' | 'downloaded' | 'cancelled'
type Axis = 'row' | 'col'

export interface TableColumn {
  /** Column id used by the table, e.g. "c3" for Excel column C. */
  key: string
  col: number
  /** Excel column letter, e.g. "C". */
  letter: string
  /** Header text; empty when the header cell is blank (the UI shows a fallback). */
  title: string
  kind: CellKind
  /** Set for short text columns with a few repeating values (status, tier…): their distinct values. */
  options?: string[]
}

/** Thrown by renameSheet; `reason` lets the UI show a translated message. */
export class SheetNameError extends Error {
  constructor(public reason: 'invalid' | 'duplicate') {
    super(`Invalid sheet name (${reason})`)
    this.name = 'SheetNameError'
  }
}

export interface TableRow {
  /** Excel row number. */
  id: number
  cells: Record<string, CellData | undefined>
}

interface Change {
  r: number
  c: number
  before: CellValue
  after: CellValue
}

type HistoryEntry =
  | { kind: 'cells'; sheetId: number; changes: Change[] }
  | { kind: 'insert'; axis: Axis; sheetId: number; at: number; from?: StyleFrom }
  | { kind: 'delete'; axis: Axis; sheetId: number; at: number; snap: LineSnapshot }
  | { kind: 'batch'; sheetId: number; entries: HistoryEntry[] }

const HISTORY_LIMIT = 300
const MAX_OPTIONS = 8
const DEFAULT_COL_PX = 100
const MAX_FROZEN_ROWS = 5
const MAX_FROZEN_COLS = 3

/** Excel column width (in characters) → pixels. */
const widthToPx = (w: number) => Math.min(480, Math.max(32, Math.round(w * 7 + 5)))
const key = (c: number) => `c${c}`

// ExcelJS is ~1 MB, so load it on first use instead of with the page.
const loadExcelJS = () => import('exceljs').then((m) => m.default)

/**
 * Holds the open workbook and every edit made to it. The ExcelJS workbook is
 * the single source of truth: edits go straight into its cells, so saving keeps
 * the original file's formatting, column widths, other sheets and so on.
 *
 * The active sheet is shown as a table: its first non-empty row gives the
 * column titles and every row below it is a record.
 */
export function useWorkbook() {
  const workbook = shallowRef<Workbook | null>(null)
  const fileName = ref('')
  const fileHandle = shallowRef<FileSystemFileHandle | null>(null)
  const activeId = ref(0)
  const revision = ref(0)
  const dirty = ref(false)

  const headerRow = ref(1)
  /** Last used row and column of the active sheet (header included), for the Excel view. */
  const extent = shallowRef({ rows: 0, cols: 0 })
  const columns = shallowRef<TableColumn[]>([])
  const rows = shallowRef<TableRow[]>([])
  let lastCol = 0
  /** Keeps freshly added empty rows visible (they have no values yet). */
  let keepRow = 0

  const edited = new Map<number, Set<string>>()
  let undoStack: HistoryEntry[] = []
  let redoStack: HistoryEntry[] = []

  const touch = () => void revision.value++
  const sheet = (): Worksheet | undefined => workbook.value?.getWorksheet(activeId.value)

  // ---- formulas ------------------------------------------------------------
  // Common formulas (SUM, AVERAGE, arithmetic…) are computed live from the current values,
  // so totals update while you edit and files saved without cached results (e.g. by
  // openpyxl) still show numbers. Results are cached until the next change.

  const formulaCache = new Map<string, Scalar | undefined>()
  const evaluating = new Set<string>()
  let formulaCacheRev = -1
  const clearFormulaCache = () => {
    formulaCache.clear()
    formulaCacheRev = revision.value
  }

  function toScalar(v: CellValue): Scalar {
    if (v === null || v === undefined) return null
    if (typeof v === 'number' || typeof v === 'string' || typeof v === 'boolean') return v
    if (v instanceof Date) return dateToSerial(v)
    if ('richText' in v) return v.richText.map((t) => t.text).join('')
    if ('hyperlink' in v) return typeof v.text === 'string' ? v.text : null
    if ('error' in v) return String(v.error)
    return toScalar((v as { result?: CellValue }).result ?? null)
  }

  function scalarAt(ws: Worksheet, r: number, c: number): Scalar {
    const cell = ws.findCell(r, c)
    if (!cell || isMergeSlave(cell)) return null
    if (cell.formula) return formulaValue(ws, cell) ?? toScalar(cell.value)
    return toScalar(cell.value)
  }

  function formulaValue(ws: Worksheet, cell: Cell): Scalar | undefined {
    const k = cell.address
    if (formulaCache.has(k)) return formulaCache.get(k)
    if (evaluating.has(k)) return undefined // circular reference
    evaluating.add(k)
    const value = evaluateFormula(cell.formula, (r, c) => scalarAt(ws, r, c), ws.rowCount)
    evaluating.delete(k)
    formulaCache.set(k, value)
    return value
  }

  const formulaResult: FormulaResult = (cell) => {
    const ws = sheet()
    return ws && cell.formula ? formulaValue(ws, cell) : undefined
  }

  const sheets = computed(() => {
    void revision.value
    return workbook.value?.worksheets.map((ws) => ({ id: ws.id, name: ws.name })) ?? []
  })
  const canUndo = computed(() => (void revision.value, undoStack.length > 0))
  const canRedo = computed(() => (void revision.value, redoStack.length > 0))

  // ---- building the table ------------------------------------------------

  /** Rows of the table holding formulas; they are re-read whenever any cell changes. */
  let formulaRows = new Set<number>()

  function buildRow(ws: Worksheet, r: number): TableRow {
    const row = ws.findRow(r)
    const cells: TableRow['cells'] = {}
    if (row) {
      for (let c = 1; c <= lastCol; c++) {
        const cell = row.findCell(c)
        if (!cell) continue
        cells[key(c)] = cellData(cell, formulaResult)
        if (cell.formula) formulaRows.add(r)
      }
    }
    return { id: r, cells }
  }

  function guessKind(list: TableRow[], k: string): CellKind {
    const counts: Record<CellKind, number> = { text: 0, number: 0, date: 0, boolean: 0 }
    let seen = 0
    for (const row of list) {
      const cell = row.cells[k]
      if (!cell) continue
      counts[cell.kind]++
      if (++seen >= 500) break
    }
    const best = (Object.keys(counts) as CellKind[]).reduce((a, b) => (counts[b] > counts[a] ? b : a))
    return counts[best] > 0 && counts[best] >= seen * 0.6 ? best : 'text'
  }

  /** Short text columns where a handful of values repeat are treated as categories. */
  function categoryOptions(list: TableRow[], k: string): string[] | undefined {
    const seen = new Set<string>()
    let filled = 0
    for (const row of list) {
      const text = row.cells[k]?.text
      if (!text) continue
      if (text.length > 24) return undefined
      filled++
      seen.add(text)
      if (seen.size > MAX_OPTIONS) return undefined
    }
    if (filled < 6 || seen.size < 2 || filled < list.length * 0.5 || seen.size > filled / 3) return undefined
    return [...seen].sort((a, b) => a.localeCompare(b, 'th'))
  }

  /** Header text of column c; a cell merged from the left (part of a wide title) counts as blank. */
  function headerTitle(ws: Worksheet, c: number): string {
    const cell = ws.findCell(headerRow.value, c)
    return cell && !isMergeSlave(cell) ? displayText(cell.value).trim() : ''
  }

  function buildColumn(ws: Worksheet, c: number, list: TableRow[]): TableColumn {
    const k = key(c)
    const kind = guessKind(list, k)
    return {
      key: k,
      col: c,
      letter: colName(c),
      title: headerTitle(ws, c),
      kind,
      options: kind === 'text' ? categoryOptions(list, k) : undefined,
    }
  }

  function rebuild() {
    const ws = sheet()
    if (!ws) return
    clearFormulaCache()
    formulaRows = new Set()
    const layout = detectLayout(ws)
    headerRow.value = layout.headerRow
    lastCol = Math.max(layout.lastCol, 1)
    const lastRow = Math.max(layout.lastTableRow, keepRow, layout.headerRow)
    extent.value = { rows: Math.max(layout.lastRow, lastRow), cols: lastCol }

    const list: TableRow[] = []
    for (let r = layout.headerRow + 1; r <= lastRow; r++) list.push(buildRow(ws, r))
    rows.value = list

    const cols: TableColumn[] = []
    for (let c = 1; c <= lastCol; c++) cols.push(buildColumn(ws, c, list))
    columns.value = cols
  }

  /** Re-read one cell's row (and its column's type) after an edit, without rebuilding everything. */
  function refreshCell(ws: Worksheet, r: number, c: number) {
    const i = r - headerRow.value - 1
    if (i < 0 || i >= rows.value.length) return rebuild()
    clearFormulaCache()
    const next = rows.value.slice()
    next[i] = buildRow(ws, r)
    // Totals and other formulas may depend on the cell that changed.
    for (const fr of formulaRows) {
      const j = fr - headerRow.value - 1
      if (fr !== r && j >= 0 && j < next.length) next[j] = buildRow(ws, fr)
    }
    rows.value = next
    columns.value = columns.value.map((col) => (col.col === c ? buildColumn(ws, c, rows.value) : col))
  }

  // ---- loading -----------------------------------------------------------

  async function load(data: ArrayBuffer, name: string, handle: FileSystemFileHandle | null = null) {
    const [ExcelJS, { normalizeXlsx }] = await Promise.all([loadExcelJS(), import('@/lib/normalizeXlsx')])
    const wb = new ExcelJS.Workbook()
    await wb.xlsx.load(normalizeXlsx(data))
    if (!wb.worksheets.length) wb.addWorksheet('Sheet1')
    reset(wb, name, handle)
  }

  /** New workbook with one sheet whose header row holds `headers`. */
  async function createBlank(headers: string[]) {
    const ExcelJS = await loadExcelJS()
    const wb = new ExcelJS.Workbook()
    const ws = wb.addWorksheet('Sheet1')
    ws.addRow(headers).font = { bold: true }
    headers.forEach((_, i) => (ws.getColumn(i + 1).width = 18))
    reset(wb, 'Book1.xlsx', null)
    keepRow = 2
    rebuild()
  }

  function reset(wb: Workbook, name: string, handle: FileSystemFileHandle | null) {
    workbook.value = markRaw(wb)
    fileName.value = name
    fileHandle.value = handle
    edited.clear()
    undoStack = []
    redoStack = []
    dirty.value = false
    selectSheet(wb.worksheets[0].id)
  }

  function selectSheet(id: number) {
    if (!workbook.value?.getWorksheet(id)) return
    activeId.value = id
    keepRow = 0
    rebuild()
    touch()
  }

  // ---- cells -------------------------------------------------------------

  function getEditText(r: number, c: number) {
    const cell = sheet()?.findCell(r, c)
    return cell ? editText(cell) : ''
  }

  const isEdited = (r: number, c: number) => edited.get(activeId.value)?.has(`${r}:${c}`) ?? false

  /** Cells changed since the last save, across all sheets. */
  const editedCount = computed(() => {
    void revision.value
    let n = 0
    for (const set of edited.values()) n += set.size
    return n
  })

  function markEdited(sheetId: number, r: number, c: number) {
    let set = edited.get(sheetId)
    if (!set) edited.set(sheetId, (set = new Set()))
    set.add(`${r}:${c}`)
  }

  function record(entry: HistoryEntry) {
    undoStack.push(entry)
    if (undoStack.length > HISTORY_LIMIT) undoStack.shift()
    redoStack = []
    dirty.value = true
    touch()
  }

  function afterCellChange(ws: Worksheet, r: number, c: number) {
    markEdited(ws.id, r, c)
    if (r === headerRow.value || c > lastCol || r < headerRow.value) rebuild()
    else refreshCell(ws, r, c)
  }

  const setCell = (r: number, c: number, text: string) => setCells([{ r, c, text }])

  /** Write several cells as one undo step (e.g. a paste). */
  function setCells(entries: { r: number; c: number; text: string }[]) {
    const ws = sheet()
    if (!ws) return
    const changes: Change[] = []
    for (const { r, c, text } of entries) {
      const cell = ws.getCell(r, c)
      if (isMergeSlave(cell) || editText(cell) === text) continue
      const before = cell.value
      const after = parseInput(text, before, cell.numFmt)
      cell.value = after
      changes.push({ r, c, before, after })
    }
    if (!changes.length) return
    if (changes.length === 1) afterCellChange(ws, changes[0].r, changes[0].c)
    else {
      for (const ch of changes) markEdited(ws.id, ch.r, ch.c)
      rebuild()
    }
    record({ kind: 'cells', sheetId: ws.id, changes })
  }

  function pasteBlock(r0: number, c0: number, block: string[][]) {
    setCells(block.flatMap((line, i) => line.map((text, j) => ({ r: r0 + i, c: c0 + j, text }))))
  }

  // ---- Excel view --------------------------------------------------------

  /** A cell as the Excel view draws it, with its Excel formatting. */
  function gridView(r: number, c: number): CellView | undefined {
    if (formulaCacheRev !== revision.value) clearFormulaCache()
    const cell = sheet()?.findCell(r, c)
    return cell ? viewOf(cell, formulaResult) : undefined
  }

  /** Pixel widths of columns 1..n, from the sheet's column widths. */
  function colWidths(n: number): number[] {
    const ws = sheet()
    const defined = ws?.columns?.length ?? 0
    const fallback = ws?.properties.defaultColWidth ? widthToPx(ws.properties.defaultColWidth) : DEFAULT_COL_PX
    const out: number[] = []
    for (let c = 1; c <= n; c++) {
      const w = ws && c <= defined ? ws.getColumn(c).width : undefined
      out.push(w ? widthToPx(w) : fallback)
    }
    return out
  }

  /** Frozen panes from the sheet's view settings. */
  const frozen = computed(() => {
    void revision.value
    const view = sheet()?.views?.[0]
    if (view?.state !== 'frozen') return { rows: 0, cols: 0 }
    return { rows: Math.min(view.ySplit ?? 0, MAX_FROZEN_ROWS), cols: Math.min(view.xSplit ?? 0, MAX_FROZEN_COLS) }
  })

  /** Merged cells are edited through their top-left (master) cell. */
  function resolve(r: number, c: number) {
    const cell = sheet()?.findCell(r, c)
    if (cell && isMergeSlave(cell)) return { r: cell.master.fullAddress.row, c: cell.master.fullAddress.col }
    return { r, c }
  }

  // ---- rows & columns ----------------------------------------------------

  function shiftEdited(sheetId: number, axis: Axis, at: number, delta: 1 | -1) {
    const set = edited.get(sheetId)
    if (!set) return
    const next = new Set<string>()
    for (const k of set) {
      const [r, c] = k.split(':').map(Number)
      const v = axis === 'row' ? r : c
      if (delta < 0 && v === at) continue
      const moved = v >= at ? v + delta : v
      next.add(axis === 'row' ? `${moved}:${c}` : `${r}:${moved}`)
    }
    edited.set(sheetId, next)
  }

  function addLine(ws: Worksheet, axis: Axis, at: number, snap?: LineSnapshot, from?: StyleFrom) {
    if (axis === 'row') {
      ops.insertRow(ws, at, snap, from)
      keepRow = Math.max(keepRow + (keepRow >= at ? 1 : 0), at)
    } else ops.insertCol(ws, at, snap)
    shiftEdited(ws.id, axis, at, 1)
  }

  function removeLine(ws: Worksheet, axis: Axis, at: number): LineSnapshot {
    const snap = axis === 'row' ? ops.snapshotRow(ws, at) : ops.snapshotCol(ws, at)
    if (axis === 'row') {
      ops.deleteRow(ws, at)
      if (keepRow >= at) keepRow--
    } else ops.deleteCol(ws, at)
    shiftEdited(ws.id, axis, at, -1)
    return snap
  }

  /** Insert an empty row at Excel row `at` (default: after the last row). Returns its row number. */
  function insertRow(at?: number): number {
    const ws = sheet()
    if (!ws) return 0
    const end = headerRow.value + rows.value.length + 1
    const r = Math.min(at ?? end, end)
    // Inside the table copy the row being pushed down; at the end copy the last row.
    const from: StyleFrom = r >= end ? 'above' : 'below'
    addLine(ws, 'row', r, undefined, from)
    rebuild()
    record({ kind: 'insert', axis: 'row', sheetId: ws.id, at: r, from })
    return r
  }

  function deleteRows(rowNumbers: number[]) {
    const ws = sheet()
    if (!ws || !rowNumbers.length) return
    // Bottom-up, so earlier deletes don't shift the rows still to delete.
    const sorted = [...new Set(rowNumbers)].sort((a, b) => b - a)
    const entries: HistoryEntry[] = sorted.map((at) => ({
      kind: 'delete',
      axis: 'row',
      sheetId: ws.id,
      at,
      snap: removeLine(ws, 'row', at),
    }))
    rebuild()
    record({ kind: 'batch', sheetId: ws.id, entries })
  }

  function addColumn(title: string) {
    const ws = sheet()
    if (!ws) return
    const c = lastCol + 1
    const h = headerRow.value
    // Make the new header look like its neighbours.
    const prev = ws.findCell(h, c - 1)
    if (prev) ws.getCell(h, c).style = structuredClone(prev.style)
    const prevWidth = ws.getColumn(c - 1).width
    if (prevWidth) ws.getColumn(c).width = prevWidth
    setCell(h, c, title)
  }

  const renameColumn = (c: number, title: string) => setCell(headerRow.value, c, title)

  /** Insert an empty column at Excel column `at` (Excel view). */
  function insertColumn(at: number) {
    const ws = sheet()
    if (!ws) return
    addLine(ws, 'col', at)
    rebuild()
    record({ kind: 'insert', axis: 'col', sheetId: ws.id, at })
  }

  function deleteColumn(c: number) {
    const ws = sheet()
    if (!ws) return
    const snap = removeLine(ws, 'col', c)
    rebuild()
    record({ kind: 'delete', axis: 'col', sheetId: ws.id, at: c, snap })
  }

  // ---- undo / redo -------------------------------------------------------

  function apply(entry: HistoryEntry, dir: 'undo' | 'redo', nested = false) {
    const ws = workbook.value?.getWorksheet(entry.sheetId)
    if (!ws) return
    if (activeId.value !== ws.id) selectSheet(ws.id)

    if (entry.kind === 'batch') {
      const list = dir === 'undo' ? [...entry.entries].reverse() : entry.entries
      for (const e of list) apply(e, dir, true)
    } else if (entry.kind === 'cells') {
      for (const ch of entry.changes) {
        ws.getCell(ch.r, ch.c).value = dir === 'undo' ? ch.before : ch.after
        markEdited(ws.id, ch.r, ch.c)
      }
    } else if ((entry.kind === 'insert') === (dir === 'redo')) {
      addLine(
        ws,
        entry.axis,
        entry.at,
        entry.kind === 'delete' ? entry.snap : undefined,
        entry.kind === 'insert' ? entry.from : undefined,
      )
    } else {
      const snap = removeLine(ws, entry.axis, entry.at)
      if (entry.kind === 'delete') entry.snap = snap
    }
    if (!nested) rebuild()
  }

  function undo() {
    const entry = undoStack.pop()
    if (!entry) return
    apply(entry, 'undo')
    redoStack.push(entry)
    dirty.value = true
    touch()
  }

  function redo() {
    const entry = redoStack.pop()
    if (!entry) return
    apply(entry, 'redo')
    undoStack.push(entry)
    dirty.value = true
    touch()
  }

  // ---- sheets ------------------------------------------------------------

  function addSheet(headers: string[]) {
    const wb = workbook.value
    if (!wb) return
    let n = wb.worksheets.length + 1
    while (wb.getWorksheet(`Sheet${n}`)) n++
    const ws = wb.addWorksheet(`Sheet${n}`)
    ws.addRow(headers).font = { bold: true }
    dirty.value = true
    selectSheet(ws.id)
    keepRow = 2
    rebuild()
  }

  /** Throws SheetNameError if Excel would reject the name. */
  function renameSheet(id: number, name: string) {
    const wb = workbook.value
    const ws = wb?.getWorksheet(id)
    if (!wb || !ws || ws.name === name) return
    if (!name || name.length > 31 || /[\\/?*[\]:]/.test(name)) throw new SheetNameError('invalid')
    if (wb.worksheets.some((w) => w.id !== id && w.name.toLowerCase() === name.toLowerCase())) {
      throw new SheetNameError('duplicate')
    }
    ws.name = name
    dirty.value = true
    touch()
  }

  // ---- saving ------------------------------------------------------------

  async function toBlob(): Promise<Blob> {
    const wb = workbook.value!
    // Formulas aren't recalculated here, so ask Excel to do it when it opens the file.
    wb.calcProperties = { ...wb.calcProperties, fullCalcOnLoad: true }
    const buf = await wb.xlsx.writeBuffer()
    return new Blob([buf], { type: XLSX_MIME })
  }

  function markSaved() {
    dirty.value = false
    edited.clear()
    touch()
  }

  /** Write back to the opened file when the browser allows it, otherwise save as / download. */
  async function save(): Promise<SaveResult> {
    const handle = fileHandle.value
    if (!handle) return canSaveAs ? saveAs() : download()
    // Ask for write permission first, while we still have the user's click.
    await ensureWritable(handle)
    await writeToHandle(handle, await toBlob())
    markSaved()
    return 'written'
  }

  async function saveAs(): Promise<SaveResult> {
    if (!canSaveAs) return download()
    const handle = await pickSaveTarget(fileName.value)
    if (!handle) return 'cancelled'
    await writeToHandle(handle, await toBlob())
    fileHandle.value = handle
    fileName.value = handle.name
    markSaved()
    return 'written'
  }

  async function download(): Promise<SaveResult> {
    downloadBlob(await toBlob(), fileName.value)
    markSaved()
    return 'downloaded'
  }

  return {
    workbook,
    fileName,
    fileHandle,
    activeId,
    revision,
    dirty,
    sheets,
    headerRow,
    columns,
    rows,
    editedCount,
    canUndo,
    canRedo,
    load,
    createBlank,
    selectSheet,
    getEditText,
    isEdited,
    setCell,
    insertRow,
    deleteRows,
    addColumn,
    renameColumn,
    deleteColumn,
    insertColumn,
    setCells,
    pasteBlock,
    extent,
    gridView,
    colWidths,
    frozen,
    resolve,
    undo,
    redo,
    addSheet,
    renameSheet,
    save,
    saveAs,
    download,
  }
}

export type WorkbookStore = ReturnType<typeof useWorkbook>
