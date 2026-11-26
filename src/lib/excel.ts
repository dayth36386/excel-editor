import type { Cell, CellFormulaValue, CellSharedFormulaValue, CellValue } from 'exceljs'
import type { Scalar } from './formula'

export type CellKind = 'text' | 'number' | 'date' | 'boolean'

/** What the table needs to show, sort and search one cell. */
export interface CellData {
  text: string
  sort: string | number
  kind: CellKind
  formula?: boolean
}

/** 1 → A, 26 → Z, 27 → AA */
export function colName(n: number): string {
  let s = ''
  while (n > 0) {
    const m = (n - 1) % 26
    s = String.fromCharCode(65 + m) + s
    n = Math.floor((n - 1) / 26)
  }
  return s
}

const pad = (n: number) => String(n).padStart(2, '0')

/** ExcelJS reads Excel dates as UTC instants, so format them in UTC. */
export function formatDate(d: Date): string {
  if (Number.isNaN(d.getTime())) return ''
  const day = `${d.getUTCFullYear()}-${pad(d.getUTCMonth() + 1)}-${pad(d.getUTCDate())}`
  const h = d.getUTCHours()
  const m = d.getUTCMinutes()
  const s = d.getUTCSeconds()
  if (!h && !m && !s) return day
  return `${day} ${pad(h)}:${pad(m)}${s ? `:${pad(s)}` : ''}`
}

/** Hide binary floating-point noise such as 0.30000000000000004. */
export function formatNumber(n: number): string {
  if (!Number.isFinite(n)) return String(n)
  return String(Number(n.toPrecision(15)))
}

// prefix, integer digits, decimals, scaling commas (each one ÷1000, as in #,##0,"k"), percent, suffix
const NUMFMT = /^([^#0?]*?)([#0?,]*[#0?])(?:\.([#0?]+))?(,*)(%?)([^#0?]*)$/

/**
 * Apply the common Excel number formats: thousands separators, fixed
 * decimals, percentages and literal prefixes/suffixes such as "฿#,##0.00".
 * Anything fancier falls back to the plain number.
 */
export function formatWithNumFmt(n: number, numFmt?: string): string {
  if (!numFmt || numFmt === 'General' || !Number.isFinite(n)) return formatNumber(n)
  const section = numFmt
    .split(';')[0]
    .replace(/"([^"]*)"|\\(.)|\[[^\]]*\]|_.|\*./g, (_, quoted?: string, escaped?: string) => quoted ?? escaped ?? '')
  const m = NUMFMT.exec(section)
  if (!m) return formatNumber(n)
  const [, prefix, int, dec = '', scaling, pct, suffix] = m
  const value = (pct ? n * 100 : n) / 1000 ** scaling.length
  const digits = new Intl.NumberFormat('en-US', {
    minimumIntegerDigits: Math.max(1, (int.match(/0/g) ?? []).length),
    minimumFractionDigits: (dec.match(/0/g) ?? []).length,
    maximumFractionDigits: dec.length,
    useGrouping: int.includes(','),
  }).format(Math.abs(value))
  const negative = value < 0 && Number(digits.replace(/,/g, '')) !== 0
  return `${negative ? '-' : ''}${prefix}${digits}${pct}${suffix}`
}

function isFormula(v: CellValue): v is CellFormulaValue | CellSharedFormulaValue {
  return typeof v === 'object' && v !== null && !(v instanceof Date) && ('formula' in v || 'sharedFormula' in v)
}

/** Plain text of a value. Formulas give their last calculated result. */
export function displayText(v: CellValue): string {
  if (v === null || v === undefined) return ''
  if (typeof v === 'number') return formatNumber(v)
  if (typeof v === 'string') return v
  if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE'
  if (v instanceof Date) return formatDate(v)
  if ('richText' in v) return v.richText.map((t) => t.text).join('')
  if (isFormula(v)) {
    const res = v.result
    if (res === undefined || res === null) return 'formula' in v ? `=${v.formula}` : ''
    return displayText(res as CellValue)
  }
  if ('hyperlink' in v) return typeof v.text === 'string' ? v.text : displayText(v.text as CellValue)
  if ('error' in v) return String(v.error)
  return String(v)
}

/** Text put in the editor. Formulas are edited as `=FORMULA`. */
export function editText(cell: Cell): string {
  const v = cell.value
  if (isFormula(v) && cell.formula) return `=${cell.formula}`
  return displayText(v)
}

export function isMergeSlave(cell: Cell): boolean {
  return cell.isMerged && cell.master.address !== cell.address
}

/** Table view of a cell, or undefined when it is empty. */
/** Live result of a formula cell, or undefined to fall back to the result cached in the file. */
export type FormulaResult = (cell: Cell) => Scalar | undefined

// Excel stores dates as days since 1899-12-30.
const EXCEL_EPOCH = Date.UTC(1899, 11, 30)
export const dateToSerial = (d: Date) => (d.getTime() - EXCEL_EPOCH) / 86_400_000
const serialToDate = (n: number) => new Date(EXCEL_EPOCH + Math.round(n * 86_400_000))

/** Whether a number format shows a date or time (so a computed serial number should display as one). */
function isDateFormat(numFmt?: string): boolean {
  if (!numFmt) return false
  const f = numFmt.replace(/"[^"]*"|\[[^\]]*\]|\\./g, '')
  return /[dmyhs]/i.test(f) && !/[#0]/.test(f.replace(/s\.0+/i, ''))
}

export function cellData(cell: Cell, formulaResult?: FormulaResult): CellData | undefined {
  if (isMergeSlave(cell)) return undefined
  const raw = cell.value
  const formula = isFormula(raw)
  let v = formula ? (raw.result as CellValue) : raw
  if (formula) {
    const live = formulaResult?.(cell)
    if (live !== undefined) v = typeof live === 'number' && isDateFormat(cell.numFmt) ? serialToDate(live) : live
  }
  if (v === null || v === undefined) {
    return formula && 'formula' in raw ? { text: `=${raw.formula}`, sort: '', kind: 'text', formula } : undefined
  }
  if (typeof v === 'number') return { text: formatWithNumFmt(v, cell.numFmt), sort: v, kind: 'number', formula }
  if (v instanceof Date) return { text: formatDate(v), sort: v.getTime(), kind: 'date', formula }
  if (typeof v === 'boolean') return { text: v ? 'TRUE' : 'FALSE', sort: v ? 1 : 0, kind: 'boolean', formula }
  const text = displayText(v)
  return text ? { text, sort: text.toLocaleLowerCase('th'), kind: 'text', formula } : undefined
}

/** Numbers (and dates) before text; text in Thai dictionary order. */
export function compareSort(a: unknown, b: unknown): number {
  const an = typeof a === 'number'
  const bn = typeof b === 'number'
  if (an && bn) return (a as number) - (b as number)
  if (an !== bn) return an ? -1 : 1
  return String(a).localeCompare(String(b), 'th')
}

const NUM = /^-?(\d+\.?\d*|\.\d+)(e[-+]?\d+)?$/i
const NUM_GROUPED = /^-?\d{1,3}(,\d{3})+(\.\d+)?$/
const LEADING_ZERO = /^-?0\d/
const DATE = /^(\d{4})-(\d{1,2})-(\d{1,2})(?:[ T](\d{1,2}):(\d{2})(?::(\d{2}))?)?$/

/**
 * Turn what the user typed into an Excel value, roughly the way Excel does:
 * numbers, TRUE/FALSE, yyyy-mm-dd dates and =formulas are converted; anything
 * else stays text. Values with leading zeros (phone numbers, codes) stay text,
 * a leading apostrophe forces text, and cells formatted as Text (`@`) stay text.
 */
export function parseInput(raw: string, prev: CellValue = null, numFmt?: string): CellValue {
  if (raw === '') return null
  if (raw.startsWith("'")) return raw.slice(1)
  if (raw.length > 1 && raw.startsWith('=')) return { formula: raw.slice(1) } as CellValue
  if (numFmt === '@') return keepHyperlink(raw, prev)

  const t = raw.trim()
  if (NUM.test(t) && !LEADING_ZERO.test(t)) return Number(t)
  if (NUM_GROUPED.test(t)) return Number(t.replace(/,/g, ''))

  const upper = t.toUpperCase()
  if (upper === 'TRUE' || upper === 'FALSE') return upper === 'TRUE'

  const m = DATE.exec(t)
  if (m) {
    const [y, mo, d, h = '0', mi = '0', s = '0'] = m.slice(1)
    const date = new Date(Date.UTC(+y, +mo - 1, +d, +h, +mi, +s))
    if (date.getUTCMonth() === +mo - 1 && date.getUTCDate() === +d) return date
  }
  return keepHyperlink(raw, prev)
}

function keepHyperlink(text: string, prev: CellValue): CellValue {
  if (prev && typeof prev === 'object' && !(prev instanceof Date) && 'hyperlink' in prev) {
    return { ...prev, text }
  }
  return text
}

// ---- spreadsheet (Excel view) rendering ------------------------------------------

/** What the Excel-style grid needs to draw one cell, including its Excel formatting. */
export interface CellView {
  text: string
  /** Inline CSS from the cell's font and fill. */
  style?: string
  align?: 'left' | 'center' | 'right'
}

// Default Office theme colours, in the order Excel's `theme` index uses.
const THEME = ['FFFFFF', '000000', 'E7E6E6', '44546A', '4472C4', 'ED7D31', 'A5A5A5', 'FFC000', '5B9BD5', '70AD47']

type ExcelColor = { argb?: string; theme?: number; tint?: number }

function colorToCss(color?: ExcelColor): string | undefined {
  if (!color) return undefined
  let hex: string | undefined
  if (color.argb && /^[0-9a-f]{6,8}$/i.test(color.argb)) hex = color.argb.slice(-6).toUpperCase()
  else if (typeof color.theme === 'number') hex = THEME[color.theme]
  if (!hex) return undefined
  const tint = color.tint ?? 0
  if (!tint) return `#${hex}`
  const rgb = [0, 2, 4].map((i) => parseInt(hex.slice(i, i + 2), 16))
  return `rgb(${rgb.map((v) => Math.round(tint > 0 ? v + (255 - v) * tint : v * (1 + tint))).join(',')})`
}

function cssOf(cell: Cell): string | undefined {
  const parts: string[] = []
  const font = cell.font
  if (font?.bold) parts.push('font-weight:600')
  if (font?.italic) parts.push('font-style:italic')
  const deco = [font?.underline && 'underline', font?.strike && 'line-through'].filter(Boolean)
  if (deco.length) parts.push(`text-decoration:${deco.join(' ')}`)

  const fill = cell.fill
  const bg = fill?.type === 'pattern' && fill.pattern === 'solid' ? colorToCss(fill.fgColor as ExcelColor) : undefined
  let color = colorToCss(font?.color as ExcelColor)
  // Excel's default text is black. Without a fill we use the theme colour instead (so
  // dark mode stays readable); on a filled cell we keep black, as Excel shows it.
  if (bg && !color) color = '#000000'
  if (!bg && color === '#000000') color = undefined
  if (color) parts.push(`color:${color}`)
  if (bg) parts.push(`background-color:${bg}`)
  return parts.length ? parts.join(';') : undefined
}

function alignOf(cell: Cell, data: CellData | undefined): CellView['align'] {
  const h = cell.alignment?.horizontal
  if (h === 'left' || h === 'center' || h === 'right') return h
  if (data?.kind === 'number' || data?.kind === 'date') return 'right'
  if (data?.kind === 'boolean') return 'center'
  return undefined
}

/** Grid view of a cell (text formatted like Excel shows it), or undefined if there is nothing to draw. */
export function viewOf(cell: Cell, formulaResult?: FormulaResult): CellView | undefined {
  if (isMergeSlave(cell)) return undefined
  const data = cellData(cell, formulaResult)
  const style = cssOf(cell)
  if (!data && !style) return undefined
  return { text: data?.text ?? '', style, align: alignOf(cell, data) }
}

export const address = (r: number, c: number) => `${colName(c)}${r}`
