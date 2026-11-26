/**
 * A small formula evaluator for the formulas people use most in simple sheets:
 * arithmetic (+ - * / ^, parentheses, unary minus, & for text), cell references,
 * ranges on the same sheet (A1:B5, A:A) and SUM, AVERAGE, MIN, MAX, COUNT, COUNTA,
 * ROUND, ABS.
 *
 * Anything else (other functions, references to other sheets, comparisons…) makes
 * `evaluateFormula` return `undefined`, and the caller falls back to the value Excel
 * cached in the file. So a formula is either computed correctly or not at all.
 */

export type Scalar = number | string | boolean | null

/** Value of a cell; rows and columns are 1-based. */
export type CellGetter = (r: number, c: number) => Scalar

class Unsupported extends Error {}

type Token =
  | { t: 'num'; v: number }
  | { t: 'str'; v: string }
  | { t: 'ref'; r: number; c: number }
  | { t: 'range'; r1: number; c1: number; r2: number; c2: number }
  | { t: 'fn'; name: string }
  | { t: 'op'; v: string }

const colNumber = (letters: string) => [...letters.toUpperCase()].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0)

const CELL = String.raw`\$?([A-Za-z]{1,3})\$?(\d+)`
const COL = String.raw`\$?([A-Za-z]{1,3})`
const TOKEN = new RegExp(
  [
    String.raw`(?<space>\s+)`,
    String.raw`(?<num>\d+(?:\.\d+)?(?:[eE][-+]?\d+)?|\.\d+)`,
    String.raw`(?<str>"(?:[^"]|"")*")`,
    String.raw`(?<sheet>(?:'[^']*'|[A-Za-z_][\w.]*)!)`,
    `(?<range>${CELL}:${CELL})`,
    `(?<cols>${COL}:${COL})(?![\\d(])`,
    String.raw`(?<fn>[A-Za-z][A-Za-z0-9.]*)\s*\(`,
    `(?<ref>${CELL})(?![\\w(])`,
    String.raw`(?<op>[-+*/^&(),])`,
  ].join('|'),
  'y',
)

function tokenize(src: string, maxRow: number): Token[] {
  const out: Token[] = []
  TOKEN.lastIndex = 0
  while (TOKEN.lastIndex < src.length) {
    const m = TOKEN.exec(src)
    if (!m?.groups) throw new Unsupported(src)
    const g = m.groups
    if (g.space) continue
    if (g.sheet) throw new Unsupported('other sheet')
    if (g.num) out.push({ t: 'num', v: Number(g.num) })
    else if (g.str) out.push({ t: 'str', v: g.str.slice(1, -1).replace(/""/g, '"') })
    else if (g.range) {
      const [a, b] = g.range.split(':').map((s) => /\$?([A-Za-z]+)\$?(\d+)/.exec(s)!)
      out.push({
        t: 'range',
        r1: Number(a[2]),
        c1: colNumber(a[1]),
        r2: Number(b[2]),
        c2: colNumber(b[1]),
      })
    } else if (g.cols) {
      const [a, b] = g.cols.replace(/\$/g, '').split(':')
      out.push({ t: 'range', r1: 1, c1: colNumber(a), r2: maxRow, c2: colNumber(b) })
    } else if (g.fn) out.push({ t: 'fn', name: g.fn.toUpperCase() })
    else if (g.ref) {
      const [, letters, row] = /\$?([A-Za-z]+)\$?(\d+)/.exec(g.ref)!
      out.push({ t: 'ref', r: Number(row), c: colNumber(letters) })
    } else out.push({ t: 'op', v: g.op })
  }
  return out
}

type Value = Scalar | Scalar[]
type ErrorValue = '#DIV/0!' | '#VALUE!'

const MAX_RANGE_CELLS = 200_000

function toNumber(v: Scalar): number {
  if (v === null || v === '') return 0
  if (typeof v === 'number') return v
  if (typeof v === 'boolean') return v ? 1 : 0
  const n = Number(v)
  if (Number.isNaN(n)) throw new Unsupported('#VALUE!' satisfies ErrorValue)
  return n
}

const toText = (v: Scalar) => (v === null ? '' : typeof v === 'boolean' ? (v ? 'TRUE' : 'FALSE') : String(v))

/** Recursive-descent parser that evaluates as it goes. */
class Evaluator {
  private i = 0
  constructor(
    private tokens: Token[],
    private get: CellGetter,
  ) {}

  run(): Scalar {
    const v = this.concat()
    if (this.i < this.tokens.length) throw new Unsupported('trailing tokens')
    return this.scalar(v)
  }

  private peek() {
    return this.tokens[this.i]
  }

  private isOp(v: string) {
    const t = this.peek()
    return t?.t === 'op' && t.v === v
  }

  private expect(v: string) {
    if (!this.isOp(v)) throw new Unsupported(`expected ${v}`)
    this.i++
  }

  private scalar(v: Value): Scalar {
    if (Array.isArray(v)) throw new Unsupported('range used as a value')
    return v
  }

  private concat(): Value {
    let left = this.additive()
    while (this.isOp('&')) {
      this.i++
      left = toText(this.scalar(left)) + toText(this.scalar(this.additive()))
    }
    return left
  }

  private additive(): Value {
    let left = this.term()
    while (this.isOp('+') || this.isOp('-')) {
      const op = (this.tokens[this.i++] as { v: string }).v
      const a = toNumber(this.scalar(left))
      const b = toNumber(this.scalar(this.term()))
      left = op === '+' ? a + b : a - b
    }
    return left
  }

  private term(): Value {
    let left = this.power()
    while (this.isOp('*') || this.isOp('/')) {
      const op = (this.tokens[this.i++] as { v: string }).v
      const a = toNumber(this.scalar(left))
      const b = toNumber(this.scalar(this.power()))
      if (op === '/' && b === 0) throw new Unsupported('#DIV/0!' satisfies ErrorValue)
      left = op === '*' ? a * b : a / b
    }
    return left
  }

  private power(): Value {
    const base = this.unary()
    if (!this.isOp('^')) return base
    this.i++
    return toNumber(this.scalar(base)) ** toNumber(this.scalar(this.power()))
  }

  private unary(): Value {
    if (this.isOp('-')) {
      this.i++
      return -toNumber(this.scalar(this.unary()))
    }
    if (this.isOp('+')) {
      this.i++
      return this.unary()
    }
    return this.primary()
  }

  private primary(): Value {
    const t = this.tokens[this.i++]
    if (!t) throw new Unsupported('unexpected end')
    switch (t.t) {
      case 'num':
      case 'str':
        return t.v
      case 'ref':
        return this.get(t.r, t.c)
      case 'range':
        return this.range(t)
      case 'fn':
        return this.call(t.name)
      case 'op':
        if (t.v === '(') {
          const v = this.concat()
          this.expect(')')
          return v
        }
    }
    throw new Unsupported('unexpected token')
  }

  private range(t: Extract<Token, { t: 'range' }>): Scalar[] {
    const r1 = Math.min(t.r1, t.r2)
    const r2 = Math.max(t.r1, t.r2)
    const c1 = Math.min(t.c1, t.c2)
    const c2 = Math.max(t.c1, t.c2)
    if ((r2 - r1 + 1) * (c2 - c1 + 1) > MAX_RANGE_CELLS) throw new Unsupported('range too large')
    const out: Scalar[] = []
    for (let r = r1; r <= r2; r++) for (let c = c1; c <= c2; c++) out.push(this.get(r, c))
    return out
  }

  private args(): Value[] {
    const list: Value[] = []
    if (this.isOp(')')) {
      this.i++
      return list
    }
    for (;;) {
      list.push(this.concat())
      if (this.isOp(',')) {
        this.i++
        continue
      }
      this.expect(')')
      return list
    }
  }

  private call(name: string): Scalar {
    const args = this.args()
    // Ranges contribute only their numbers; values typed directly into the call are converted.
    const numbers = () =>
      args.flatMap((a) => (Array.isArray(a) ? a.filter((v): v is number => typeof v === 'number') : [toNumber(a)]))
    switch (name) {
      case 'SUM':
        return numbers().reduce((s, n) => s + n, 0)
      case 'AVERAGE': {
        const ns = numbers()
        if (!ns.length) throw new Unsupported('#DIV/0!' satisfies ErrorValue)
        return ns.reduce((s, n) => s + n, 0) / ns.length
      }
      case 'MIN': {
        const ns = numbers()
        return ns.length ? Math.min(...ns) : 0
      }
      case 'MAX': {
        const ns = numbers()
        return ns.length ? Math.max(...ns) : 0
      }
      case 'COUNT':
        return args.flatMap((a) => (Array.isArray(a) ? a : [a])).filter((v) => typeof v === 'number').length
      case 'COUNTA':
        return args.flatMap((a) => (Array.isArray(a) ? a : [a])).filter((v) => v !== null && v !== '').length
      case 'ROUND': {
        const [x, digits = 0] = args.map((a) => toNumber(this.scalar(a)))
        const f = 10 ** digits
        return Math.round(x * f) / f
      }
      case 'ABS':
        return Math.abs(toNumber(this.scalar(args[0] ?? null)))
    }
    throw new Unsupported(`function ${name}`)
  }
}

/**
 * Evaluate `formula` (without the leading "="). Returns undefined when the formula uses
 * something this evaluator doesn't support, or produces an error such as #DIV/0!.
 * `maxRow` bounds whole-column ranges like A:A.
 */
export function evaluateFormula(formula: string, get: CellGetter, maxRow = 1_048_576): Scalar | undefined {
  try {
    const result = new Evaluator(tokenize(formula, maxRow), get).run()
    return typeof result === 'number' && !Number.isFinite(result) ? undefined : result
  } catch (e) {
    if (e instanceof Unsupported) return undefined
    throw e
  }
}
