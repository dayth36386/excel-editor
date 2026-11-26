import { describe, expect, it } from 'vitest'
import { evaluateFormula, type Scalar } from './formula'

// A1..C3 grid:   1  2  x
//                3  4  TRUE
//                (blank) 5 "7"
const grid: Record<string, Scalar> = {
  '1:1': 1,
  '1:2': 2,
  '1:3': 'x',
  '2:1': 3,
  '2:2': 4,
  '2:3': true,
  '3:2': 5,
  '3:3': '7',
}
const get = (r: number, c: number) => grid[`${r}:${c}`] ?? null
const ev = (f: string) => evaluateFormula(f, get, 3)

describe('evaluateFormula', () => {
  it('sums, averages and counts ranges like Excel (numbers only)', () => {
    expect(ev('SUM(A1:C3)')).toBe(15)
    expect(ev('sum(a1:b2)')).toBe(10)
    expect(ev('SUM(A1, B1, 10)')).toBe(13)
    expect(ev('AVERAGE(A1:B2)')).toBe(2.5)
    expect(ev('MIN(A1:B3)')).toBe(1)
    expect(ev('MAX(A1:B3)')).toBe(5)
    expect(ev('COUNT(A1:C3)')).toBe(5)
    expect(ev('COUNTA(A1:C3)')).toBe(8)
    expect(ev('SUM(B:B)')).toBe(11)
    expect(ev('SUM($A$1:$A$2)')).toBe(4)
  })

  it('does arithmetic with precedence, parentheses and text joins', () => {
    expect(ev('A1+B1*B2')).toBe(9)
    expect(ev('(A1+B1)*B2')).toBe(12)
    expect(ev('-A2^2')).toBe(9) // Excel applies the minus before ^
    expect(ev('2^3^2')).toBe(512)
    expect(ev('C3*2')).toBe(14)
    expect(ev('ROUND(B3/3, 2)')).toBe(1.67)
    expect(ev('ABS(-A2)')).toBe(3)
    expect(ev('C1&"-"&A1')).toBe('x-1')
    expect(ev('A3+1')).toBe(1)
  })

  it('gives up (undefined) instead of guessing', () => {
    expect(ev('VLOOKUP(A1, A1:B3, 2, FALSE)')).toBeUndefined()
    expect(ev("SUM('Other sheet'!A1:A3)")).toBeUndefined()
    expect(ev('A1/0')).toBeUndefined()
    expect(ev('AVERAGE(C1)')).toBeUndefined()
    expect(ev('C1*2')).toBeUndefined()
    expect(ev('A1>2')).toBeUndefined()
    expect(ev('SUM(A1:B2')).toBeUndefined()
  })
})
