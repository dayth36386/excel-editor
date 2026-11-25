/**
 * Parse tab-separated text as Excel and Google Sheets put it on the clipboard.
 * Cells containing tabs, newlines or quotes arrive wrapped in double quotes.
 */
export function parseTSV(input: string): string[][] {
  let text = input.replace(/\r\n?/g, '\n')
  if (text.endsWith('\n')) text = text.slice(0, -1)

  const rows: string[][] = []
  let row: string[] = []
  let field = ''
  let quoted = false

  for (let i = 0; i < text.length; i++) {
    const ch = text[i]
    if (quoted) {
      if (ch !== '"') field += ch
      else if (text[i + 1] === '"') {
        field += '"'
        i++
      } else quoted = false
    } else if (ch === '"' && field === '') quoted = true
    else if (ch === '\t') {
      row.push(field)
      field = ''
    } else if (ch === '\n') {
      row.push(field)
      rows.push(row)
      row = []
      field = ''
    } else field += ch
  }
  row.push(field)
  rows.push(row)
  return rows
}

/** Quote a value for the clipboard the way Excel does. */
export function toTSVCell(text: string): string {
  return /[\t\n"]/.test(text) ? `"${text.replace(/"/g, '""')}"` : text
}
