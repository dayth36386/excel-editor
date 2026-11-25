// Builds public/sample.xlsx — a made-up member list used by "ลองเปิดไฟล์ตัวอย่าง".
// Run with: npm run sample
import ExcelJS from 'exceljs'
import { fileURLToPath } from 'node:url'

const out = fileURLToPath(new URL('../public/sample.xlsx', import.meta.url))

// Small deterministic random generator so the file is the same every run.
let seed = 42
const rand = () => (seed = (seed * 1103515245 + 12345) % 2 ** 31) / 2 ** 31
const pick = (list) => list[Math.floor(rand() * list.length)]

const first = ['สมชาย', 'สมหญิง', 'วิชัย', 'ปราณี', 'อนุชา', 'กาญจนา', 'ธนพล', 'ศิริพร', 'ณัฐวุฒิ', 'พิมพ์ชนก', 'กิตติ', 'อรุณี']
const last = ['ใจดี', 'รักไทย', 'ศรีสุข', 'ทองมา', 'มั่นคง', 'แสงทอง', 'บุญมี', 'วงศ์ใหญ่', 'สุขสันต์', 'พรหมมา']
const tiers = ['Silver', 'Gold', 'Platinum']

const wb = new ExcelJS.Workbook()
wb.creator = 'excel-editor sample'

const ws = wb.addWorksheet('สมาชิก', { views: [{ state: 'frozen', ySplit: 1, xSplit: 0, topLeftCell: 'A2' }] })
ws.columns = [
  { header: 'รหัสสมาชิก', key: 'id', width: 13 },
  { header: 'ชื่อ-นามสกุล', key: 'name', width: 24 },
  { header: 'เบอร์โทร', key: 'phone', width: 14, style: { numFmt: '@' } },
  { header: 'ระดับ', key: 'tier', width: 11 },
  { header: 'แต้มสะสม', key: 'points', width: 12, style: { numFmt: '#,##0' } },
  { header: 'ยอดซื้อรวม (บาท)', key: 'spend', width: 17, style: { numFmt: '#,##0.00' } },
  { header: 'วันที่สมัคร', key: 'joined', width: 13, style: { numFmt: 'yyyy-mm-dd' } },
  { header: 'ใช้งานอยู่', key: 'active', width: 10 },
]

const N = 150
let totalPoints = 0
let totalSpend = 0
for (let i = 1; i <= N; i++) {
  const points = Math.floor(rand() * 20000)
  const spend = Math.round(rand() * 5_000_000) / 100
  totalPoints += points
  totalSpend += spend
  ws.addRow({
    id: `M${String(i).padStart(5, '0')}`,
    name: `${pick(first)} ${pick(last)}`,
    phone: `08${Math.floor(rand() * 1e8)
      .toString()
      .padStart(8, '0')}`,
    tier: pick(tiers),
    points,
    spend,
    joined: new Date(Date.UTC(2022 + Math.floor(rand() * 4), Math.floor(rand() * 12), 1 + Math.floor(rand() * 28))),
    active: rand() > 0.2,
  })
}

const header = ws.getRow(1)
header.font = { bold: true, color: { argb: 'FFFFFFFF' } }
header.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF107C41' } }
header.alignment = { vertical: 'middle' }

const totalRow = ws.addRow({ name: 'รวม' })
totalRow.getCell('points').value = { formula: `SUM(E2:E${N + 1})`, result: totalPoints }
totalRow.getCell('spend').value = { formula: `SUM(F2:F${N + 1})`, result: Math.round(totalSpend * 100) / 100 }
totalRow.font = { bold: true }
totalRow.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE3F1E8' } }

// Second sheet: formulas that point at the first.
const sum = wb.addWorksheet('สรุปตามระดับ')
sum.columns = [
  { header: 'ระดับ', width: 12 },
  { header: 'จำนวนสมาชิก', width: 15 },
  { header: 'แต้มรวม', width: 14, style: { numFmt: '#,##0' } },
]
sum.getRow(1).font = { bold: true }
const rows = ws.getRows(2, N)
tiers.forEach((tier, i) => {
  const mine = rows.filter((r) => r.getCell(4).value === tier)
  const r = i + 2
  sum.getCell(r, 1).value = tier
  sum.getCell(r, 2).value = { formula: `COUNTIF('สมาชิก'!D:D,A${r})`, result: mine.length }
  sum.getCell(r, 3).value = {
    formula: `SUMIF('สมาชิก'!D:D,A${r},'สมาชิก'!E:E)`,
    result: mine.reduce((s, row) => s + row.getCell(5).value, 0),
  }
})

await wb.xlsx.writeFile(out)
console.log(`wrote ${out}`)
