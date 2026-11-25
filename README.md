# Excel Editor

เปิดไฟล์ Excel (.xlsx) มาแก้ในตารางสไตล์ shadcn แล้วบันทึกกลับเป็น Excel — ทำงานในเบราว์เซอร์ทั้งหมด ไฟล์ไม่ถูกอัปโหลดไปไหน

**Stack:** Vue 3 · Vite · TypeScript · Tailwind v4 · shadcn-vue (reka-ui) · TanStack Table v8 · ExcelJS · vue-i18n

- ภาษาไทย / English สลับได้ที่ปุ่ม TH | EN มุมขวาบน (จำค่าไว้ในเบราว์เซอร์; ครั้งแรกเลือกตามภาษาเบราว์เซอร์)
- ธีมสว่าง / มืด / ตามระบบ
- ข้อความทั้งหมดอยู่ใน `src/i18n/en.ts` และ `src/i18n/th.ts` (TypeScript เช็คว่า key ตรงกัน)

```bash
npm install
npm run dev      # http://localhost:5173
npm test         # unit tests (อ่าน/เขียน/แทรก/ลบ + round-trip ไฟล์)
npm run build    # type-check + production build
npm run sample   # สร้าง public/sample.xlsx ใหม่
npm run lint     # ESLint (lint:fix เพื่อแก้อัตโนมัติ)
npm run format   # Prettier จัดทุกไฟล์ (format:check แค่ตรวจ)
npm run check    # typecheck + lint + format:check + test รวดเดียว
```

### Code style

- **Prettier** — ไม่มี `;`, single quote, กว้าง 120, เรียง class Tailwind อัตโนมัติ (`prettier-plugin-tailwindcss`)
- **ESLint** — `eslint-plugin-vue` (recommended) + `typescript-eslint`; กฎเรื่อง format ปิดไว้ให้ Prettier ดูแล
- **Husky**
  - `pre-commit` → `lint-staged`: ESLint `--fix` + Prettier เฉพาะไฟล์ที่ stage
  - `pre-push` → `typecheck` + `test`
- `src/components/ui/` (shadcn-vue) ไม่ถูก lint/format เพื่อให้ตรงกับที่ CLI generate

## ใช้งาน

มี 2 มุมมอง สลับได้ที่ปุ่ม **ตาราง | Excel** (แก้ข้อมูลชุดเดียวกัน, undo ใช้ร่วมกัน, จำมุมมองล่าสุดไว้)

- **ตาราง** — data table สไตล์ shadcn: ค้นหา กรอง เรียง แบ่งหน้า เลือกหลายแถว
- **Excel** — ตารางแบบสเปรดชีต: คอลัมน์ A B C, เลขแถว, สี/ฟอนต์จากไฟล์, freeze pane, แถบสูตร fx,
  เลื่อนด้วยลูกศร, พิมพ์ทับได้เลย, Enter/F2 แก้, Alt+Enter ขึ้นบรรทัดใหม่, Del ล้าง,
  ⌘/Ctrl+C คัดลอก, ⌘/Ctrl+V วางหลายเซลล์จาก Excel/Google Sheets, คลิกขวาเพื่อแทรก/ลบแถว-คอลัมน์

- **เปิดไฟล์**: ปุ่ม "เปิดไฟล์", ลากไฟล์มาวาง, หรือ ⌘/Ctrl+O
- แถวแรกที่มีข้อมูลของชีตคือ **หัวตาราง**; แถวที่อยู่ข้างล่างคือข้อมูล
- **แก้เซลล์**: คลิกเซลล์ → พิมพ์ → Enter (Tab ไปช่องถัดไป, Esc ยกเลิก)
  - พิมพ์ `=SUM(E2:E10)` เพื่อใส่สูตร, `'0123` เพื่อบังคับเป็นข้อความ
  - เลขที่ขึ้นต้นด้วย 0 (เบอร์โทร/รหัส) เก็บเป็นข้อความให้อัตโนมัติ
  - คอลัมน์ TRUE/FALSE แสดงเป็น checkbox
- ค้นหาทุกคอลัมน์, เรียงตามคอลัมน์, ซ่อนคอลัมน์, แบ่งหน้า, เลือกหลายแถวแล้วลบ
- คอลัมน์ข้อความที่มีค่าซ้ำไม่กี่แบบ (เช่น ระดับ/สถานะ) แสดงเป็น badge สี มีตัวกรองพร้อมจำนวน และ autocomplete ตอนพิมพ์
- แทรก/ลบแถว (เมนู ⋯ ท้ายแถว), เพิ่ม/เปลี่ยนชื่อ/ลบคอลัมน์ (เมนูที่หัวคอลัมน์), เพิ่ม/เปลี่ยนชื่อชีต
- เลิกทำ/ทำซ้ำ: ⌘/Ctrl+Z, ⌘/Ctrl+Shift+Z
- **บันทึก** (⌘/Ctrl+S): Chrome/Edge เขียนทับไฟล์เดิมได้เลย; เบราว์เซอร์อื่นดาวน์โหลดเป็นไฟล์ใหม่

## สิ่งที่คงไว้ตอนบันทึก

แก้ไขลงใน workbook เดิมโดยตรง จึงคงไว้: สี/ฟอนต์/เส้นขอบ, รูปแบบตัวเลข, ความกว้างคอลัมน์, freeze pane, สูตร, ชีตอื่นๆ
สูตรไม่ถูกคำนวณใหม่ในแอป — ไฟล์ถูกตั้งให้ Excel คำนวณใหม่ตอนเปิด

## ข้อจำกัด (จาก ExcelJS)

- รองรับเฉพาะ `.xlsx` (ไม่รองรับ `.xls` รุ่นเก่า, `.xlsm` จะเสีย macro)
- กราฟ (charts) และ pivot table จะหายเมื่อบันทึก
- แทรก/ลบแถว-คอลัมน์ไม่ได้ปรับ reference ในสูตรให้อัตโนมัติ

# excel-editor
