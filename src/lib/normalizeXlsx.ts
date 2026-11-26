import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate'

const SHEET_RELS = /^xl\/worksheets\/_rels\/[^/]+\.xml\.rels$/
const SHEET_DIR = 'xl/worksheets/'

/** A relationship target (absolute "/xl/…" or relative to xl/worksheets/) → its path inside the zip. */
function toZipPath(target: string): string {
  if (target.startsWith('/')) return target.slice(1)
  return new URL(target, `file:///${SHEET_DIR}`).pathname.slice(1)
}

/** A path inside the zip → a target relative to xl/worksheets/, the form ExcelJS looks up. */
function toSheetTarget(zipPath: string): string {
  if (zipPath.startsWith(SHEET_DIR)) return zipPath.slice(SHEET_DIR.length)
  return zipPath.startsWith('xl/') ? `../${zipPath.slice(3)}` : `/${zipPath}`
}

/**
 * ExcelJS only finds a sheet's comments if they are stored as `xl/commentsN.xml` and
 * `xl/drawings/vmlDrawingN.vml`, referenced with relative paths, which is how Excel writes them.
 * Other writers (openpyxl, for one) use `xl/comments/comment1.xml`, `commentsDrawing1.vml` and
 * absolute "/xl/…" targets, and ExcelJS then fails with
 * "Cannot read properties of undefined (reading 'comments')".
 *
 * This renames those parts and rewrites the sheet relationships to the Excel layout. Nothing else
 * in the file changes, and a file that is already in that layout is returned as is.
 */
export function normalizeXlsx(data: ArrayBuffer): ArrayBuffer {
  const zip = new Uint8Array(data)
  // Only the small .rels files are decompressed to decide whether anything needs fixing.
  const rels = unzipSync(zip, { filter: (f) => SHEET_RELS.test(f.name) })
  const needsFix = Object.values(rels).some((u8) =>
    /Target="\/|Target="[^"]*comments\/|commentsDrawing/.test(strFromU8(u8)),
  )
  if (!needsFix) return data

  const files = unzipSync(zip)
  const taken = new Set(Object.keys(files))
  const freshName = (prefix: string, ext: string) => {
    let n = 1
    while (taken.has(`${prefix}${n}${ext}`)) n++
    const name = `${prefix}${n}${ext}`
    taken.add(name)
    return name
  }

  for (const relsPath of Object.keys(files).filter((p) => SHEET_RELS.test(p))) {
    const xml = strFromU8(files[relsPath])
    const fixed = xml.replace(/<Relationship\b[^>]*>/g, (tag) => {
      if (/TargetMode="External"/.test(tag)) return tag
      const target = /Target="([^"]*)"/.exec(tag)?.[1]
      const type = /Type="([^"]*)"/.exec(tag)?.[1] ?? ''
      if (!target) return tag

      const from = toZipPath(target)
      let to = from
      if (type.endsWith('/comments') && !/^xl\/comments\d+\.xml$/.test(from)) {
        to = freshName('xl/comments', '.xml')
      } else if (type.endsWith('/vmlDrawing') && !/^xl\/drawings\/vmlDrawing\d+\.vml$/.test(from)) {
        to = freshName('xl/drawings/vmlDrawing', '.vml')
      }
      if (to !== from && files[from]) {
        files[to] = files[from]
        delete files[from]
      }
      const next = toSheetTarget(to)
      return next === target ? tag : tag.replace(/Target="[^"]*"/, `Target="${next}"`)
    })
    if (fixed !== xml) files[relsPath] = strToU8(fixed)
  }
  const out = zipSync(files)
  return out.buffer.slice(out.byteOffset, out.byteOffset + out.byteLength) as ArrayBuffer
}
