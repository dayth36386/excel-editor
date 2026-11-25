export const XLSX_MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'

const types: FilePickerAcceptType[] = [{ description: 'Excel Workbook (.xlsx)', accept: { [XLSX_MIME]: ['.xlsx'] } }]

/**
 * Chrome and Edge can open a file and later write straight back to it
 * (File System Access API). Other browsers fall back to <input type=file>
 * for opening and a download for saving.
 */
export const canWriteBack = typeof window !== 'undefined' && typeof window.showOpenFilePicker === 'function'
export const canSaveAs = typeof window !== 'undefined' && typeof window.showSaveFilePicker === 'function'

const isAbort = (e: unknown) => e instanceof DOMException && e.name === 'AbortError'

export async function pickExcelFile(): Promise<FileSystemFileHandle | null> {
  try {
    const [handle] = await window.showOpenFilePicker!({ types, multiple: false })
    return handle ?? null
  } catch (e) {
    if (isAbort(e)) return null
    throw e
  }
}

export async function pickSaveTarget(suggestedName: string): Promise<FileSystemFileHandle | null> {
  try {
    return await window.showSaveFilePicker!({ types, suggestedName })
  } catch (e) {
    if (isAbort(e)) return null
    throw e
  }
}

/** The user (or browser) refused write access to the opened file. */
export class WritePermissionError extends Error {
  constructor() {
    super('Write permission was not granted')
    this.name = 'WritePermissionError'
  }
}

/** Ask for write access to a file we opened (Chrome asks the user once). */
export async function ensureWritable(handle: FileSystemFileHandle) {
  const opts = { mode: 'readwrite' } as const
  if (!handle.requestPermission) return
  if ((await handle.queryPermission?.(opts)) === 'granted') return
  if ((await handle.requestPermission(opts)) !== 'granted') throw new WritePermissionError()
}

export async function writeToHandle(handle: FileSystemFileHandle, blob: Blob) {
  const writable = await handle.createWritable()
  await writable.write(blob)
  await writable.close()
}

export function downloadBlob(blob: Blob, name: string) {
  const url = URL.createObjectURL(blob)
  const a = document.createElement('a')
  a.href = url
  a.download = name
  document.body.append(a)
  a.click()
  a.remove()
  setTimeout(() => URL.revokeObjectURL(url), 1000)
}

/**
 * Read a dropped file. In Chrome/Edge this also returns a handle, so a dropped
 * file can be saved back in place too. Must be called synchronously from the
 * drop handler (before any await), because the DataTransfer is cleared after.
 */
export async function fileFromDrop(e: DragEvent): Promise<{ file: File; handle: FileSystemFileHandle | null } | null> {
  const item = e.dataTransfer?.items?.[0]
  const pending = item?.kind === 'file' && item.getAsFileSystemHandle ? item.getAsFileSystemHandle() : null
  const file = e.dataTransfer?.files?.[0]
  if (!file) return null
  const handle = await pending?.catch(() => null)
  if (handle && handle.kind === 'file') return { file, handle: handle as FileSystemFileHandle }
  return { file, handle: null }
}
