/// <reference types="vite/client" />

// File System Access API (Chrome/Edge). Not in TypeScript's DOM lib yet.
interface FilePickerAcceptType {
  description?: string
  accept: Record<string, string[]>
}

interface Window {
  showOpenFilePicker?(options?: {
    types?: FilePickerAcceptType[]
    multiple?: boolean
    excludeAcceptAllOption?: boolean
  }): Promise<FileSystemFileHandle[]>
  showSaveFilePicker?(options?: {
    types?: FilePickerAcceptType[]
    suggestedName?: string
  }): Promise<FileSystemFileHandle>
}

interface DataTransferItem {
  getAsFileSystemHandle?(): Promise<FileSystemHandle | null>
}

interface FileSystemHandle {
  queryPermission?(descriptor: { mode: 'read' | 'readwrite' }): Promise<PermissionState>
  requestPermission?(descriptor: { mode: 'read' | 'readwrite' }): Promise<PermissionState>
}
