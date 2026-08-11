// 渲染进程暴露给 React 使用的 Electron IPC API 类型声明
export interface PreviewRow {
  cells: Array<{ value: string; label?: string }>
}

// 最近修改历史记录
export interface HistoryRecord {
  filePath: string
  fileName: string
  time: string
}

export interface ElectronAPI {
  selectDirectory: () => Promise<string | null>
  getWorkDir: () => Promise<string>
  scanFiles: (dir: string) => Promise<{ error?: boolean; message?: string; files?: FileEntry[] }>
  openFile: (filePath: string) => Promise<{ error?: boolean; message?: string }>
  openDir: (dirPath: string) => Promise<{ error?: boolean; message?: string }>
  previewRows: (filePath: string, limit: number) => Promise<{
    error?: boolean
    message?: string
    sheetName?: string
    totalRows?: number
    headerLabels?: string[]
    rows?: PreviewRow[]
  }>
  appendMemo: (filePath: string, rows: Array<{
    no: string
    date: string
    name: string
    unit: string
    qty: number
    price: number
    amount: number | ''
    person: string
    remark: string
  }>) => Promise<{
    error?: boolean
    message?: string
    rowNumber?: number
    count?: number
  }>
  updateMemo: (filePath: string, rows: Array<{
    no: string
    date: string
    name: string
    unit: string
    qty: number
    price: number
    amount: number | ''
    person: string
    remark: string
  }>) => Promise<{
    error?: boolean
    message?: string
    rowNumber?: number
    count?: number
  }>
  getHistory: () => Promise<HistoryRecord[]>
  clearHistory: () => Promise<void>
}

export interface FileEntry {
  fileName: string
  filePath: string
  fileNameLower: string
  pathLower: string
  pyFull: string      // 汉字全拼
  pyInit: string      // 汉字拼音首字母缩写
}

declare global {
  interface Window {
    electronAPI: ElectronAPI
  }
}

export {}
