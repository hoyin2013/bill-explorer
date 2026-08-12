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

// AI 接口配置
export interface AIConfig {
  baseURL: string
  apiKey: string
  model: string
  temperature: number
}

// AI 识别出的单条小票字段
export interface AIRecognizedRow {
  no?: string | number
  date?: string
  name?: string
  unit?: string
  qty?: string | number
  price?: string | number
  amount?: string | number
  person?: string
  remark?: string
  /** 仅前端使用：记录该结果来自哪张图片（多图扫描时便于核对），AI 不会返回此字段 */
  source?: string
}

// 应用设置
export interface AppSettings {
  aiConfig: AIConfig
  imageDir: string
  prompt: string
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
  // 读取整张表（全量，字符串矩阵）
  loadSheet: (filePath: string) => Promise<{
    error?: boolean
    message?: string
    sheetName?: string
    headerLabels: string[]
    rows: string[][]
  }>
  // 覆盖保存整张表
  saveSheet: (filePath: string, rows: string[][]) => Promise<{
    error?: boolean
    message?: string
    rowNumber?: number
    count?: number
  }>
  // 在文件管理器中打开文件所在文件夹并选中该文件
  revealFile: (filePath: string) => Promise<{ error?: boolean; message?: string }>
  // 列出某文件可用的备份版本（新 → 旧），time 形如 20260812-101530
  listBackups: (filePath: string) => Promise<{
    error?: boolean
    message?: string
    backups?: Array<{ path: string; time: string }>
  }>
  // 恢复某文件最近一份备份
  restoreBackup: (filePath: string) => Promise<{
    error?: boolean
    message?: string
    backupTime?: string
  }>
  // 获取设置
  getSettings: () => Promise<AppSettings>
  // 保存设置
  saveSettings: (settings: {
    aiConfig?: { baseURL?: string; apiKey?: string; model?: string; temperature?: number }
    imageDir?: string
    prompt?: string
  }) => Promise<{ error?: boolean; message?: string }>
  // 选择图片目录
  selectImageDirectory: () => Promise<string | null>
  // 列出图片目录中的图片
  listImages: (dir?: string) => Promise<{
    error?: boolean
    message?: string
    images?: Array<{ name: string; path: string }>
  }>
  // 读取图片为 base64
  readImageBase64: (imagePath: string) => Promise<{
    error?: boolean
    message?: string
    base64?: string
    mime?: string
  }>
  // AI 识别小票
  aiRecognize: (imagePath: string) => Promise<{
    error?: boolean
    message?: string
    rows?: AIRecognizedRow[]
  }>
  // 读取某 Excel 已自动填入过的图片列表（去重用）
  getApplied: (filePath: string) => Promise<string[]>
  // 记录某张图片已自动填入
  addApplied: (filePath: string, imagePath: string) => Promise<void>

  // 打开独立的小票识图窗口（不影响录入）
  openImageWindow: () => Promise<void>
  // 订阅主进程发往渲染进程的事件（返回取消订阅函数）
  on: (channel: string, listener: (...args: unknown[]) => void) => () => void
  // 识图窗口 → 主进程：上报本次识别出的人名（用于把左侧列表对应文件置顶）
  reportPersons: (persons: string[]) => void
  // 识图窗口 → 主进程：把识别结果回填到主窗口当前打开的录入网格（用户主动点击才触发）
  applyToMain: (rows: AIRecognizedRow[]) => void
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
