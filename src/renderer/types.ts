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
  /** 快速模式：关闭模型思考/推理(reasoning)，加快响应；默认开启 */
  fastMode?: boolean
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
  /** 前端使用：人名是否已被「人名清单」自动修正过 */
  personCorrected?: boolean
}

// 应用设置
export interface AppSettings {
  aiConfig: AIConfig
  imageDir: string
  prompt: string
  // 小票检测增强（YOLOv8）相关
  pythonPath?: string
  detectModel?: string
  enableDetect?: boolean
}

// YOLOv8 检测出的小票边界框（基于预览缩放图像素坐标）
export interface DetectedBox {
  x: number
  y: number
  w: number
  h: number
  conf: number
  cls: number
  label?: string
  /** 该小票裁剪图被自动旋转的角度（度） */
  angle?: number
}

// 单张拆分出来的小票（用于逐个放大查看与录入）
export interface RecognizedTicket {
  index: number
  box: DetectedBox
  crop: string
  angle: number
  rows: AIRecognizedRow[]
}

// 「小票识图」窗口的可序列化状态快照：用于拆分/合并独立窗口时保留已识别结果与进度
export interface ImageSnapshot {
  imageDir: string
  images: Array<{ name: string; path: string }>
  selected: string
  preview: string
  rotate: number
  rows: AIRecognizedRow[]
  boxes: DetectedBox[]
  imgNatural: { w: number; h: number }
  tickets: RecognizedTicket[]
  viewMode: 'overview' | 'single'
  active: number
  singleRotate: number
  recogState: Record<number, 'pending' | 'busy' | 'done' | 'error'>
  zoom: number
  pan: { x: number; y: number }
  singleZoom: number
  singlePan: { x: number; y: number }
  listWidth: number
  listCollapsed: boolean
  dateAnchor: string | null
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
    pythonPath?: string
    detectModel?: string
    enableDetect?: boolean
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
  // 读取全局默认图片旋转角（度，0/90/180/270）：旋转任一图片后自动保存，其后每张图套用此方向
  getImageRotation: () => Promise<number>
  // 保存全局默认图片旋转角（自动保存用户设定的方向）
  setImageRotation: (angle: number) => Promise<void>
  // 获取当前账单目录的人名清单（从文件名提取）
  getNameList: () => Promise<{ names: string[] }>
  // AI 识别小票
  aiRecognize: (imagePath: string) => Promise<{
    error?: boolean
    message?: string
    rows?: AIRecognizedRow[]
  }>
  // 仅检测小票边界框（返回矩形框 + 逐张裁剪图，用于在前端预览上画出并点击放大）。
  // 可传 imageBase64（前端已旋转/缩放后的图），否则用 imagePath 读原图。
  aiDetect: (payload: { imagePath?: string; imageBase64?: string }) => Promise<{
    ok?: boolean
    modelAvailable?: boolean
    message?: string
    boxes?: DetectedBox[]
    imageWidth?: number
    imageHeight?: number
    tickets?: RecognizedTicket[]
  }>
  // 检测增强识别：YOLOv8 框出小票 → 逐张裁剪 → AI 识别（返回识别结果 + 框坐标 + 逐张小票）。
  // 可传 imageBase64（前端已旋转/缩放后的图），否则用 imagePath 读原图。
  aiRecognizeDetected: (payload: { imagePath?: string; imageBase64?: string }) => Promise<{
    error?: boolean
    message?: string
    rows?: AIRecognizedRow[]
    boxes?: DetectedBox[]
    imageWidth?: number
    imageHeight?: number
    detected?: boolean
    modelAvailable?: boolean
    tickets?: RecognizedTicket[]
  }>
  // 仅识别单张裁剪小票（crop 为 base64 jpeg），用于「先框出、再逐张识别」流程
  aiRecognizeCrop: (cropBase64: string) => Promise<{
    error?: boolean
    message?: string
    rows?: AIRecognizedRow[]
  }>
  // 读取某 Excel 已自动填入过的图片列表（去重用）
  getApplied: (filePath: string) => Promise<string[]>
  // 记录某张图片已自动填入
  addApplied: (filePath: string, imagePath: string) => Promise<void>

  // 探测检测增强环境（ONNX 模型与运行时是否就绪），供 UI 提前提示
  detectEnvironment: () => Promise<{
    modelExists: boolean
    runtimeReady: boolean
    detail: string
  }>
  // 订阅主进程发往渲染进程的事件（返回取消订阅函数）
  on: (channel: string, listener: (...args: unknown[]) => void) => () => void
  // 识图窗口 → 主进程：上报本次识别出的人名（用于把左侧列表对应文件置顶）
  reportPersons: (persons: string[]) => void
  // 识图窗口 → 主进程：把识别结果回填到主窗口当前打开的录入网格（用户主动点击才触发）
  applyToMain: (rows: AIRecognizedRow[]) => void
  // 把「小票识图」面板拆成独立窗口（携带当前状态快照，主进程创建 BrowserWindow 后回传）
  openImageDetached: (state: ImageSnapshot) => Promise<void>
  // 把独立识图窗口合并回主窗口（携带最新状态快照，主进程关闭该窗口并把状态回传给主窗口）
  attachImageDetached: (state: ImageSnapshot) => Promise<void>
  // 独立窗口启动时拉取拆分时传入的状态快照（保留结果与进度）
  getDetachedInit: () => Promise<ImageSnapshot | null>
  // 独立窗口关闭前回传最新状态（X 关闭 / 页面卸载时触发），保证主窗口合并回的是最新结果
  detachedStateUpdate: (state: ImageSnapshot) => void
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
