import { app, BrowserWindow, ipcMain, dialog } from 'electron'
import Store from 'electron-store'
import { join, basename } from 'path'
import { scanDirectory, openFile, revealFile } from './file-service'
import { appendRows, previewRows, updateRows, loadSheet, saveSheet, listBackups, restoreBackup } from './excel-memo'
import { listImages, readImageBase64 } from './image-service'
import { recognizeReceipt, DEFAULT_PROMPT, AIConfig, AIRecognizedRow } from './ai-service'

// 最近修改历史的单条记录
interface HistoryRecord {
  filePath: string
  fileName: string
  time: string
}

interface AppSettings {
  aiConfig?: AIConfig
  imageDir?: string
  prompt?: string
}

// 历史记录最多保留条数
const HISTORY_LIMIT = 50

// 持久化配置存储：工作目录 + 最近修改历史 + AI/图片设置
const store = new Store<{
  workDir?: string
  history?: HistoryRecord[]
  settings?: AppSettings
  applied?: Record<string, string[]>
}>({
  schema: {
    workDir: { type: 'string', default: '' },
    history: { type: 'array', default: [] },
    settings: {
      type: 'object',
      default: {},
      properties: {
        aiConfig: {
          type: 'object',
          default: {},
          properties: {
            baseURL: { type: 'string', default: '' },
            apiKey: { type: 'string', default: '' },
            model: { type: 'string', default: 'gpt-4o-mini' },
            temperature: { type: 'number', default: 0.2 },
          },
        },
        imageDir: { type: 'string', default: '' },
        prompt: { type: 'string', default: DEFAULT_PROMPT },
      },
    },
    // 已自动填入过的图片：filePath -> 图片路径数组，避免重复打开反复添加同一张
    applied: {
      type: 'object',
      default: {},
    },
  },
})

// 读取某个 Excel 已自动填入过的图片路径列表
function getApplied(filePath: string): string[] {
  const all = (store.get('applied') as Record<string, string[]>) || {}
  return Array.isArray(all[filePath]) ? all[filePath] : []
}

// 记录某张图片已自动填入（去重）
function addApplied(filePath: string, imagePath: string) {
  const all = (store.get('applied') as Record<string, string[]>) || {}
  const list = Array.isArray(all[filePath]) ? all[filePath].slice() : []
  if (!list.includes(imagePath)) list.push(imagePath)
  all[filePath] = list
  store.set('applied', all)
}

function getSettings(): AppSettings {
  const raw = store.get('settings') || {}
  return {
    aiConfig: {
      baseURL: raw.aiConfig?.baseURL || '',
      apiKey: raw.aiConfig?.apiKey || '',
      model: raw.aiConfig?.model || 'gpt-4o-mini',
      temperature: typeof raw.aiConfig?.temperature === 'number' ? raw.aiConfig.temperature : 0.2,
    },
    imageDir: raw.imageDir || '',
    prompt: raw.prompt || DEFAULT_PROMPT,
  }
}

function saveSettings(patch: AppSettings): void {
  const current = getSettings()
  const next: AppSettings = {
    aiConfig: patch.aiConfig
      ? {
          baseURL: patch.aiConfig.baseURL?.trim() || current.aiConfig?.baseURL || '',
          apiKey: patch.aiConfig.apiKey || current.aiConfig?.apiKey || '',
          model: patch.aiConfig.model?.trim() || current.aiConfig?.model || 'gpt-4o-mini',
          temperature: typeof patch.aiConfig.temperature === 'number'
            ? patch.aiConfig.temperature
            : current.aiConfig?.temperature ?? 0.2,
        }
      : current.aiConfig,
    imageDir: patch.imageDir !== undefined ? patch.imageDir : current.imageDir,
    prompt: patch.prompt !== undefined ? patch.prompt : current.prompt,
  }
  store.set('settings', next)
}

// 记录一次成功保存：最新在前、按文件去重、最多保留 HISTORY_LIMIT 条
function recordHistory(filePath: string) {
  const time = new Date().toLocaleString('zh-CN', { hour12: false })
  const entry: HistoryRecord = { filePath, fileName: basename(filePath), time }
  const history = store.get('history') || []
  const next = [entry, ...history.filter((h) => h.filePath !== filePath)]
  store.set('history', next.slice(0, HISTORY_LIMIT))
}

// 记录每个文件最近一次写入的行范围（给"同一文件连续编辑 → 覆盖更新"用）
const lastWriteMap = new Map<string, { startRow: number; count: number }>()

// 行类型：amount 允许空字符串（表示金额留空）
type MemoRow = {
  no: string; date: string; name: string; unit: string
  qty: number; price: number; amount: number | ''; person: string; remark: string
}

// 过滤空行 + 规范化：数量/单价转数字，金额留空则保持 ''
function sanitizeRows(rows?: MemoRow[]): MemoRow[] {
  return (rows || [])
    .filter(
      (r) => r.date || r.name || r.unit || r.person || r.remark ||
        r.qty !== 0 || r.price !== 0 ||
        (r.amount !== '' && r.amount != null && r.amount !== 0),
    )
    .map((r) => {
      const qty = Number(r.qty) || 0
      const price = Number(r.price) || 0
      return {
        ...r,
        no: r.no || '',
        qty,
        price,
        amount: (r.amount === '' || r.amount == null) ? '' : Number(r.amount),
      }
    })
}

// 处理 append / update 的结果：成功时记录历史 + 记录写入位置
function afterWrite(filePath: string, result: { rowNumber?: number; count?: number; error?: boolean }) {
  if (result.error || result.rowNumber == null) return
  lastWriteMap.set(filePath, { startRow: result.rowNumber, count: result.count || 0 })
  recordHistory(filePath)
}

// 文件路径（生产模式通过 file:// 加载）
const PRELOAD_PATH = join(__dirname, 'preload.js')
const DIST_HTML = join(__dirname, '../dist/index.html')
const DIST_IMAGE_HTML = join(__dirname, '../dist/image.html')

let mainWindow: BrowserWindow | null = null
let imageWindow: BrowserWindow | null = null

function createWindow() {
  mainWindow = new BrowserWindow({
    width: 1220,
    height: 780,
    minWidth: 980,
    minHeight: 560,
    title: '账单录入器',
    webPreferences: {
      preload: PRELOAD_PATH,
      contextIsolation: true,
      nodeIntegration: false,
    },
  })

  // 开发模式走 Vite dev server，生产模式走本地文件
  if (process.env.VITE_DEV_SERVER_URL) {
    mainWindow.loadURL(process.env.VITE_DEV_SERVER_URL)
  } else {
    mainWindow.loadFile(DIST_HTML)
  }
}

// 创建独立的小票识图窗口（不影响主窗口录入；可同时打开、自由摆放）
function createImageWindow() {
  if (imageWindow && !imageWindow.isDestroyed()) {
    imageWindow.focus()
    return
  }
  imageWindow = new BrowserWindow({
    width: 560,
    height: 840,
    minWidth: 420,
    minHeight: 560,
    title: '账单录入器 - 小票识图',
    show: true,
    webPreferences: {
      preload: PRELOAD_PATH,
      contextIsolation: true,
      nodeIntegration: false,
    },
  })

  if (process.env.VITE_DEV_SERVER_URL) {
    imageWindow.loadURL(process.env.VITE_DEV_SERVER_URL + '/image.html')
  } else {
    imageWindow.loadFile(DIST_IMAGE_HTML)
  }

  // 关闭识图窗口时，清除主窗口左侧列表的人名置顶
  imageWindow.on('closed', () => {
    imageWindow = null
    mainWindow?.webContents.send('recognized-persons', [])
  })
}

app.whenReady().then(() => {
  createWindow()

  app.on('activate', () => {
    if (BrowserWindow.getAllWindows().length === 0) createWindow()
  })
})

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') app.quit()
})

// ============ IPC 通道：选择工作目录 ============
ipcMain.handle('select-directory', async () => {
  const result = await dialog.showOpenDialog(mainWindow!, {
    properties: ['openDirectory'],
  })
  if (result.canceled || result.filePaths.length === 0) {
    return null
  }
  const dir = result.filePaths[0]
  // 持久保存选中的工作目录
  store.set('workDir', dir)
  return dir
})

// ============ IPC 通道：获取持久化的工作目录 ============
ipcMain.handle('get-work-dir', () => {
  return store.get('workDir')
})

// ============ IPC 通道：递归扫描目录下的 xlsx/xls 文件 ============
ipcMain.handle('scan-files', async (_event, dir: string) => {
  try {
    return await scanDirectory(dir)
  } catch (err) {
    const message = err instanceof Error ? err.message : '未知错误'
    return { error: true, message }
  }
})

// ============ IPC 通道：调用系统默认程序打开文件 ============
ipcMain.handle('open-file', async (_event, filePath: string) => {
  return openFile(filePath)
})

// ============ IPC 通道：在文件管理器中打开文件所在文件夹并选中 ============
ipcMain.handle('reveal-file', (_event, filePath: string) => {
  return revealFile(filePath)
})

// ============ IPC 通道：用系统默认管理器打开文件夹 ============
ipcMain.handle('open-dir', (_event, dirPath: string) => {
  return openFile(dirPath)
})

// ============ IPC 通道：读取 Excel 文件最近 N 行（预览，不修改） ============
ipcMain.handle('preview-rows', async (_event, filePath: string, limit?: number) => {
  try {
    return await previewRows(filePath, limit || 10)
  } catch (err) {
    const message = err instanceof Error ? err.message : '预览失败'
    return { error: true, message }
  }
})

// ============ IPC 通道：向 Excel 文件批量追加多条账单记录 ============
ipcMain.handle(
  'append-memo',
  async (_event, payload: { filePath: string; rows: MemoRow[] }) => {
    const { filePath, rows } = payload
    if (!filePath) {
      return { error: true, message: '文件路径不能为空。' }
    }
    const valid = sanitizeRows(rows)
    if (valid.length === 0) {
      return { error: true, message: '请至少填写一行有效数据。' }
    }
    try {
      const result = await appendRows(filePath, valid)
      afterWrite(filePath, result)
      return result
    } catch (err) {
      return {
        error: true,
        message: err instanceof Error ? err.message : '写入失败',
      }
    }
  },
)

// ============ IPC 通道：覆盖更新同一文件连续编辑时上次写入的行 ============
ipcMain.handle(
  'update-memo',
  async (_event, payload: { filePath: string; rows: MemoRow[] }) => {
    const { filePath, rows } = payload
    if (!filePath) {
      return { error: true, message: '文件路径不能为空。' }
    }
    const prev = lastWriteMap.get(filePath)
    if (!prev) {
      return { error: true, message: '该文件还没有可更新的记录。' }
    }
    const valid = sanitizeRows(rows)
    if (valid.length === 0) {
      return { error: true, message: '请至少填写一行有效数据。' }
    }
    try {
      const result = await updateRows(filePath, valid, prev.startRow, prev.count)
      afterWrite(filePath, result)
      return result
    } catch (err) {
      return {
        error: true,
        message: err instanceof Error ? err.message : '更新失败',
      }
    }
  },
)

// ============ IPC 通道：读取整张表（类 Excel 网格编辑：全量加载） ============
ipcMain.handle('load-sheet', async (_event, filePath: string) => {
  try {
    return await loadSheet(filePath)
  } catch (err) {
    const message = err instanceof Error ? err.message : '读取失败'
    return { error: true, message }
  }
})

// ============ IPC 通道：覆盖保存整张表（保存 = 更新，而非追加） ============
ipcMain.handle(
  'save-sheet',
  async (_event, payload: { filePath: string; rows: string[][] }) => {
    const { filePath, rows } = payload
    if (!filePath) {
      return { error: true, message: '文件路径不能为空。' }
    }
    try {
      const result = await saveSheet(filePath, rows)
      if (!result.error) afterWrite(filePath, result)
      return result
    } catch (err) {
      return {
        error: true,
        message: err instanceof Error ? err.message : '保存失败',
      }
    }
  },
)

// ============ IPC 通道：列出某文件可用的备份版本（新 → 旧） ============
ipcMain.handle('list-backups', (_event, filePath: string) => {
  try {
    return { backups: listBackups(filePath) }
  } catch (err) {
    return { error: true, message: err instanceof Error ? err.message : '读取备份失败' }
  }
})

// ============ IPC 通道：恢复某文件最近一份备份 ============
ipcMain.handle('restore-backup', async (_event, filePath: string) => {
  try {
    return await restoreBackup(filePath)
  } catch (err) {
    return { error: true, message: err instanceof Error ? err.message : '恢复失败' }
  }
})

// ============ IPC 通道：获取最近修改历史 ============
ipcMain.handle('get-history', () => {
  return store.get('history') || []
})

// ============ IPC 通道：清空最近修改历史 ============
ipcMain.handle('clear-history', () => {
  store.set('history', [])
})

// ============ IPC 通道：获取设置（AI + 图片目录 + 提示词） ============
ipcMain.handle('get-settings', () => {
  return getSettings()
})

// ============ IPC 通道：保存设置 ============
ipcMain.handle('save-settings', (_event, patch: AppSettings) => {
  try {
    saveSettings(patch)
    return { error: false }
  } catch (err) {
    return { error: true, message: err instanceof Error ? err.message : '保存设置失败' }
  }
})

// ============ IPC 通道：选择图片目录 ============
ipcMain.handle('select-image-directory', async () => {
  const result = await dialog.showOpenDialog(mainWindow!, {
    properties: ['openDirectory'],
    title: '选择小票图片保存目录',
  })
  if (result.canceled || result.filePaths.length === 0) {
    return null
  }
  const dir = result.filePaths[0]
  saveSettings({ imageDir: dir })
  return dir
})

// ============ IPC 通道：列出图片目录中的图片 ============
ipcMain.handle('list-images', (_event, dir?: string) => {
  const target = dir || getSettings().imageDir || ''
  return listImages(target)
})

// ============ IPC 通道：读取图片为 base64（并压缩，用于预览和 AI） ============
ipcMain.handle('read-image-base64', (_event, imagePath: string) => {
  return readImageBase64(imagePath)
})

// ============ IPC 通道：AI 识别小票 ============
ipcMain.handle('ai-recognize', async (_event, imagePath: string) => {
  const settings = getSettings()
  const base = readImageBase64(imagePath)
  if (base.error) return base
  const res = await recognizeReceipt(base.base64 || '', settings.aiConfig || { baseURL: '', apiKey: '', model: '', temperature: 0.2 }, settings.prompt)
  return res
})

// 读取某 Excel 已自动填入过的图片列表（去重用）
ipcMain.handle('get-applied', (_event, filePath: string) => {
  return getApplied(filePath)
})

// 记录某张图片已自动填入
ipcMain.handle('add-applied', (_event, filePath: string, imagePath: string) => {
  addApplied(filePath, imagePath)
})

// ============ 独立"小票识图"窗口 ============
// 打开独立识图窗口（不影响主窗口录入）
ipcMain.handle('open-image-window', () => {
  createImageWindow()
})

// 识图窗口 → 主窗口：转发识别出的人名（主窗口据此把对应 Excel 置顶）
ipcMain.on('image:recognized-persons', (_event, persons: string[]) => {
  mainWindow?.webContents.send('recognized-persons', persons || [])
})

// 识图窗口 → 主窗口：把识别结果回填到当前打开的录入网格（用户主动点击触发）
ipcMain.on('image:apply-rows', (_event, rows: unknown) => {
  mainWindow?.webContents.send('apply-recognized-rows', rows)
})
