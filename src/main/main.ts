import { app, BrowserWindow, ipcMain, dialog } from 'electron'
import Store from 'electron-store'
import { join, basename } from 'path'
import { scanDirectory, openFile, revealFile } from './file-service'
import { appendRows, previewRows, updateRows, loadSheet, saveSheet, listBackups, restoreBackup } from './excel-memo'
import { listImages, readImageBase64 } from './image-service'
import { recognizeReceipt, DEFAULT_PROMPT, AIConfig, AIRecognizedRow, buildNameList, buildAugmentedPrompt, correctPersonNames, recognizeTicketsWithDetection, recognizeSingleCrop, type RecognizedTicket } from './ai-service'
import { getDefaultModelPath, detectTickets, detectEnvironment } from './detection'
import type { ImageSnapshot } from '../renderer/types'

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
  // 小票检测增强（YOLOv8）：本地检测模型逐张裁剪后再识别，提高识别率
  pythonPath?: string
  detectModel?: string
  enableDetect?: boolean
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
            fastMode: { type: 'boolean', default: true },
          },
        },
        imageDir: { type: 'string', default: '' },
        prompt: { type: 'string', default: DEFAULT_PROMPT },
        pythonPath: { type: 'string', default: '' },
        detectModel: { type: 'string', default: '' },
        enableDetect: { type: 'boolean', default: true },
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
  // 旧版默认提示词结构已变更（含"你是一位票据录入助手…"及上一版"识别不到则置 null"的 name 规则），
  // 自动升级为新的默认提示词，仅对"从未自定义过"的用户生效；自定义过的不动。
  let prompt = raw.prompt || DEFAULT_PROMPT
  if (prompt.includes('你是一位票据录入助手') || prompt.includes('识别不到则置 null')) prompt = DEFAULT_PROMPT
  return {
    aiConfig: {
      baseURL: raw.aiConfig?.baseURL || '',
      apiKey: raw.aiConfig?.apiKey || '',
      model: raw.aiConfig?.model || 'gpt-4o-mini',
      temperature: typeof raw.aiConfig?.temperature === 'number' ? raw.aiConfig.temperature : 0.2,
      fastMode: raw.aiConfig?.fastMode !== false,
    },
    imageDir: raw.imageDir || '',
    prompt,
    pythonPath: raw.pythonPath || '',
    detectModel: raw.detectModel || '',
    enableDetect: typeof raw.enableDetect === 'boolean' ? raw.enableDetect : true,
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
          fastMode: patch.aiConfig.fastMode !== undefined ? patch.aiConfig.fastMode : (current.aiConfig?.fastMode ?? true),
        }
      : current.aiConfig,
    imageDir: patch.imageDir !== undefined ? patch.imageDir : current.imageDir,
    prompt: patch.prompt !== undefined ? patch.prompt : current.prompt,
    pythonPath: patch.pythonPath !== undefined ? patch.pythonPath : current.pythonPath,
    detectModel: patch.detectModel !== undefined ? patch.detectModel : current.detectModel,
    enableDetect: patch.enableDetect !== undefined ? patch.enableDetect : current.enableDetect,
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

let mainWindow: BrowserWindow | null = null
// 「小票识图」独立窗口（从主窗口右侧面板拆出来，方便在大屏单独查看）
let detachedImageWin: BrowserWindow | null = null
// 拆分时由主窗口传入的状态快照，独立窗口启动后通过 get-detached-init 取回，以保留结果/进度
let pendingDetachedState: ImageSnapshot | null = null

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

// 创建「小票识图」独立窗口：只渲染 ImageWindow 组件（通过 ?detached=image 查询参数识别），
// 关闭时通知主窗口恢复右侧面板。
function createDetachedImageWindow() {
  if (detachedImageWin && !detachedImageWin.isDestroyed()) {
    detachedImageWin.focus()
    return
  }
  const win = new BrowserWindow({
    width: 960,
    height: 720,
    minWidth: 640,
    minHeight: 480,
    title: '小票识图（独立窗口）',
    backgroundColor: '#f5f6f8',
    webPreferences: {
      preload: PRELOAD_PATH,
      contextIsolation: true,
      nodeIntegration: false,
    },
  })
  if (process.env.VITE_DEV_SERVER_URL) {
    win.loadURL(process.env.VITE_DEV_SERVER_URL + '?detached=image')
  } else {
    win.loadFile(DIST_HTML, { search: 'detached=image' })
  }
  win.on('closed', () => {
    detachedImageWin = null
    // 关闭前若独立窗口已回传最新状态（点「合并」或 X 关闭前），先恢复到主窗口保留工作成果
    if (pendingDetachedState) {
      mainWindow?.webContents.send('image-detached-merged', pendingDetachedState)
    }
    // 通知主窗口：独立窗口已关闭，请恢复右侧「小票识图」面板
    mainWindow?.webContents.send('image-detached-closed')
    pendingDetachedState = null
  })
  detachedImageWin = win
}

// 创建主窗口后即可启动应用
app.whenReady().then(() => {
  createWindow()

  app.on('activate', () => {
    if (BrowserWindow.getAllWindows().length === 0) createWindow()
  })
})

app.on('window-all-closed', () => {
  // 关掉主窗口即整体退出（同时会先销毁 AI 识图窗口），不再区分平台保留在后台
  app.quit()
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
  // 遍历账单目录，取每个 Excel 的完整文件名作为清单（~ 开头的临时文件已由 scanDirectory 过滤）
  const nameList = await buildNameList(store.get('workDir') || '')
  const base = await readImageBase64(imagePath)
  if (base.error) return base
  // 把文件名清单附加到提示词之后，让 AI 在已知范围内识别并做模糊匹配
  const prompt = buildAugmentedPrompt(settings.prompt, nameList)
  const res = await recognizeReceipt(base.base64 || '', settings.aiConfig || { baseURL: '', apiKey: '', model: '', temperature: 0.2 }, prompt)
  // 识别完成后，把人名模糊匹配并修正为清单中的完整文件名写法
  if (res.rows && res.rows.length) {
    const { rows, corrected } = correctPersonNames(res.rows, nameList)
    res.rows = rows
    if (corrected > 0) {
      res.message = (res.message ? res.message + '\n' : '') + `已按人名清单自动修正 ${corrected} 处人名`
    }
  }
  return res
})

// ============ IPC 通道：仅检测小票边界框（用矩形框标出每张小票，不调用 AI 识别） ============
ipcMain.handle('ai-detect', async (_event, payload: { imagePath?: string; imageBase64?: string }) => {
  const settings = getSettings()
  const modelPath = settings.detectModel && settings.detectModel.trim() ? settings.detectModel.trim() : getDefaultModelPath()
  // 未显式指定时传空串，由 detection 自动探测可用解释器（python3 优先），
  // 不再写死 'python'（macOS 上默认就没有 python 命令）。
  const pythonPath = settings.pythonPath && settings.pythonPath.trim() ? settings.pythonPath.trim() : ''
  const enableDetect = settings.enableDetect !== false
  if (!enableDetect) {
    return { ok: false, modelAvailable: false, message: '检测增强已关闭，请在设置中启用。', boxes: [], imageWidth: 0, imageHeight: 0 }
  }
  try {
    const det = await detectTickets({ imagePath: payload.imagePath, imageBase64: payload.imageBase64, modelPath, pythonPath, crops: true })
    if (!det.modelAvailable) {
      return { ok: false, modelAvailable: false, message: det.message || '检测模型不可用', boxes: [], imageWidth: 0, imageHeight: 0, tickets: [] }
    }
    const tickets: RecognizedTicket[] = det.boxes.map((box, i) => ({
      index: i + 1,
      box,
      crop: det.crops[i] || '',
      angle: box.angle || 0,
      rows: [],
    }))
    return {
      ok: det.ok,
      modelAvailable: true,
      message: det.message,
      boxes: det.boxes,
      imageWidth: det.imageWidth,
      imageHeight: det.imageHeight,
      tickets,
    }
  } catch (err) {
    return { ok: false, modelAvailable: false, message: err instanceof Error ? err.message : '检测失败', boxes: [], imageWidth: 0, imageHeight: 0, tickets: [] }
  }
})

// ============ IPC 通道：探测检测增强的运行环境（Python / ultralytics / 模型是否就绪） ============
// 供识图窗口提前给出提示，而不是让用户点「检测」后才慢悠悠地失败。
ipcMain.handle('detect-environment', async () => {
  const env = await detectEnvironment()
  return env
})

// ============ IPC 通道：仅识别单张裁剪小票（用于「先框出、再逐张识别」流程） ============
ipcMain.handle('ai-recognize-crop', async (_event, cropBase64: string) => {
  const settings = getSettings()
  const nameList = await buildNameList(store.get('workDir') || '')
  try {
    const res = await recognizeSingleCrop(
      cropBase64,
      settings.aiConfig || { baseURL: '', apiKey: '', model: '', temperature: 0.2 },
      settings.prompt,
      nameList,
    )
    return res
  } catch (err) {
    return { error: true, message: err instanceof Error ? err.message : '单张识别失败' }
  }
})

// ============ IPC 通道：检测增强识别（YOLOv8 框出小票 → 逐张裁剪 → AI 识别人名与内容） ============
ipcMain.handle('ai-recognize-detected', async (_event, payload: { imagePath?: string; imageBase64?: string }) => {
  const settings = getSettings()
  const nameList = await buildNameList(store.get('workDir') || '')
  const modelPath = settings.detectModel && settings.detectModel.trim() ? settings.detectModel.trim() : getDefaultModelPath()
  const pythonPath = settings.pythonPath && settings.pythonPath.trim() ? settings.pythonPath.trim() : ''
  const enableDetect = settings.enableDetect !== false
  try {
    const res = await recognizeTicketsWithDetection(
      payload.imagePath || '',
      settings.aiConfig || { baseURL: '', apiKey: '', model: '', temperature: 0.2 },
      settings.prompt,
      nameList,
      { modelPath, pythonPath, enableDetect, imageBase64: payload.imageBase64 },
    )
    return res
  } catch (err) {
    return { error: true, message: err instanceof Error ? err.message : '检测增强识别失败', detected: false }
  }
})

// ============ IPC 通道：获取当前账单目录的人名清单（供设置界面展示） ============
ipcMain.handle('get-name-list', async () => {
  const list = await buildNameList(store.get('workDir') || '')
  return { names: list }
})

// 读取某 Excel 已自动填入过的图片列表（去重用）
ipcMain.handle('get-applied', (_event, filePath: string) => {
  return getApplied(filePath)
})

// 记录某张图片已自动填入
ipcMain.handle('add-applied', (_event, filePath: string, imagePath: string) => {
  addApplied(filePath, imagePath)
})

// ============ 识图面板 → 主窗口 ============
// 识图面板（已内嵌于主窗口右侧）上报识别出的人名，主窗口据此把对应 Excel 置顶
ipcMain.on('image:recognized-persons', (_event, persons: string[]) => {
  mainWindow?.webContents.send('recognized-persons', persons || [])
})

// 识图面板把识别结果回填到当前打开的录入网格（用户主动点击触发）
ipcMain.on('image:apply-rows', (_event, rows: unknown) => {
  mainWindow?.webContents.send('apply-recognized-rows', rows)
})

// ============ 独立「小票识图」窗口：拆分 / 合并 ============
// 把右侧面板拆成独立窗口（携带当前状态快照，独立窗口启动后取回以保留结果/进度）
ipcMain.handle('open-image-detached', (_event, state: unknown) => {
  pendingDetachedState = (state as ImageSnapshot) ?? null
  createDetachedImageWindow()
})

// 独立窗口启动时拉取拆分时传入的状态快照
ipcMain.handle('get-detached-init', () => {
  return pendingDetachedState
})

// 独立窗口关闭前回传最新状态（X 关闭 / 页面卸载时触发），保证主窗口合并回的是最新结果
ipcMain.on('detached-state-update', (_event, state: unknown) => {
  pendingDetachedState = (state as ImageSnapshot) ?? pendingDetachedState
})

// 把独立窗口合并回主窗口：记录最新状态后关闭该窗口，closed 事件会统一把最新状态恢复回主窗口
ipcMain.handle('attach-image-detached', (_event, state: unknown) => {
  pendingDetachedState = (state as ImageSnapshot) ?? pendingDetachedState
  if (detachedImageWin && !detachedImageWin.isDestroyed()) {
    detachedImageWin.close()
  }
  detachedImageWin = null
})
