import { app, BrowserWindow, ipcMain, dialog } from 'electron'
import Store from 'electron-store'
import { join, basename } from 'path'
import { scanDirectory, openFile } from './file-service'
import { appendRows, previewRows, updateRows, loadSheet, saveSheet } from './excel-memo'

// 最近修改历史的单条记录
interface HistoryRecord {
  filePath: string
  fileName: string
  time: string
}

// 历史记录最多保留条数
const HISTORY_LIMIT = 50

// 持久化配置存储：工作目录 + 最近修改历史
const store = new Store<{ workDir?: string; history?: HistoryRecord[] }>({
  schema: {
    workDir: { type: 'string', default: '' },
    history: { type: 'array', default: [] },
  },
})

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

function createWindow() {
  mainWindow = new BrowserWindow({
    width: 1220,
    height: 780,
    minWidth: 980,
    minHeight: 560,
    title: '账单查找器',
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

// ============ IPC 通道：获取最近修改历史 ============
ipcMain.handle('get-history', () => {
  return store.get('history') || []
})

// ============ IPC 通道：清空最近修改历史 ============
ipcMain.handle('clear-history', () => {
  store.set('history', [])
})
