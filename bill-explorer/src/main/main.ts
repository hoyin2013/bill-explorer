import { app, BrowserWindow, ipcMain, dialog } from 'electron'
import Store from 'electron-store'
import { join } from 'path'
import { scanDirectory, openFile } from './file-service'
import { appendRows, previewRows } from './excel-memo'

// 持久化配置存储，只保存工作目录路径
const store = new Store<{ workDir?: string }>({
  schema: {
    workDir: { type: 'string', default: '' },
  },
})

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
  async (_event, payload: { filePath: string; rows: Array<{ date: string; name: string; unit: string; qty: number; price: number; amount: number; person: string; remark: string }> }) => {
    const { filePath, rows } = payload
    if (!filePath) {
      return { error: true, message: '文件路径不能为空。' }
    }
    // 过滤空行 + 兜底补算金额
    const valid = (rows || [])
      .filter((r) => r.name || r.remark || r.unit || r.person || r.date || r.qty !== 0 || r.price !== 0)
      .map((r) => {
        const qty = Number(r.qty) || 0
        const price = Number(r.price) || 0
        return {
          ...r,
          qty,
          price,
          amount: !isNaN(qty) && !isNaN(price) ? +(qty * price) : (Number(r.amount) || 0),
        }
      })
    if (valid.length === 0) {
      return { error: true, message: '请至少填写一行有效数据。' }
    }
    try {
      return await appendRows(filePath, valid)
    } catch (err) {
      return {
        error: true,
        message: err instanceof Error ? err.message : '写入失败',
      }
    }
  },
)
