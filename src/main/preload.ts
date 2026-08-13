import { contextBridge, ipcRenderer } from 'electron'
import type { AIRecognizedRow } from '../renderer/types'

// 暴露给渲染进程的安全 IPC API（contextIsolation 下必须通过此桥接）
contextBridge.exposeInMainWorld('electronAPI', {
  // 选择文件夹，返回绝对路径或 null
  selectDirectory: () => ipcRenderer.invoke('select-directory'),
  // 获取上次持久化的工作目录
  getWorkDir: () => ipcRenderer.invoke('get-work-dir'),
  // 递归扫描目录下的 xlsx/xls 文件
  scanFiles: (dir: string) => ipcRenderer.invoke('scan-files', dir),
  // 调用系统默认程序打开文件
  openFile: (filePath: string) => ipcRenderer.invoke('open-file', filePath),
  // 用系统默认管理器打开文件夹
  openDir: (dirPath: string) => ipcRenderer.invoke('open-dir', dirPath),
  // 读取 Excel 文件最近 N 行预览
  previewRows: (filePath: string, limit: number) => ipcRenderer.invoke('preview-rows', filePath, limit),
  // 向 Excel 文件批量追加多条账单记录
  appendMemo: (filePath: string, rows: Array<{
    no: string; date: string; name: string; unit: string; qty: number; price: number; amount: number | ''; person: string; remark: string
  }>) => ipcRenderer.invoke('append-memo', { filePath, rows }),
  // 覆盖更新同一文件连续编辑时上次写入的行
  updateMemo: (filePath: string, rows: Array<{
    no: string; date: string; name: string; unit: string; qty: number; price: number; amount: number | ''; person: string; remark: string
  }>) => ipcRenderer.invoke('update-memo', { filePath, rows }),
  // 读取整张表（全量）
  loadSheet: (filePath: string) => ipcRenderer.invoke('load-sheet', filePath),
  // 覆盖保存整张表
  saveSheet: (filePath: string, rows: string[][]) => ipcRenderer.invoke('save-sheet', { filePath, rows }),
  // 在文件管理器中打开文件所在文件夹并选中
  revealFile: (filePath: string) => ipcRenderer.invoke('reveal-file', filePath),
  // 列出某文件可用的备份版本（新 → 旧）
  listBackups: (filePath: string) => ipcRenderer.invoke('list-backups', filePath),
  // 恢复某文件最近一份备份
  restoreBackup: (filePath: string) => ipcRenderer.invoke('restore-backup', filePath),
  // 获取最近修改历史
  getHistory: () => ipcRenderer.invoke('get-history'),
  // 清空最近修改历史
  clearHistory: () => ipcRenderer.invoke('clear-history'),
  // 获取设置
  getSettings: () => ipcRenderer.invoke('get-settings'),
  // 保存设置
  saveSettings: (settings: {
    aiConfig?: { baseURL?: string; apiKey?: string; model?: string; temperature?: number }
    imageDir?: string
    prompt?: string
    pythonPath?: string
    detectModel?: string
    enableDetect?: boolean
  }) => ipcRenderer.invoke('save-settings', settings),
  // 选择图片目录
  selectImageDirectory: () => ipcRenderer.invoke('select-image-directory'),
  // 列出图片目录中的图片
  listImages: (dir?: string) => ipcRenderer.invoke('list-images', dir),
  // 读取图片为 base64
  readImageBase64: (imagePath: string) => ipcRenderer.invoke('read-image-base64', imagePath),
  // 获取当前账单目录的人名清单（文件名提取）
  getNameList: () => ipcRenderer.invoke('get-name-list'),
  // AI 识别小票
  aiRecognize: (imagePath: string) => ipcRenderer.invoke('ai-recognize', imagePath),
  // 仅检测小票边界框
  aiDetect: (payload) => ipcRenderer.invoke('ai-detect', payload),
  // 检测增强识别（检测 + 逐张裁剪 + AI）
  aiRecognizeDetected: (payload) => ipcRenderer.invoke('ai-recognize-detected', payload),
  // 仅识别单张裁剪小票
  aiRecognizeCrop: (cropBase64: string) => ipcRenderer.invoke('ai-recognize-crop', cropBase64),
  // 读取某 Excel 已自动填入过的图片列表（去重用）
  getApplied: (filePath: string) => ipcRenderer.invoke('get-applied', filePath),
  // 记录某张图片已自动填入
  addApplied: (filePath: string, imagePath: string) => ipcRenderer.invoke('add-applied', filePath, imagePath),
  // 打开独立的小票识图窗口
  openImageWindow: () => ipcRenderer.invoke('open-image-window'),
  // 探测检测增强的运行环境（Python / ultralytics / 模型是否就绪）
  detectEnvironment: () => ipcRenderer.invoke('detect-environment'),
  // 订阅主进程事件（返回取消订阅函数）；用于接收 recognized-persons / apply-recognized-rows
  on: (channel: string, listener: (...args: unknown[]) => void) => {
    const sub = (_e: unknown, ...args: unknown[]) => listener(...args)
    ipcRenderer.on(channel, sub)
    return () => ipcRenderer.removeListener(channel, sub)
  },
  // 上报识别出的人名（主进程转发给主窗口，用于左侧列表置顶）
  reportPersons: (persons: string[]) => ipcRenderer.send('image:recognized-persons', persons),
  // 把识别结果回填到主窗口当前录入网格
  applyToMain: (rows: AIRecognizedRow[]) => ipcRenderer.send('image:apply-rows', rows),
  // 识图窗口即将关闭时通知主进程（清空主窗口搜索框 + 清除命中置顶）
  notifyImageClosing: () => ipcRenderer.send('image:closing'),
})
