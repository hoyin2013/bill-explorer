import { contextBridge, ipcRenderer } from 'electron'

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
  // 获取最近修改历史
  getHistory: () => ipcRenderer.invoke('get-history'),
  // 清空最近修改历史
  clearHistory: () => ipcRenderer.invoke('clear-history'),
})
