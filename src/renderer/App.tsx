import { useEffect, useMemo, useRef, useState } from 'react'
import { ElectronAPI, FileEntry, HistoryRecord } from './types'
import { filterFiles } from './utils'
import { DirectoryBar } from './components/DirectoryBar'
import { ResultList } from './components/ResultList'
import { MemoPanel } from './components/MemoPanel'
import { ErrorMessage } from './components/ErrorMessage'
import { HistoryList } from './components/HistoryList'

// 安全检查：electronAPI 来自 preload，若为 undefined 说明 Electron 未正确加载 preload
const api = window.electronAPI

// 当 api 缺失时展示明确错误，避免用户看到一片空白
function BrokenState() {
  return (
    <div className="app">
      <h1 className="app-title">账单文件查找器</h1>
      <div className="error-msg" style={{ margin: '20px 0' }}>
        <span className="error-icon">!</span>
        <span>
          无法连接主进程（electronAPI 未加载）。
          通常是因为从错误的目录启动了 Electron，或缺少 dist-electron/preload.js。
        </span>
      </div>
      <div className="error-msg" style={{ margin: '8px 0', background: '#fef0f0' }}>
        <span className="error-icon">→</span>
        <span>
          请先在项目目录 bill-explorer 下执行 <code style={{ fontWeight: 600 }}>npm run build:preload</code>，
          然后<strong>在 bill-explorer 目录内</strong>运行 <code style={{ fontWeight: 600 }}>npx electron .</code>
        </span>
      </div>
    </div>
  )
}

export default function App() {
  if (!api) return <BrokenState />

  return <AppInner api={api} />
}

function AppInner({ api }: { api: ElectronAPI }) {
  const [workDir, setWorkDir] = useState('')
  const [files, setFiles] = useState<FileEntry[]>([])
  const [query, setQuery] = useState('')
  const [loading, setLoading] = useState(false)
  const [status, setStatus] = useState('')
  const [error, setError] = useState('')

  // ---- 录入面板状态 ----
  const [activeFile, setActiveFile] = useState<FileEntry | null>(null)
  const [memoSaving, setMemoSaving] = useState(false)
  const [memoStatus, setMemoStatus] = useState('')

  // ---- 最近修改历史 ----
  const [history, setHistory] = useState<HistoryRecord[]>([])

  const searchRef = useRef<HTMLInputElement>(null)
  // 记录最近一次保存的文件路径（用于"同一文件连续编辑 → 覆盖更新"）
  const lastSavedFile = useRef<string | null>(null)

  // 每次窗口激活（从后台 / Dock 切回来）都把光标定位到搜索框，方便直接打字
  useEffect(() => {
    const onFocus = () => {
      // 仅在没有正在录入面板时聚焦搜索框，避免打断用户录入
      if (!activeFile) {
        searchRef.current?.focus()
      }
    }
    window.addEventListener('focus', onFocus)
    return () => window.removeEventListener('focus', onFocus)
  }, [activeFile])

  // 软件启动时自动加载上次持久化的工作目录
  useEffect(() => {
    loadWorkDir()
    refreshHistory()
  }, [])

  // 加载最近修改历史
  async function refreshHistory() {
    try {
      setHistory((await api.getHistory()) || [])
    } catch {
      setHistory([])
    }
  }

  async function loadWorkDir() {
    try {
      const dir = await api.getWorkDir()
      if (dir) {
        setWorkDir(dir)
        scan(dir)
      }
    } catch (err) {
      setError(err instanceof Error ? err.message : '加载配置失败')
    }
  }

  // 扫描：选择新目录时触发；刷新时也触发
  async function scan(dir: string) {
    setLoading(true)
    setError('')
    setFiles([])
    setStatus('扫描中，请稍候...')

    let result: { error?: boolean; message?: string; files?: FileEntry[] }
    try {
      result = await api.scanFiles(dir)
    } catch (err) {
      result = { error: true, message: err instanceof Error ? err.message : '扫描失败' }
    }

    if (result.error) {
      setError(result.message || '扫描失败')
      setFiles([])
    } else {
      setFiles(result.files || [])
      setStatus(`共 ${result.files?.length ?? 0} 个文件`)
    }

    setLoading(false)
    setStatus('')
  }

  // 选择工作目录
  async function onChooseDir() {
    const dir = await api.selectDirectory()
    if (dir) {
      setWorkDir(dir)
      scan(dir)
    }
  }

  // 刷新：重新扫描当前工作目录
  async function onRefresh() {
    if (workDir) scan(workDir)
  }

  // 双击路径：用系统默认管理器打开工作目录文件夹
  async function onOpenDir() {
    if (!workDir) return
    try {
      const result = await api.openDir(workDir)
      if (result.error) {
        setMemoStatus('打开目录失败：' + (result.message || '未知错误'))
        setTimeout(() => setMemoStatus(''), 3000)
      }
    } catch (err) {
      setMemoStatus('打开目录失败')
      setTimeout(() => setMemoStatus(''), 3000)
    }
  }

  // ---- 单击条目：在应用内打开录入面板 ----
  function onOpen(index: number) {
    const file = filtered[index]
    if (!file) return
    setActiveFile(file)
    setMemoStatus('')
  }

  // ---- 双击条目：用系统默认程序（Excel）打开 ----
  async function onOpenInExcel(index: number) {
    const file = filtered[index]
    if (!file) return
    try {
      await api.openFile(file.filePath)
    } catch (err) {
      const msg = err instanceof Error ? err.message : '打开失败'
      setMemoStatus('打开失败：' + msg)
      setTimeout(() => setMemoStatus(''), 3000)
    }
  }

  // ---- 关闭录入面板 ----（关闭后视为一次独立编辑会话，下次保存回到"追加"）
  function onCloseMemo() {
    setActiveFile(null)
    setMemoStatus('')
    lastSavedFile.current = null
  }

  // ---- 点击历史记录：若在当前扫描结果中则打开录入，否则用系统程序打开 ----
  async function onHistorySelect(filePath: string) {
    const file = files.find((f) => f.filePath === filePath)
    if (file) {
      setActiveFile(file)
      setMemoStatus('')
    } else {
      try {
        await api.openFile(filePath)
      } catch {
        setMemoStatus('打开失败：' + filePath)
        setTimeout(() => setMemoStatus(''), 3000)
      }
    }
  }

  // ---- 清空历史记录 ----
  async function onClearHistory() {
    try {
      await api.clearHistory()
      setHistory([])
    } catch {
      /* ignore */
    }
  }

  // ---- 保存：把面板内的多行数据一次性写回 Excel ----
  // 若与上次保存的是同一个文件（未切换其他 Excel），则覆盖更新上次写入的行，否则追加新记录
  async function onSaveMemo(rows: Array<{
    no: string; date: string; name: string; unit: string; qty: number; price: number; amount: number | ''; person: string; remark: string
  }>) {
    if (!activeFile) return
    setMemoSaving(true)
    setMemoStatus('')
    const isUpdate = lastSavedFile.current === activeFile.filePath
    try {
      const result = isUpdate
        ? await api.updateMemo(activeFile.filePath, rows)
        : await api.appendMemo(activeFile.filePath, rows)
      if (result.error) {
        setMemoStatus('保存失败：' + (result.message || '未知错误'))
      } else {
        lastSavedFile.current = activeFile.filePath
        setMemoStatus(result.message || '已保存')
        refreshHistory()
      }
    } catch (err) {
      setMemoStatus('保存失败：' + (err instanceof Error ? err.message : '未知错误'))
    } finally {
      setMemoSaving(false)
    }
  }

  // ---- 保存并前进到下一条 ----
  async function onSaveAndNext(rows: Array<{
    no: string; date: string; name: string; unit: string; qty: number; price: number; amount: number | ''; person: string; remark: string
  }>) {
    await onSaveMemo(rows)
    if (!activeFile) return
    const nextIdx = filtered.findIndex((f) => f.filePath === activeFile.filePath) + 1
    if (nextIdx < filtered.length) {
      setActiveFile(filtered[nextIdx])
      setMemoStatus('')
    } else {
      setMemoStatus('已是最后一条文件')
      setTimeout(() => setMemoStatus(''), 2500)
    }
  }

  // 实时过滤：每次输入字符即在内存索引中过滤
  const filtered = useMemo(
    () => filterFiles(files, query),
    [files, query],
  )

  return (
    <div className="app">
      <div className="app-top">
        <h1 className="app-title">账单查找器</h1>
        <DirectoryBar
          workDir={workDir}
          status={status}
          loading={loading}
          onChoose={onChooseDir}
          onRefresh={onRefresh}
          onOpenDir={onOpenDir}
        />
      </div>

      <div className="app-body">
        <aside className="app-side">
          <input
            ref={searchRef}
            className="search-input"
            type="text"
            placeholder="搜文件名 / 路径 / 拼音（如 wjy 找王金玉）"
            value={query}
            onChange={(e) => setQuery(e.target.value)}
          />
          <ErrorMessage message={error} onClear={() => setError('')} />
          <ResultList
            api={api}
            files={filtered}
            totalCount={files.length}
            filteredCount={filtered.length}
            query={query}
            loading={loading}
            activeFile={activeFile}
            onOpen={onOpen}
            onOpenInExcel={onOpenInExcel}
          />
          <HistoryList
            records={history}
            onSelect={onHistorySelect}
            onClear={onClearHistory}
          />
        </aside>

        <main className="app-main">
          {activeFile ? (
            <MemoPanel
              key={activeFile.filePath}
              file={activeFile}
              api={api}
              saving={memoSaving}
              status={memoStatus}
              onSave={onSaveMemo}
              onSaveAndNext={onSaveAndNext}
              onClose={onCloseMemo}
            />
          ) : (
            <div className="idle-hint">
              <div className="idle-title">左侧选中文件，开始录入</div>
              <ul className="idle-list">
                <li><b>←→↑↓</b>　方向键切换焦点</li>
                <li><b>Enter</b>　向右移动，到行尾 / 最后一行自动换到下一行起始</li>
                <li><b>Ctrl+Enter</b>　保存到当前文件</li>
                <li><b>Ctrl+Shift+Enter</b>　保存并跳到下一个文件</li>
                <li>点进空格自动弹出<b>最近输入</b>，<b>Alt+↓</b> 手动弹出</li>
                <li>搜索支持拼音首字母，如 <b>wjy</b> 找「王金玉」</li>
              </ul>
            </div>
          )}
        </main>
      </div>
    </div>
  )
}
