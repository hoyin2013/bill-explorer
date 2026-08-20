import { useEffect, useMemo, useRef, useState } from 'react'
import { ElectronAPI, FileEntry, HistoryRecord, ImageSnapshot } from './types'
import { filterFiles } from './utils'
import { DirectoryBar } from './components/DirectoryBar'
import { ResultList } from './components/ResultList'
import { UniverSheet } from './components/UniverSheet'
import { ImageWindow, type ImageWindowHandle } from './components/ImageWindow'
import { ErrorMessage } from './components/ErrorMessage'
import { HistoryList } from './components/HistoryList'
import { SettingsModal } from './components/SettingsModal'

// 安全检查：electronAPI 来自 preload，若为 undefined 说明 Electron 未正确加载 preload
const api = window.electronAPI

// 当 api 缺失时展示明确错误，避免用户看到一片空白
function BrokenState() {
  return (
    <div className="app">
      <h1 className="app-title">账单录入器</h1>
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

  // 通过 ?detached=image 查询参数识别「独立识图窗口」模式：只渲染 ImageWindow 全屏
  const detached = new URLSearchParams(window.location.search).get('detached') === 'image'
  return detached ? <DetachedImageApp api={api} /> : <AppInner api={api} />
}

// 独立「小票识图」窗口：仅渲染 ImageWindow，并带「合并回主窗口」能力
function DetachedImageApp({ api }: { api: ElectronAPI }) {
  const [init, setInit] = useState<ImageSnapshot | null>(null)
  const [failed, setFailed] = useState(false)
  useEffect(() => {
    let alive = true
    let settled = false
    api
      .getDetachedInit()
      .then((s) => {
        settled = true
        if (alive && s) setInit(s as ImageSnapshot)
      })
      .catch(() => {
        settled = true
        if (alive) setFailed(true)
      })
    // 兜底：若 3 秒内仍未取回快照（极端情况），升级为空白面板渲染，避免一直卡在加载
    const t = setTimeout(() => {
      if (alive && !settled) setFailed(true)
    }, 3000)
    return () => {
      alive = false
      clearTimeout(t)
    }
  }, [api])

  // 取回快照前先不渲染内容，避免「先空白再填充」的闪烁；
  // 快照到位后 ImageWindow 首帧即带 initialState，配合 useLayoutEffect 在绘制前完成恢复。
  if (!init) {
    return (
      <div className="detached-root">
        <div className="detached-loading">
          {failed ? '未能载入识图内容，请重新点「拆分窗口」' : '正在载入识图内容…'}
        </div>
      </div>
    )
  }
  return (
    <div className="detached-root">
      <ImageWindow
        api={api}
        detached
        initialState={init}
        onAttach={(state) => {
          api.attachImageDetached(state)
          window.close()
        }}
      />
    </div>
  )
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

  // ---- AI / 图片设置 ----
  const [settingsOpen, setSettingsOpen] = useState(false)

  // 右侧"小票识图"面板是否展开 + 其宽度（px），可拖动分隔条调整
  const [showImagePanel, setShowImagePanel] = useState(true)
  const [imgWidth, setImgWidth] = useState(360)
  const imgSplitDragRef = useRef<{ startX: number; startW: number } | null>(null)
  // 是否已把识图面板拆成独立窗口（拆分后主窗口面板隐藏但保持挂载，以保留结果与进度）
  const [isDetached, setIsDetached] = useState(false)
  const imageWinRef = useRef<ImageWindowHandle>(null)

  // 识图窗口上报的"已识别人名"，用于把左侧列表中对应 Excel 置顶
  const [recognizedPersons, setRecognizedPersons] = useState<string[]>([])
  // 当前是否在主窗口打开了录入面板（供"识图回填"反馈判断）
  const activeFileRef = useRef<FileEntry | null>(null)
  activeFileRef.current = activeFile

  // ---- 左侧查找区宽度（px），可拖动中间分隔条调整（参照"最近修改"窗口的可调设计） ----
  const [sideWidth, setSideWidth] = useState(200)
  const splitDragRef = useRef<{ startX: number; startW: number } | null>(null)

  const searchRef = useRef<HTMLInputElement>(null)
  // 记录最近一次保存的文件路径（用于"同一文件连续编辑 → 覆盖更新"）
  const lastSavedFile = useRef<string | null>(null)

  // 切换文件前：等待 UniverSheet 完成保存（若有未保存修改）
  const sheetSaveBeforeSwitchRef = useRef<() => Promise<unknown>>(async () => {})

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

  // 软件启动时自动加载上次持久化的工作目录和设置
  useEffect(() => {
    loadWorkDir()
    refreshHistory()
    loadSettings()
  }, [])

  // 接收独立识图窗口上报的人名，用于把左侧列表对应 Excel 置顶
  useEffect(() => {
    const off = api.on('recognized-persons', (persons) => {
      setRecognizedPersons(Array.isArray(persons) ? (persons as string[]) : [])
    })
    return off
  }, [api])

  // 识图窗口点"填入当前录入"时，若主窗口尚未打开任何账单文件，给出引导提示
  useEffect(() => {
    const off = api.on('apply-recognized-rows', (rows) => {
      if (!activeFileRef.current) {
        const n = Array.isArray(rows) ? rows.length : 0
        setMemoStatus(`已收到识图窗口的 ${n} 条识别结果，但左侧还没打开目标 Excel。请先选中一个文件进入录入，再点网格即可核对后保存。`)
        setTimeout(() => setMemoStatus(''), 5000)
      }
    })
    return off
  }, [api])

  // 独立「小票识图」窗口关闭时，取消 detached 标记，右侧面板恢复可见（状态始终保留在隐藏面板中）
  useEffect(() => {
    const off = api.on('image-detached-closed', () => setIsDetached(false))
    return off
  }, [api])

  // 独立窗口「合并回主窗口」时，把独立窗口里最新的结果/进度同步回主面板
  useEffect(() => {
    const off = api.on('image-detached-merged', (state) => {
      imageWinRef.current?.restoreState(state as ImageSnapshot)
    })
    return off
  }, [api])

  // 把右侧「小票识图」面板拆成独立窗口（自身先隐藏，但保持挂载以保留结果与进度）
  function handleDetach(state: ImageSnapshot) {
    setIsDetached(true)
    api.openImageDetached(state)
  }

  async function loadSettings() {
    try {
      await api.getSettings()
    } catch {
      // ignore
    }
  }

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
  // 若有未保存修改，先保存当前文件再打开下一个，避免静默丢数据（修复 U5）
  function onOpen(index: number) {
    const file = filtered[index]
    if (!file) return
    const prev = activeFileRef.current
    if (prev && file.filePath !== prev.filePath) {
      // 等 UniverSheet 保存完当前文件，再切换 key → 销毁旧实例（不等则旧 sheet 卸载时数据丢失）
      sheetSaveBeforeSwitchRef.current().then(() => {
        setActiveFile(file)
        setMemoStatus('')
      })
      return
    }
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
    setQuery('') // 关闭录入面板即清空左侧文件搜索框，避免残留检索词
    setMemoStatus('')
    lastSavedFile.current = null
    // 关闭后把焦点交还左侧搜索框：编辑表格时 Univer 的隐藏编辑器可能
    // 仍抢占焦点，导致搜索框“卡住无法输入”；显式聚焦保证其可用，
    // 也符合“正常关闭后焦点定位到搜索框”的预期行为。
    requestAnimationFrame(() => searchRef.current?.focus())
  }

  // ---- 点击历史记录：若在当前扫描结果中则打开录入，否则用系统程序打开 ----
  // 若有未保存修改，先保存再打开，避免静默丢数据（修复 U5）
  async function onHistorySelect(filePath: string) {
    const file = files.find((f) => f.filePath === filePath)
    if (file) {
      const prev = activeFileRef.current
      if (prev && filePath !== prev.filePath) {
        await sheetSaveBeforeSwitchRef.current()
      }
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

  // ---- 拖动中间分隔条，左右调节查找区宽度 ----
  function onSplitterDown(e: React.MouseEvent) {
    e.preventDefault()
    e.stopPropagation()
    splitDragRef.current = { startX: e.clientX, startW: sideWidth }
    const onMove = (ev: MouseEvent) => {
      const d = splitDragRef.current
      if (!d) return
      // 鼠标向右拉使左侧更宽，向左拉使其更窄；限制在合理范围内
      const w = Math.max(160, Math.min(window.innerWidth * 0.7, d.startW + (ev.clientX - d.startX)))
      setSideWidth(w)
    }
    const onUp = () => {
      splitDragRef.current = null
      window.removeEventListener('mousemove', onMove)
      window.removeEventListener('mouseup', onUp)
    }
    window.addEventListener('mousemove', onMove)
    window.addEventListener('mouseup', onUp)
  }

  // ---- 拖动右侧分隔条，调节「小票识图」面板宽度（向左拖变宽） ----
  function onImgSplitterDown(e: React.MouseEvent) {
    e.preventDefault()
    e.stopPropagation()
    imgSplitDragRef.current = { startX: e.clientX, startW: imgWidth }
    const onMove = (ev: MouseEvent) => {
      const d = imgSplitDragRef.current
      if (!d) return
      // 鼠标向左移使右侧面板更宽，向右移使其更窄；限制在合理范围内
      const w = Math.max(260, Math.min(window.innerWidth * 0.6, d.startW - (ev.clientX - d.startX)))
      setImgWidth(w)
    }
    const onUp = () => {
      imgSplitDragRef.current = null
      window.removeEventListener('mousemove', onMove)
      window.removeEventListener('mouseup', onUp)
    }
    window.addEventListener('mousemove', onMove)
    window.addEventListener('mouseup', onUp)
  }

  // ---- 保存：把面板内的多行数据一次性写回 Excel ----
  // 若与上次保存的是同一个文件（未切换其他 Excel），则覆盖更新上次写入的行，否则追加新记录
  // 返回是否保存成功（供「保存并下一条」复用，避免在推进逻辑里重复这段）
  async function doSaveMemo(rows: Array<{
    no: string; date: string; name: string; unit: string; qty: number; price: number; amount: number | ''; person: string; remark: string
  }>): Promise<boolean> {
    if (!activeFile) return false
    setMemoSaving(true)
    setMemoStatus('')
    const isUpdate = lastSavedFile.current === activeFile.filePath
    try {
      const result = isUpdate
        ? await api.updateMemo(activeFile.filePath, rows)
        : await api.appendMemo(activeFile.filePath, rows)
      if (result.error) {
        setMemoStatus('保存失败：' + (result.message || '未知错误'))
        return false
      } else {
        lastSavedFile.current = activeFile.filePath
        setMemoStatus(result.message || '已保存')
        refreshHistory()
        return true
      }
    } catch (err) {
      setMemoStatus('保存失败：' + (err instanceof Error ? err.message : '未知错误'))
      return false
    } finally {
      setMemoSaving(false)
    }
  }

  // 普通保存：成功后清空左侧搜索框（回到全部文件视图，方便继续找下一张）
  async function onSaveMemo(rows: Array<{
    no: string; date: string; name: string; unit: string; qty: number; price: number; amount: number | ''; person: string; remark: string
  }>) {
    const ok = await doSaveMemo(rows)
    if (ok) setQuery('')
  }

  // ---- 保存并前进到下一条 ----
  async function onSaveAndNext(rows: Array<{
    no: string; date: string; name: string; unit: string; qty: number; price: number; amount: number | ''; person: string; remark: string
  }>) {
    const ok = await doSaveMemo(rows)
    if (!ok || !activeFile) return
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
        <h1 className="app-title">账单录入器</h1>
        <DirectoryBar
          workDir={workDir}
          status={status}
          loading={loading}
          onChoose={onChooseDir}
          onRefresh={onRefresh}
          onOpenDir={onOpenDir}
        />
        <div className="app-top-actions">
          <button
            className={'btn ' + (showImagePanel ? 'btn-primary' : 'btn-outline')}
            onClick={() => setShowImagePanel((v) => !v)}
            title={showImagePanel ? '收起右侧「小票识图」面板' : '展开右侧「小票识图」面板'}
          >
            小票识图
          </button>
          <button className="btn btn-outline" onClick={() => setSettingsOpen(true)} title="配置 AI 接口、图片目录、识别提示词">
            AI 设置
          </button>
        </div>
      </div>

      <div className="app-body">
        <aside className="app-side" style={{ width: sideWidth }}>
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
            recognizedPersons={recognizedPersons}
            onOpen={onOpen}
            onOpenInExcel={onOpenInExcel}
          />
          <HistoryList
            records={history}
            onSelect={onHistorySelect}
            onClear={onClearHistory}
          />
        </aside>

        <div className="pane-splitter" onMouseDown={onSplitterDown} title="拖动调整左右宽度" />

        <main className="app-main">
          {activeFile ? (
            <UniverSheet
              key={activeFile.filePath}
              file={activeFile}
              api={api}
              onClose={onCloseMemo}
              onSaved={() => {
                setQuery('')      // 保存后清空左侧搜索框，回到全部文件视图
                refreshHistory()
              }}
              onRegisterSwitchFile={(fn) => { sheetSaveBeforeSwitchRef.current = fn }}
            />
          ) : (
            <div className="idle-hint">
              <div className="idle-title">左侧选中文件，开始录入</div>
              <ul className="idle-list">
                <li><b>打开文件</b>　整张表直接显示，光标自动落在底部空行</li>
                <li><b>↑↓←→</b>　在单元格间移动焦点</li>
                <li><b>Enter / 双击</b>　进入编辑；编辑中 <b>Enter</b> 下一行、<b>Tab</b> 右移</li>
                <li><b>Ctrl+C / V / X</b>　复制 / 粘贴 / 剪切，支持从 Excel 直接粘贴整块</li>
                <li><b>日期列</b>　直接输入或粘贴，支持 2026/8/11、8-11 等多种格式自动转换</li>
                <li><b>Ctrl+S</b>　保存（覆盖整表，自动去掉尾部空行）</li>
                <li>搜索支持拼音首字母，如 <b>wjy</b> 找「王金玉」</li>
              </ul>
            </div>
          )}
        </main>

        {(showImagePanel || isDetached) && (
          <>
            {!isDetached && (
              <div
                className="pane-splitter"
                onMouseDown={onImgSplitterDown}
                title="拖动调整右侧「小票识图」面板宽度"
              />
            )}
            <aside
              className={'image-panel-col' + (isDetached ? ' detached-hidden' : '')}
              style={{ width: imgWidth }}
            >
              <ImageWindow ref={imageWinRef} api={api} onDetach={handleDetach} />
            </aside>
          </>
        )}
      </div>

      <SettingsModal
        api={api}
        open={settingsOpen}
        onClose={() => setSettingsOpen(false)}
        onSaved={loadSettings}
      />
    </div>
  )
}
