import { useEffect, useMemo, useRef, useState } from 'react'
import { ElectronAPI, FileEntry, PreviewRow } from '../types'
import { personMatches } from '../utils/person'

interface Props {
  api: ElectronAPI
  files: FileEntry[]
  totalCount: number
  filteredCount: number
  query: string
  loading: boolean
  activeFile: FileEntry | null
  recognizedPersons?: string[]
  onOpen: (index: number) => void
  onOpenInExcel?: (index: number) => void
}

interface PreviewState {
  file: FileEntry
  data: {
    sheetName?: string
    totalRows?: number
    headerLabels?: string[]
    rows?: PreviewRow[]
    error?: boolean
    message?: string
  }
  pos: { left: number; top: number }
  loading: boolean
}

export function ResultList({
  api,
  files,
  totalCount,
  filteredCount,
  query,
  loading,
  activeFile,
  recognizedPersons = [],
  onOpen,
  onOpenInExcel,
}: Props) {
  // ---- 单击/双击节流 ----
  const timerRef = useRef<ReturnType<typeof setTimeout> | null>(null)
  const lastIdxRef = useRef<number | null>(null)
  const [, setTick] = useState(0)

  const handleClick = (index: number) => {
    if (timerRef.current) clearTimeout(timerRef.current)
    lastIdxRef.current = index
    timerRef.current = setTimeout(() => {
      if (lastIdxRef.current === index) {
        onOpen(index)
        setTick((t) => t + 1)
      }
    }, 300)
  }
  const handleDoubleClick = (index: number) => {
    if (timerRef.current) { clearTimeout(timerRef.current); timerRef.current = null }
    lastIdxRef.current = null
    onOpenInExcel?.(index)
  }

  // ---- 悬停预览开关（默认关闭，避免弹出挡视线；选择持久化） ----
  const PREVIEW_KEY = 'bill-explorer:preview-enabled'
  const [previewEnabled, setPreviewEnabled] = useState<boolean>(() => {
    try {
      return localStorage.getItem(PREVIEW_KEY) === '1'
    } catch {
      return false
    }
  })

  function togglePreview() {
    const next = !previewEnabled
    setPreviewEnabled(next)
    try {
      localStorage.setItem(PREVIEW_KEY, next ? '1' : '0')
    } catch {
      /* ignore */
    }
    if (!next) setPreview(null)
  }

  // ---- 悬停预览 ----
  const [preview, setPreview] = useState<PreviewState | null>(null)
  const wrapRef = useRef<HTMLDivElement | null>(null)
  const previewTimer = useRef<ReturnType<typeof setTimeout> | null>(null)
  const hideTimer = useRef<ReturnType<typeof setTimeout> | null>(null)

  // ---- 右键菜单（文件列表项） ----
  const [ctxMenu, setCtxMenu] = useState<{ x: number; y: number; index: number } | null>(null)

  function handleItemContextMenu(e: React.MouseEvent, index: number) {
    e.preventDefault()
    e.stopPropagation()
    const MW = 200
    const MH = 132
    const x = Math.min(e.clientX, window.innerWidth - MW - 8)
    const y = Math.min(e.clientY, window.innerHeight - MH - 8)
    setCtxMenu({ x: Math.max(4, x), y: Math.max(4, y), index })
  }

  // 菜单打开时，点击别处 / 滚动 / Esc / 失焦即关闭
  useEffect(() => {
    if (!ctxMenu) return
    const close = () => setCtxMenu(null)
    const onKey = (ev: KeyboardEvent) => { if (ev.key === 'Escape') close() }
    window.addEventListener('mousedown', close)
    window.addEventListener('resize', close)
    window.addEventListener('blur', close)
    window.addEventListener('keydown', onKey)
    window.addEventListener('scroll', close, true)
    return () => {
      window.removeEventListener('mousedown', close)
      window.removeEventListener('resize', close)
      window.removeEventListener('blur', close)
      window.removeEventListener('keydown', onKey)
      window.removeEventListener('scroll', close, true)
    }
  }, [ctxMenu])

  function getPos(index: number) {
    const item = wrapRef.current?.querySelector(`[data-file-idx="${index}"]`) as HTMLElement
    if (!item) return { left: 10, top: 10 }
    const rect = item.getBoundingClientRect()
    const cardWidth = 720
    const rightSpace = window.innerWidth - rect.right
    const leftSpace = rect.left
    let left: number
    if (rightSpace >= cardWidth + 10) {
      left = rect.right + 10
    } else if (leftSpace >= cardWidth + 10) {
      left = rect.left - cardWidth - 10
    } else {
      left = Math.max(8, (window.innerWidth - cardWidth) / 2)
    }
    const top = Math.min(rect.top, window.innerHeight - 440)
    return { left, top: Math.max(8, top) }
  }

  function loadPreview(file: FileEntry, index: number) {
    if (!file.fileName.toLowerCase().endsWith('.xlsx')) {
      setPreview({
        file,
        data: { error: true, message: '仅支持 .xlsx 文件预览' },
        pos: getPos(index),
        loading: false,
      })
      return
    }
    setPreview({ file, data: {}, pos: getPos(index), loading: true })
    api.previewRows(file.filePath, 10)
      .then((result) => setPreview({ file, data: result, pos: getPos(index), loading: false }))
      .catch(() => setPreview({ file, data: { error: true, message: '读取失败' }, pos: getPos(index), loading: false }))
  }

  const handleItemEnter = (index: number) => {
    if (!previewEnabled) return
    const file = files[index]
    if (!file) return
    clearTimeout(previewTimer.current || 0)
    clearTimeout(hideTimer.current || 0)
    previewTimer.current = setTimeout(() => loadPreview(file, index), 120)
  }

  const handleItemLeave = () => {
    // 仅取消待加载的预览请求，不在此隐藏卡片（由卡片的 mouseLeave 处理）
    clearTimeout(previewTimer.current || 0)
    previewTimer.current = null
  }

  const handlePreviewEnter = () => {
    // 取消任何待执行的隐藏
    clearTimeout(hideTimer.current || 0)
  }

  const handlePreviewLeave = () => {
    clearTimeout(hideTimer.current || 0)
    hideTimer.current = setTimeout(() => {
      setPreview(null)
    }, 300)
  }

  // 滚动时重新定位
  useEffect(() => {
    const onScroll = () => {
      if (!preview) return
      const idx = files.findIndex((f) => f.filePath === preview.file.filePath)
      if (idx >= 0) {
        setPreview((p) => (p ? { ...p, pos: getPos(idx) } : null))
      }
    }
    window.addEventListener('scroll', onScroll, true)
    return () => window.removeEventListener('scroll', onScroll, true)
  }, [preview, files])

  // 把"识图识别出的人名"对应的 Excel 排到最上方；其余文件保持原顺序
  const display = useMemo(() => {
    if (!recognizedPersons.length) {
      return files.map((file, idx) => ({ file, idx, hit: false }))
    }
    const matched: Array<{ file: FileEntry; idx: number; hit: boolean }> = []
    const others: Array<{ file: FileEntry; idx: number; hit: boolean }> = []
    files.forEach((file, idx) => {
      const hit = recognizedPersons.some((p) =>
        personMatches(file.fileName.replace(/\.(xlsx|xls)$/i, ''), p),
      )
      ;(hit ? matched : others).push({ file, idx, hit })
    })
    return [...matched, ...others]
  }, [files, recognizedPersons])

  if (!files.length && totalCount === 0 && !loading) {
    return (
      <div className="empty-state">
        尚未扫描到 Excel 文件，请先选择工作目录并刷新。
      </div>
    )
  }

  if (!files.length && query.trim()) {
    return (
      <div className="empty-state">
        <span>无匹配结果</span>
        <br />
        <span style={{ fontSize: 11, color: '#c0c4cc', marginTop: 6, display: 'block' }}>
          已尝试 文件名 / 路径 / 全拼 / 拼音首字母 匹配，仍未找到
        </span>
      </div>
    )
  }

  return (
    <div ref={wrapRef} className="result-wrap">
      <div className="result-stats">
        <span>
          {query.trim()
            ? `共 ${totalCount} 个文件，匹配 ${filteredCount} 个`
            : `共 ${totalCount} 个文件`}
        </span>
        <span className="result-tip">单击录入 · 双击 Excel 打开</span>
        <button
          className={previewEnabled ? 'preview-toggle is-on' : 'preview-toggle'}
          onClick={togglePreview}
          title={previewEnabled ? '关闭悬停预览（默认关闭，避免挡视线）' : '开启悬停预览'}
        >
          悬停预览 {previewEnabled ? '开' : '关'}
        </button>
      </div>
      <ul className="file-list">
        {recognizedPersons.length > 0 && display.some((d) => d.hit) && (
          <li className="hit-group-header" key="__hit">
            识图命中 {display.filter((d) => d.hit).length} 个（按识别人名置顶）
          </li>
        )}
        {display.map(({ file, idx, hit }) => (
          <li
            data-file-idx={idx}
            className={
              (file.filePath === activeFile?.filePath ? 'file-item is-active' : 'file-item') +
              (hit ? ' is-hit' : '')
            }
            key={file.filePath}
            title="单击打开录入，双击用 Excel 打开，右键更多操作"
            onClick={() => handleClick(idx)}
            onDoubleClick={(e) => {
              e.stopPropagation()
              handleDoubleClick(idx)
            }}
            onContextMenu={(e) => handleItemContextMenu(e, idx)}
            onMouseEnter={() => handleItemEnter(idx)}
            onMouseLeave={handleItemLeave}
          >
            <div className="file-name">
              {file.fileName}
              {hit && <span className="hit-badge">命中</span>}
            </div>
            <div className="file-path">{file.filePath}</div>
          </li>
        ))}
      </ul>

      {preview && (
        <div
          className="preview-card"
          style={{ left: preview.pos.left, top: preview.pos.top }}
          onMouseEnter={handlePreviewEnter}
          onMouseLeave={handlePreviewLeave}
        >
          <div className="preview-header">
            <span className="preview-file">{preview.file.fileName}</span>
            {preview.data.sheetName && preview.data.totalRows !== undefined && (
              <span className="preview-meta">
                {preview.data.sheetName} · {preview.data.totalRows} 行
              </span>
            )}
            <button
              className="preview-close"
              title="关闭预览"
              onClick={() => setPreview(null)}
            >
              ✕
            </button>
          </div>
          {preview.loading ? (
            <div className="preview-body">加载中...</div>
          ) : preview.data.error && preview.data.message ? (
            <div className="preview-body preview-error">{preview.data.message}</div>
          ) : preview.data.rows && preview.data.rows.length > 0 ? (
            <div className="preview-body">
              <table className="preview-table">
                <colgroup>
                  {Array.from({ length: 9 }, () => (
                    <col style={{ width: `${720 / 9}px` }} />
                  ))}
                </colgroup>
                <thead>
                  <tr>
                    {((preview.data.headerLabels && preview.data.headerLabels.length > 0)
                        ? preview.data.headerLabels.slice(0, 9)
                        : Array.from({ length: Math.min(9, (preview.data.rows[0] && preview.data.rows[0].cells.length) || 9) }, () => ''))
                      .map((h, i) => <th key={i}>{h}</th>)}
                  </tr>
                </thead>
                <tbody>
                  {preview.data.rows!.map((row, i) => (
                    <tr key={i}>
                      {row.cells.slice(0, 9).map((cell, j) => (
                        <td key={j}>{cell.value}</td>
                      ))}
                      {row.cells.length > 9 && (
                        <td style={{ color: '#909399', textAlign: 'center' }}>…</td>
                      )}
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          ) : (
            <div className="preview-body">暂无数据</div>
          )}
        </div>
      )}

      {ctxMenu && (() => {
        const f = files[ctxMenu.index]
        if (!f) return null
        const item = (label: string, hint: string, fn: () => void) => (
          <button
            className="ctx-item"
            onMouseDown={(e) => e.stopPropagation()}
            onClick={() => { setCtxMenu(null); fn() }}
          >
            <span>{label}</span>
            {hint && <span className="ctx-hint">{hint}</span>}
          </button>
        )
        return (
          <div className="ctx-menu" style={{ left: ctxMenu.x, top: ctxMenu.y }}>
            {item('打开文件所在文件夹', '', () => { void api.revealFile(f.filePath) })}
            {item('用 Excel 打开', '', () => { onOpenInExcel?.(ctxMenu.index) })}
            <div className="ctx-sep" />
            {item('打开录入', '单击', () => { onOpen(ctxMenu.index) })}
          </div>
        )
      })()}
    </div>
  )
}