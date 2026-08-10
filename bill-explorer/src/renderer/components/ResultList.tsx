import { useEffect, useRef, useState } from 'react'
import { ElectronAPI, FileEntry, PreviewRow } from '../types'

interface Props {
  api: ElectronAPI
  files: FileEntry[]
  totalCount: number
  filteredCount: number
  query: string
  loading: boolean
  activeFile: FileEntry | null
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

  // ---- 悬停预览 ----
  const [preview, setPreview] = useState<PreviewState | null>(null)
  const wrapRef = useRef<HTMLDivElement | null>(null)
  const previewTimer = useRef<ReturnType<typeof setTimeout> | null>(null)
  const hideTimer = useRef<ReturnType<typeof setTimeout> | null>(null)

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

  if (!files.length && totalCount === 0 && !loading) {
    return (
      <div className="empty-state">
        尚未扫描到 Excel 文件，请先选择工作目录并刷新。
      </div>
    )
  }

  if (!files.length && query.trim()) {
    return <div className="empty-state">无匹配结果</div>
  }

  return (
    <div ref={wrapRef} className="result-wrap">
      <div className="result-stats">
        {query.trim()
          ? `共 ${totalCount} 个文件，匹配 ${filteredCount} 个`
          : `共 ${totalCount} 个文件`}
        <span className="result-tip">单击录入 · 双击 Excel 打开 · 悬停预览</span>
      </div>
      <ul className="file-list">
        {files.map((f, i) => (
          <li
            data-file-idx={i}
            className={
              f.filePath === activeFile?.filePath ? 'file-item is-active' : 'file-item'
            }
            key={f.filePath + i}
            title="单击打开录入，双击用 Excel 打开，悬停预览"
            onClick={() => handleClick(i)}
            onDoubleClick={(e) => {
              e.stopPropagation()
              handleDoubleClick(i)
            }}
            onMouseEnter={() => handleItemEnter(i)}
            onMouseLeave={handleItemLeave}
          >
            <div className="file-name">{f.fileName}</div>
            <div className="file-path">{f.filePath}</div>
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
    </div>
  )
}