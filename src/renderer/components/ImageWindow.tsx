import { useEffect, useRef, useState, type MouseEvent as ReactMouseEvent } from 'react'
import { ElectronAPI, AIRecognizedRow } from '../types'

interface Props {
  api: ElectronAPI
}

// 从一批识别结果里提取去重后的人名
function uniquePersons(rows: AIRecognizedRow[]): string[] {
  const set = new Set<string>()
  for (const r of rows) {
    const p = (r.person || '').trim()
    if (p) set.add(p)
  }
  return Array.from(set)
}

// 取图片文件名（去掉路径）作为"来源"标记
function baseName(path: string): string {
  const idx = Math.max(path.lastIndexOf('/'), path.lastIndexOf('\\'))
  return idx >= 0 ? path.slice(idx + 1) : path
}

// 结果表列定义（与账单 9 列一致）
const RESULT_COLS: Array<{ field: keyof AIRecognizedRow; label: string; w: string }> = [
  { field: 'no', label: '序号', w: '42px' },
  { field: 'date', label: '日期', w: '92px' },
  { field: 'name', label: '货品', w: 'auto' },
  { field: 'unit', label: '单位', w: '54px' },
  { field: 'qty', label: '数量', w: '54px' },
  { field: 'price', label: '单价', w: '70px' },
  { field: 'amount', label: '金额', w: '82px' },
  { field: 'person', label: '调货人', w: '88px' },
  { field: 'remark', label: '备注', w: '130px' },
]

export function ImageWindow({ api }: Props) {
  const [imageDir, setImageDir] = useState('')
  const [images, setImages] = useState<Array<{ name: string; path: string }>>([])
  const [selected, setSelected] = useState('')
  const [preview, setPreview] = useState('')
  const [rotate, setRotate] = useState(0)
  const [loading, setLoading] = useState(false)
  const [aiLoading, setAiLoading] = useState(false)
  const [scanning, setScanning] = useState(false)
  const [status, setStatus] = useState('')
  // 完整识别结果（不再截断），既用于展示，也用于"填入当前录入"回填
  const [rows, setRows] = useState<AIRecognizedRow[]>([])
  const [recognized, setRecognized] = useState<string[]>([])
  const scanningRef = useRef(false)
  const cancelRef = useRef(false)
  // 预览缩放 / 平移状态
  const [zoom, setZoom] = useState(1)
  const [pan, setPan] = useState<{ x: number; y: number }>({ x: 0, y: 0 })
  const [isDragging, setIsDragging] = useState(false)
  const wrapRef = useRef<HTMLDivElement>(null)
  const zoomRef = useRef(1)
  const panRef = useRef({ x: 0, y: 0 })
  const draggingRef = useRef<{ startX: number; startY: number; panX: number; panY: number } | null>(null)

  // 把当前已识别的人名实时上报给主窗口（主窗口据此把对应 Excel 置顶）
  useEffect(() => {
    api.reportPersons(recognized)
  }, [recognized, api])

  // 同步最新的 zoom/pan 到 ref（供原生事件处理器读取）
  useEffect(() => {
    zoomRef.current = zoom
  }, [zoom])
  useEffect(() => {
    panRef.current = pan
  }, [pan])

  // 滚轮缩放（以光标为锚点）
  useEffect(() => {
    const el = wrapRef.current
    if (!el) return
    const onWheel = (e: WheelEvent) => {
      if (!preview) return
      e.preventDefault()
      const rect = el.getBoundingClientRect()
      const cx = e.clientX - (rect.left + rect.width / 2)
      const cy = e.clientY - (rect.top + rect.height / 2)
      const z = zoomRef.current
      const factor = e.deltaY < 0 ? 1.12 : 1 / 1.12
      const newZoom = Math.min(8, Math.max(0.2, z * factor))
      if (newZoom === z) return
      const p = panRef.current
      const newPan = {
        x: cx - (newZoom / z) * (cx - p.x),
        y: cy - (newZoom / z) * (cy - p.y),
      }
      zoomRef.current = newZoom
      panRef.current = newPan
      setZoom(newZoom)
      setPan(newPan)
    }
    el.addEventListener('wheel', onWheel, { passive: false })
    return () => el.removeEventListener('wheel', onWheel)
  }, [preview])

  // 左键拖拽平移（监听 window，拖出容器也能继续）
  useEffect(() => {
    const onMove = (e: MouseEvent) => {
      const d = draggingRef.current
      if (!d) return
      const nx = d.panX + (e.clientX - d.startX)
      const ny = d.panY + (e.clientY - d.startY)
      panRef.current = { x: nx, y: ny }
      setPan({ x: nx, y: ny })
    }
    const onUp = () => {
      if (draggingRef.current) {
        draggingRef.current = null
        setIsDragging(false)
      }
    }
    window.addEventListener('mousemove', onMove)
    window.addEventListener('mouseup', onUp)
    return () => {
      window.removeEventListener('mousemove', onMove)
      window.removeEventListener('mouseup', onUp)
    }
  }, [])

  async function loadImages(dir?: string) {
    const target = dir || imageDir
    if (!target) return
    setImageDir(target)
    setLoading(true)
    setStatus('')
    try {
      const res = await api.listImages(target)
      if (res.error) {
        setStatus(res.message || '读取图片列表失败')
        setImages([])
      } else {
        const imgs = res.images || []
        setImages(imgs)
        if (imgs.length) {
          setSelected(imgs[0].path)
          await loadPreview(imgs[0].path)
        } else {
          setStatus('该图片目录下没有图片')
        }
      }
    } catch (err) {
      setStatus('读取图片列表失败：' + (err instanceof Error ? err.message : ''))
    } finally {
      setLoading(false)
    }
  }

  async function loadPreview(path: string) {
    setSelected(path)
    setRotate(0)
    setRows([])
    setPreview('')
    zoomRef.current = 1
    panRef.current = { x: 0, y: 0 }
    setZoom(1)
    setPan({ x: 0, y: 0 })
    const res = await api.readImageBase64(path)
    if (res.error) {
      setStatus(res.message || '读取图片失败')
    } else if (res.base64 && res.mime) {
      setPreview(`data:${res.mime};base64,${res.base64}`)
    }
  }

  // 预览区：左键按下开始拖拽平移
  function onPreviewMouseDown(e: ReactMouseEvent) {
    if (e.button !== 0 || !preview) return
    const p = panRef.current
    draggingRef.current = { startX: e.clientX, startY: e.clientY, panX: p.x, panY: p.y }
    setIsDragging(true)
  }

  // 复位缩放与平移
  function resetView() {
    zoomRef.current = 1
    panRef.current = { x: 0, y: 0 }
    setZoom(1)
    setPan({ x: 0, y: 0 })
  }

  // 行内编辑：更新某一行某一字段
  function updateCell(index: number, field: keyof AIRecognizedRow, value: string) {
    setRows((prev) => {
      const next = prev.slice()
      const row = { ...next[index] }
      if (field === 'no' || field === 'qty' || field === 'price' || field === 'amount') {
        row[field] = value
      } else {
        ;(row as Record<string, string>)[field as string] = value
      }
      next[index] = row
      return next
    })
  }

  // 删除某一行（识别结果里有错行时手动剔除）
  function deleteRow(index: number) {
    setRows((prev) => prev.filter((_, i) => i !== index))
  }

  // 识别单张图片：只展示结果，由用户点「填入当前录入」回填
  async function doRecognize(path: string) {
    if (!path) {
      setStatus('请先选择一张图片')
      return
    }
    setAiLoading(true)
    setStatus('AI 识别中…')
    setRows([])
    try {
      const res = await api.aiRecognize(path)
      if (res.error) {
        setStatus(res.message || '识别失败')
      } else {
        const rs = (res.rows || []).map((r) => ({ ...r, source: baseName(path) }))
        setRows(rs)
        const ps = uniquePersons(rs)
        if (ps.length) {
          setRecognized((prev) => Array.from(new Set([...prev, ...ps])))
          setStatus(`识别完成，共 ${rs.length} 条记录，提取人名：${ps.join('、')}`)
        } else {
          setStatus(`识别完成，共 ${rs.length} 条记录（未识别到人名）`)
        }
        if (res.message) setStatus((s) => s + '\n' + res.message)
      }
    } catch (err) {
      setStatus('识别失败：' + (err instanceof Error ? err.message : '未知错误'))
    } finally {
      setAiLoading(false)
    }
  }

  // 扫描全部图片：逐张识别，聚合并上报所有人名（用于把左侧列表对应文件置顶）
  async function scanAll() {
    if (scanningRef.current) return
    if (!images.length) {
      setStatus('没有图片可扫描')
      return
    }
    scanningRef.current = true
    cancelRef.current = false
    setScanning(true)
    setAiLoading(true)
    const all: AIRecognizedRow[] = []
    try {
      for (let i = 0; i < images.length; i++) {
        if (cancelRef.current) break
        const img = images[i]
        setStatus(`正在扫描图片 ${i + 1}/${images.length}：${img.name}`)
        let res: { error?: boolean; message?: string; rows?: AIRecognizedRow[] } | undefined
        try {
          res = await api.aiRecognize(img.path)
        } catch {
          res = undefined
        }
        if (res && res.rows) {
          for (const r of res.rows) all.push({ ...r, source: img.name })
        }
      }
      const ps = uniquePersons(all)
      setRows(all) // 保留全部结果，不再截断
      setRecognized((prev) => Array.from(new Set([...prev, ...ps])))
      if (cancelRef.current) {
        setStatus(
          ps.length
            ? `已停止扫描（${all.length} 条小票、${ps.length} 个人名）。已把对应 Excel 置顶到左侧列表。`
            : `已停止扫描（${all.length} 条小票）。`,
        )
      } else {
        setStatus(
          ps.length
            ? `扫描完成：${images.length} 张图片、${all.length} 条小票，提取到 ${ps.length} 个人名（${ps.join('、')}）。已把对应 Excel 置顶到左侧列表。`
            : `扫描完成：${images.length} 张图片、${all.length} 条小票，未识别到人名。`,
        )
      }
    } finally {
      scanningRef.current = false
      setScanning(false)
      setAiLoading(false)
    }
  }

  function stopScan() {
    cancelRef.current = true
  }

  async function chooseDir() {
    const dir = await api.selectImageDirectory()
    if (dir) {
      await loadImages(dir)
    }
  }

  function clearHits() {
    setRecognized([])
    setStatus('已清除左侧列表的人名置顶')
  }

  function clearRows() {
    setRows([])
    setStatus('已清空识别结果')
  }

  function applyToMain() {
    if (!rows.length) {
      setStatus('没有可回填的识别结果')
      return
    }
    api.applyToMain(rows)
    setStatus(
      `已把 ${rows.length} 条识别结果发送到「账单录入器」主窗口。\n请在主窗口左侧打开目标 Excel，点击网格即可核对/修改后保存（不会影响当前录入，除非你主动编辑）。`,
    )
  }

  // 初始化：读取设置中的图片目录
  useEffect(() => {
    api
      .getSettings()
      .then((s) => {
        if (s.imageDir) {
          loadImages(s.imageDir)
        } else {
          setStatus('请先选择小票图片目录')
        }
      })
      .catch(() => setStatus('读取设置失败'))
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [])

  // 是否展示"来源"列：仅当结果来自多张不同图片时（单图识别不显示，避免冗余）
  const distinctSources = new Set(rows.map((r) => r.source).filter(Boolean))
  const showSource = distinctSources.size > 1

  return (
    <div className="image-window">
      <div className="image-panel">
        <div className="image-panel-header">
          <span className="image-panel-title">小票识图（独立窗口）</span>
          {imageDir && (
            <span className="image-panel-dir" title={imageDir}>
              {imageDir}
            </span>
          )}
        </div>

        <button className="btn btn-small btn-outline" onClick={chooseDir} style={{ alignSelf: 'flex-start' }}>
          选择图片目录
        </button>

        <div className="image-content">
          {/* 左侧：可见的图片列表（替换原下拉框，解决「看不到图片列表」问题） */}
          <div className="image-list-side">
            <div className="image-list-title">图片列表（{images.length}）</div>
            <div className="image-list">
              {loading && <div className="image-list-empty">读取中…</div>}
              {!loading && images.length === 0 && (
                <div className="image-list-empty">
                  {imageDir ? '该目录下没有图片' : '请先选择图片目录'}
                </div>
              )}
              {images.map((img) => (
                <button
                  key={img.path}
                  className={'image-list-item' + (selected === img.path ? ' selected' : '')}
                  onClick={() => loadPreview(img.path)}
                  title={img.name}
                >
                  {img.name}
                </button>
              ))}
            </div>
          </div>

          {/* 右侧：预览 + 操作 + 识别结果 */}
          <div className="image-main-area">
            <div className="image-toolbar">
              <button
                className="btn btn-small btn-outline"
                onClick={() => setRotate((r) => (r + 90) % 360)}
                disabled={!preview}
                title="旋转 90°（仅预览，识别仍按原图）"
              >
                旋转
              </button>
              <button
                className="btn btn-small btn-outline"
                onClick={resetView}
                disabled={!preview}
                title="复位缩放与位置"
              >
                复位
              </button>
              <span className="image-zoom">{Math.round(zoom * 100)}%</span>
              {selected && (
                <span className="image-current-name" title={selected}>
                  {baseName(selected)}
                </span>
              )}
            </div>

            <div
              className="image-preview-wrap"
              ref={wrapRef}
              onMouseDown={onPreviewMouseDown}
              onDoubleClick={resetView}
              style={{ cursor: isDragging ? 'grabbing' : preview ? 'grab' : 'default' }}
              title="滚轮缩放（以光标为中心）· 左键拖拽平移 · 双击复位"
            >
              {preview ? (
                <img
                  src={preview}
                  alt="小票预览"
                  className="image-preview"
                  draggable={false}
                  style={{
                    transform: `translate(${pan.x}px, ${pan.y}px) rotate(${rotate}deg) scale(${zoom})`,
                    transition: isDragging ? 'none' : 'transform 0.2s',
                  }}
                />
              ) : (
                <div className="image-empty">选择图片后在此预览，可对照录入</div>
              )}
            </div>
            <div className="image-preview-hint">
              滚轮缩放（以光标为中心）· 左键拖拽平移 · 双击复位
            </div>

            <div className="ai-actions">
              <button
                className="btn btn-primary"
                onClick={() => doRecognize(selected)}
                disabled={aiLoading || scanning || !selected}
              >
                {aiLoading && !scanning ? '识别中…' : 'AI 识别当前图片'}
              </button>
              {scanning ? (
                <button className="btn btn-warn" onClick={stopScan} title="停止当前扫描">
                  停止扫描
                </button>
              ) : (
                <button
                  className="btn btn-outline"
                  onClick={scanAll}
                  disabled={aiLoading || loading || images.length === 0}
                  title="逐张识别全部图片，提取人名并置顶左侧对应 Excel"
                >
                  扫描全部图片（提取人名）
                </button>
              )}
            </div>

            {recognized.length > 0 && (
              <div className="person-chips">
                <span className="person-chips-label">已识别人名：</span>
                {recognized.map((p) => (
                  <span className="person-chip" key={p}>
                    {p}
                  </span>
                ))}
                <button className="btn btn-small btn-link" onClick={clearHits}>
                  清除置顶
                </button>
              </div>
            )}

            {rows.length > 0 && (
              <div className="ai-results">
                <div className="ai-results-header">
                  <span>识别结果（{rows.length} 条，可直接修改/删除后回填）</span>
                  <div className="ai-result-actions">
                    <button className="btn btn-small btn-primary" onClick={applyToMain}>
                      填入当前录入
                    </button>
                    <button className="btn btn-small btn-link" onClick={clearRows}>
                      清空
                    </button>
                  </div>
                </div>
                <div className="ai-results-table-wrap">
                  <table className="ai-results-table">
                    <thead>
                      <tr>
                        <th className="ai-col-del" title="删除该行">×</th>
                        {RESULT_COLS.map((c) => (
                          <th key={String(c.field)} style={{ minWidth: c.w }}>
                            {c.label}
                          </th>
                        ))}
                        {showSource && <th style={{ minWidth: '110px' }}>来源</th>}
                      </tr>
                    </thead>
                    <tbody>
                      {rows.map((r, i) => (
                        <tr key={i}>
                          <td className="ai-col-del">
                            <button className="ai-row-del" onClick={() => deleteRow(i)} title="删除该行">
                              ×
                            </button>
                          </td>
                          {RESULT_COLS.map((c) => (
                            <td key={String(c.field)}>
                              <input
                                className="ai-cell-input"
                                value={String(r[c.field] ?? '')}
                                onChange={(e) => updateCell(i, c.field, e.target.value)}
                              />
                            </td>
                          ))}
                          {showSource && <td className="ai-source" title={r.source}>{r.source}</td>}
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>
            )}

            {status && (
              <div className={status.includes('失败') ? 'image-status error' : 'image-status'}>
                {status.split('\n').map((line, i) => (
                  <div key={i}>{line}</div>
                ))}
              </div>
            )}
          </div>
        </div>
      </div>
    </div>
  )
}
