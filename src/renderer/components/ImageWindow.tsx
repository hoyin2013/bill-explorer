import { useEffect, useRef, useState, type MouseEvent as ReactMouseEvent } from 'react'
import { ElectronAPI, AIRecognizedRow, DetectedBox, RecognizedTicket } from '../types'

interface Props {
  api: ElectronAPI
}

// 从一批识别结果里提取去重后的客户人名。
// 注意：扁平行里 r.person 才是客户人名（由 normalizeRows 从顶层 name 字段填入），
// r.name 是商品品名（goods），误用它会导致左侧按人名置顶/命中的功能失效。
function uniquePersons(rows: AIRecognizedRow[]): string[] {
  const set = new Set<string>()
  for (const r of rows) {
    const n = (r.person || '').trim()
    if (n) set.add(n)
  }
  return Array.from(set)
}

// 取图片文件名（去掉路径）作为"来源"标记
function baseName(path: string): string {
  const idx = Math.max(path.lastIndexOf('/'), path.lastIndexOf('\\'))
  return idx >= 0 ? path.slice(idx + 1) : path
}

// 把 data URL 还原成纯 base64（去掉 data:...;base64, 前缀）
function rawB64(dataUrl: string): string {
  const idx = dataUrl.indexOf(',')
  return idx >= 0 ? dataUrl.slice(idx + 1) : dataUrl
}

// 用 canvas 把图片按 angleDeg（顺时针，90 的整数倍）旋转，返回新的 data URL。
// 用于在送检测前把整图手动旋转到正向，使检测框坐标与旋转后的展示图对齐。
function rotateBase64Image(dataUrl: string, angleDeg: number): Promise<string> {
  return new Promise((resolve, reject) => {
    if (angleDeg % 360 === 0) {
      resolve(dataUrl)
      return
    }
    const img = new Image()
    img.onload = () => {
      const rad = (angleDeg * Math.PI) / 180
      const sin = Math.abs(Math.sin(rad))
      const cos = Math.abs(Math.cos(rad))
      const w = img.width
      const h = img.height
      const nw = Math.round(w * cos + h * sin)
      const nh = Math.round(w * sin + h * cos)
      const canvas = document.createElement('canvas')
      canvas.width = nw
      canvas.height = nh
      const ctx = canvas.getContext('2d')
      if (!ctx) {
        reject(new Error('无法创建画布上下文'))
        return
      }
      ctx.translate(nw / 2, nh / 2)
      ctx.rotate(rad)
      ctx.drawImage(img, -w / 2, -h / 2)
      resolve(canvas.toDataURL('image/jpeg', 0.92))
    }
    img.onerror = () => reject(new Error('图片加载失败'))
    img.src = dataUrl
  })
}

// 把识别行转成 TSV 文本，便于直接粘贴进 Excel（去掉不展示的调货人/备注列）
function rowsToTsv(rs: AIRecognizedRow[]): string {
  const header = ['日期', '货品', '单位', '数量', '单价', '金额']
  const lines = [header.join('\t')]
  for (const r of rs) {
    lines.push(
      [r.date ?? '', r.name ?? '', r.unit ?? '', r.qty ?? '', r.price ?? '', r.amount ?? ''].join('\t'),
    )
  }
  return lines.join('\n')
}

// 结果表列定义（已去掉调货人 / 备注：识别不准且无需记录）
const RESULT_COLS: Array<{ field: keyof AIRecognizedRow; label: string; w: string }> = [
  { field: 'no', label: '序号', w: '42px' },
  { field: 'date', label: '日期', w: '92px' },
  { field: 'name', label: '货品', w: 'auto' },
  { field: 'unit', label: '单位', w: '54px' },
  { field: 'qty', label: '数量', w: '54px' },
  { field: 'price', label: '单价', w: '70px' },
  { field: 'amount', label: '金额', w: '82px' },
]

export function ImageWindow({ api }: Props) {
  const [imageDir, setImageDir] = useState('')
  const [images, setImages] = useState<Array<{ name: string; path: string }>>([])
  const [selected, setSelected] = useState('')
  const [preview, setPreview] = useState('')
  const [rotate, setRotate] = useState(0)
  const [loading, setLoading] = useState(false)
  const [aiLoading, setAiLoading] = useState(false)
  const [status, setStatus] = useState('')
  // 完整识别结果（非检测流程：整图识别 / 扫描全部），既用于展示，也用于"填入当前录入"回填
  const [rows, setRows] = useState<AIRecognizedRow[]>([])
  const [recognized, setRecognized] = useState<string[]>([])
  // YOLOv8 检测出的小票边界框（预览缩放图像素坐标）与对应原图尺寸，用于画框叠加
  const [boxes, setBoxes] = useState<DetectedBox[]>([])
  const [imgNatural, setImgNatural] = useState<{ w: number; h: number }>({ w: 0, h: 0 })
  // 逐张拆分出的小票（含裁剪图 + 各自识别结果），用于逐个放大查看 / 录入
  const [tickets, setTickets] = useState<RecognizedTicket[]>([])
  // 视图模式：总览（带检测框的原图）/ 单张（放大某张小票）
  const [viewMode, setViewMode] = useState<'overview' | 'single'>('overview')
  const [active, setActive] = useState(0) // 当前查看的 ticket 下标
  const [singleRotate, setSingleRotate] = useState(0) // 单张视图内手动微调旋转
  const [singleLoading, setSingleLoading] = useState(false) // 逐张识别中
  const [copyMsg, setCopyMsg] = useState('')
  // 检测增强环境探测结果（ONNX 模型与运行时是否就绪）；null=尚未探测完成
  const [detectEnv, setDetectEnv] = useState<{
    modelExists: boolean
    runtimeReady: boolean
    detail: string
  } | null>(null)
  // 预览缩放 / 平移状态
  const [zoom, setZoom] = useState(1)
  const [pan, setPan] = useState<{ x: number; y: number }>({ x: 0, y: 0 })
  const [isDragging, setIsDragging] = useState(false)
  // 图片列表（左）与预览区（右）的可调分隔宽度（px）
  const [listWidth, setListWidth] = useState(200)
  // 复位旋转瞬间禁用过渡，避免「从旋转角回正」触发可见的旋转动画
  const [instant, setInstant] = useState(false)
  const wrapRef = useRef<HTMLDivElement>(null)
  const zoomRef = useRef(1)
  const panRef = useRef({ x: 0, y: 0 })
  const draggingRef = useRef<{ startX: number; startY: number; panX: number; panY: number } | null>(null)
  // 日期统一基准：首次编辑某张日期后记住，并同步到当时所有小票的每一行；
  // 之后跨图片保留（不在 loadPreview / detectOnly 重置），后续识别默认套用该日期，
  // 避免每次靠 OCR 识别日期带来的误差。
  const dateAnchorRef = useRef<string | null>(null)

  // 是否走"检测拆分"流程（用 tickets 作为数据源）；否则用 rows
  const usingTickets = tickets.length > 0
  const displayRows = usingTickets ? tickets.flatMap((t) => t.rows) : rows

  // 检测增强环境是否不可用：探测完成后，运行时或模型任意缺失即不可用
  const detectNoRuntime = detectEnv ? !detectEnv.runtimeReady : false
  const detectUnavailable = detectEnv ? !detectEnv.runtimeReady || !detectEnv.modelExists : false
  const detectHint = detectEnv && detectEnv.detail ? detectEnv.detail : ''

  // 把当前已识别的人名实时上报给主窗口（主窗口据此把对应 Excel 置顶）
  useEffect(() => {
    api.reportPersons(recognized)
  }, [recognized, api])

  // 根据 displayRows 自动更新已识别人名。
  // 单张视图下，只上报「当前正在查看的那张」小票的人名 —— 逐张识别时左侧「命中」置顶
  // 才不会不断叠加；总览视图仍用全部小票的并集（便于「填入全部」批量核对）。
  useEffect(() => {
    if (viewMode === 'single') {
      const activeRows = tickets[active]?.rows ?? []
      setRecognized(activeRows.length ? uniquePersons(activeRows) : [])
    } else {
      setRecognized(uniquePersons(displayRows))
    }
  }, [tickets, rows, viewMode, active])

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
    setTickets([])
    setBoxes([])
    setViewMode('overview')
    setActive(0)
    setSingleRotate(0)
    setImgNatural({ w: 0, h: 0 })
    // 注意：不再重置 dateAnchorRef —— 第一张输入的日期要作为整场会话基准，
    // 跨图片保留，后续识别默认套用该日期（避免每次识别误差）。
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

  // 拖动分隔条调整「图片列表 / 预览」左右宽度
  function onSplitterDown(e: ReactMouseEvent) {
    e.preventDefault()
    const startX = e.clientX
    const startW = listWidth
    const onMove = (ev: MouseEvent) => {
      const nw = Math.min(460, Math.max(140, startW + (ev.clientX - startX)))
      setListWidth(nw)
    }
    const onUp = () => {
      window.removeEventListener('mousemove', onMove)
      window.removeEventListener('mouseup', onUp)
    }
    window.addEventListener('mousemove', onMove)
    window.addEventListener('mouseup', onUp)
  }

  // 行内编辑：更新非检测流程的某一行某一字段
  function updateCell(index: number, field: keyof AIRecognizedRow, value: string) {
    setRows((prev) => {
      const next = prev.slice()
      const row = { ...next[index] }
      ;(row as Record<string, string>)[field as string] = value
      next[index] = row
      return next
    })
  }

  // 删除非检测流程的某一行
  function deleteRow(index: number) {
    setRows((prev) => prev.filter((_, i) => i !== index))
  }

  // 编辑检测流程中某张小票的某一行
  function updateTicketCell(ti: number, ri: number, field: keyof AIRecognizedRow, value: string) {
    setTickets((prev) => {
      const next = prev.slice()
      const t = { ...next[ti], rows: next[ti].rows.slice() }
      const row = { ...t.rows[ri] }
      ;(row as Record<string, string>)[field as string] = value
      t.rows[ri] = row
      next[ti] = t
      // 日期统一：识别出的日期很不准，通常一图同日。用户首次编辑某张日期（非空）后，
      // 记住该日期并把它同步到所有小票的每一行；之后再单独改某张日期，只改当前单元格。
      if (field === 'date' && value.trim()) {
        if (dateAnchorRef.current == null) {
          dateAnchorRef.current = value
          for (const nt of next) {
            nt.rows = nt.rows.map((rr) => ({ ...rr, date: value }))
          }
        }
      }
      return next
    })
  }

  // 删除检测流程中某张小票的某一行
  function deleteTicketRow(ti: number, ri: number) {
    setTickets((prev) => {
      const next = prev.slice()
      const t = { ...next[ti], rows: next[ti].rows.filter((_, i) => i !== ri) }
      next[ti] = t
      return next
    })
  }

  // 把当前预览按 rotate 旋转，返回旋转后的纯 base64，并顺手把展示图替换为旋转版本
  async function rotatePreviewForDetect(): Promise<string> {
    const rotated = await rotateBase64Image(preview, rotate)
    setPreview(rotated)
    return rawB64(rotated)
  }

  // 仅检测小票边界框：调用 YOLOv8 标出每张小票矩形框（并取回裁剪图），用于点击放大查看+识别
  async function detectOnly(path: string) {
    if (!path) {
      setStatus('请先选择一张图片')
      return
    }
    if (!preview) {
      setStatus('图片尚未加载完成')
      return
    }
    setAiLoading(true)
    setStatus('检测小票边界框中…')
    setBoxes([])
    setTickets([])
    setRows([])
    // 注意：保留 dateAnchorRef（不重置），让日期基准跨图片延续
    setViewMode('overview')
    try {
      // 按当前旋转把预览图旋转后送检测，使检测图与展示图一致（框坐标对齐）
      const b64 = rotate % 360 === 0 ? rawB64(preview) : await rotatePreviewForDetect()
      // 展示图已替换为旋转后的版本，取消 CSS 旋转并复位缩放/平移，保证框对齐。
      // 瞬间禁用过渡，避免「从旋转角回正」触发一段可见的旋转动画（画面闪一下）。
      setInstant(true)
      setRotate(0)
      resetView()
      setTimeout(() => setInstant(false), 120)
      const res = await api.aiDetect({ imagePath: path, imageBase64: b64 })
      if (res.modelAvailable === false) {
        setStatus(res.message || '检测模型不可用，请检查设置中的 Python 解释器与模型路径')
      } else if (res.tickets) {
        setBoxes(res.boxes || [])
        setTickets(res.tickets)
        if (res.imageWidth) setImgNatural({ w: res.imageWidth, h: res.imageHeight || 0 })
        setStatus(`检测到 ${res.tickets.length} 张小票，点击下方检测框或下方缩略图放大，再点「识别此张」读取文字`)
      } else {
        setStatus(res.message || '未检测到小票')
      }
    } catch (err) {
      setStatus('检测失败：' + (err instanceof Error ? err.message : '未知错误'))
    } finally {
      setAiLoading(false)
    }
  }

  // 打开单张小票视图（仅放大，不自动识别；识别由用户点「识别此张」触发）
  function openTicket(i: number) {
    if (i < 0 || i >= tickets.length) return
    setActive(i)
    setSingleRotate(0)
    setViewMode('single')
  }

  // 逐张识别：对检测出来但还没识别的小票，调用 AI 识别这一张裁剪图
  async function recognizeThisCrop(i: number) {
    const t = tickets[i]
    if (!t || !t.crop) {
      setStatus('该小票没有可用的裁剪图')
      return
    }
    setSingleLoading(true)
    setStatus(`正在识别第 ${i + 1} 张小票…`)
    try {
      const res = await api.aiRecognizeCrop(t.crop)
      if (res.error) {
        setStatus(res.message || '识别失败')
      } else {
        const src = baseName(selected)
        // 备注(remark)按需求留空，不追加「#序号」（小票序号由 tickets[].index 承载）
        // 已设过统一日期基准时，把新识别出的小票日期也套用基准，保持一致
        const anchor = dateAnchorRef.current
        const rs = (res.rows || []).map((r) => ({ ...r, source: src, date: anchor ?? r.date }))
        setTickets((prev) => {
          const next = prev.slice()
          next[i] = { ...next[i], rows: rs }
          return next
        })
        setStatus(`第 ${i + 1} 张识别完成，共 ${rs.length} 条记录` + (res.message ? '\n' + res.message : ''))
      }
    } catch (err) {
      setStatus('识别失败：' + (err instanceof Error ? err.message : '未知错误'))
    } finally {
      setSingleLoading(false)
    }
  }

  // 复制文本到剪贴板
  async function copyText(text: string, msg: string) {
    try {
      await navigator.clipboard.writeText(text)
      setCopyMsg(msg)
    } catch {
      setCopyMsg('复制失败，请手动选择文本')
    }
    setTimeout(() => setCopyMsg(''), 2200)
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
    setTickets([])
    setStatus('已清空识别结果')
  }

  function applyRows(target: AIRecognizedRow[]) {
    if (!target.length) {
      setStatus('没有可回填的识别结果')
      return
    }
    api.applyToMain(target)
    setStatus(
      `已把 ${target.length} 条识别结果发送到「账单录入器」主窗口。\n请在主窗口左侧打开目标 Excel，点击网格即可核对/修改后保存（不会影响当前录入，除非你主动编辑）。`,
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

  // 挂载时探测检测增强环境，提前告知用户是否可用（而非点「检测」后才报错）
  useEffect(() => {
    api
      .detectEnvironment()
      .then(setDetectEnv)
      .catch(() => setDetectEnv(null))
  }, [api])

  // 窗口即将关闭（pagehide，关闭前必触发）时通知主进程：
  // 清空主窗口左侧搜索框 + 清除人名命中置顶。与 reportPersons 同一条已被验证可用的 IPC 路径，
  // 比单纯依赖主进程 imageWindow.on('closed') 更可靠，确保「关闭录入窗口即清空搜索框」稳定生效。
  useEffect(() => {
    const onHide = () => {
      try {
        api.notifyImageClosing()
      } catch {
        /* 渲染进程即将卸载，忽略 */
      }
    }
    window.addEventListener('pagehide', onHide)
    return () => window.removeEventListener('pagehide', onHide)
  }, [api])

  // 是否展示"来源"列：仅当结果来自多张不同图片时（单图识别不显示，避免冗余）
  const distinctSources = new Set(rows.map((r) => r.source).filter(Boolean))
  const showSource = distinctSources.size > 1

  const activeTicket = tickets[active]
  const isLast = active >= tickets.length - 1
  const isFirst = active <= 0

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
          {/* 左侧：可见的图片列表 */}
          <div className="image-list-side" style={{ width: listWidth, flex: '0 0 auto' }}>
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
          <div
            className="image-splitter"
            onMouseDown={onSplitterDown}
            title="拖动调整左侧列表与右侧预览的宽度"
          />
          <div className="image-main-area">
            {viewMode === 'single' && activeTicket ? (
              /* ====== 单张小票放大视图 ====== */
              <div className="single-ticket-view">
                <div className="image-toolbar">
                  <button className="btn btn-small btn-outline" onClick={() => setViewMode('overview')} title="返回带检测框的总览">
                    ← 返回总览
                  </button>
                  <span className="single-counter">
                    第 <b>{active + 1}</b> / {tickets.length} 张
                  </span>
                  {activeTicket.box.conf != null && (
                    <span className="single-conf">置信度 {Math.round(activeTicket.box.conf * 100)}%</span>
                  )}
                  {activeTicket.angle ? <span className="single-rot">已自动旋转 {activeTicket.angle}°</span> : null}
                </div>

                <div className="single-image-wrap">
                  <img
                    src={`data:image/jpeg;base64,${activeTicket.crop}`}
                    alt={`第 ${active + 1} 张小票`}
                    className="single-image"
                    draggable={false}
                    style={{ transform: `rotate(${singleRotate}deg)` }}
                  />
                </div>
                <div className="single-rotate-bar">
                  <button className="btn btn-small btn-outline" onClick={() => setSingleRotate((r) => r - 90)}>↺ 左转90°</button>
                  <button className="btn btn-small btn-outline" onClick={() => setSingleRotate((r) => r + 90)}>↻ 右转90°</button>
                  <button className="btn btn-small btn-link" onClick={() => setSingleRotate(0)}>复位</button>
                </div>

                {/* 该张小票的识别信息（可编辑，便于核对后录入） */}
                <div className="single-info">
                  <div className="single-info-header">
                    <span>本张识别信息（{activeTicket.rows.length} 条，可直接修改）</span>
                    <div className="ai-result-actions">
                      <button
                        className="btn btn-small btn-primary"
                        disabled={!activeTicket.rows.length}
                        onClick={() => copyText(rowsToTsv(activeTicket.rows), `已复制第 ${active + 1} 张的 ${activeTicket.rows.length} 条记录`)}
                      >
                        复制本张
                      </button>
                      <button
                        className="btn btn-small btn-primary"
                        disabled={!activeTicket.rows.length}
                        onClick={() => applyRows(activeTicket.rows)}
                      >
                        填入本张
                      </button>
                      {activeTicket.rows.length === 0 ? (
                        <button
                          className="btn btn-small btn-outline"
                          onClick={() => recognizeThisCrop(active)}
                          disabled={singleLoading || !activeTicket.crop}
                        >
                          {singleLoading ? '识别中…' : '识别此张'}
                        </button>
                      ) : (
                        <button
                          className="btn btn-small btn-outline"
                          onClick={() => recognizeThisCrop(active)}
                          disabled={singleLoading || !activeTicket.crop}
                          title="重新调用模型 API 识别本张小票"
                        >
                          重新识别
                        </button>
                      )}
                    </div>
                  </div>

                  {activeTicket.rows.length > 0 ? (
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
                          </tr>
                        </thead>
                        <tbody>
                          {activeTicket.rows.map((r, ri) => (
                            <tr key={ri}>
                              <td className="ai-col-del">
                                <button className="ai-row-del" onClick={() => deleteTicketRow(active, ri)} title="删除该行">×</button>
                              </td>
                              {RESULT_COLS.map((c) => (
                                <td key={String(c.field)}>
                                  <input
                                    className="ai-cell-input"
                                    value={String(r[c.field] ?? '')}
                                    onChange={(e) => updateTicketCell(active, ri, c.field, e.target.value)}
                                  />
                                  {c.field === 'person' && r.personCorrected && (
                                    <span className="person-fixed" title="已按人名清单自动修正">已修正</span>
                                  )}
                                </td>
                              ))}
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    </div>
                  ) : (
                    <div className="single-empty">
                      {singleLoading ? '识别中…' : '本张小票尚未识别，点击右上角「识别此张」进行 AI 识别。'}
                    </div>
                  )}
                </div>

                {/* 上一张 / 下一张 导航（下一张自动跳转） */}
                <div className="single-nav">
                  <button
                    className="btn btn-outline"
                    onClick={() => openTicket(active - 1)}
                    disabled={isFirst || aiLoading}
                  >
                    ← 上一张
                  </button>
                  <button
                    className="btn btn-primary"
                    onClick={() => openTicket(active + 1)}
                    disabled={isLast || aiLoading}
                  >
                    下一张 →
                  </button>
                </div>
              </div>
            ) : (
              /* ====== 总览视图（带检测框的原图） ====== */
              <>
                <div className="image-toolbar">
                  <button
                    className="btn btn-small btn-outline"
                    onClick={() => setRotate((r) => (r + 270) % 360)}
                    disabled={!preview}
                    title="向左旋转 90°"
                  >
                    ↺ 左转
                  </button>
                  <button
                    className="btn btn-small btn-outline"
                    onClick={() => setRotate((r) => (r + 90) % 360)}
                    disabled={!preview}
                    title="向右旋转 90°"
                  >
                    ↻ 右转
                  </button>
                  <button
                    className="btn btn-small btn-link"
                    onClick={() => { setRotate(0); resetView() }}
                    disabled={!preview}
                    title="复位旋转与缩放/位置"
                  >
                    复位
                  </button>
                  <span className="image-rotate-badge">旋转 {rotate}°</span>
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
                        transition: instant || isDragging ? 'none' : 'transform 0.2s',
                      }}
                    />
                  ) : (
                    <div className="image-empty">选择图片后在此预览，可对照录入</div>
                  )}
                  {/* YOLOv8 检测框叠加层：矩形可点击 → 放大该张小票 */}
                  {preview && boxes.length > 0 && imgNatural.w > 0 && (
                    <div
                      className="ticket-box-overlay"
                      style={{
                        transform: `translate(${pan.x}px, ${pan.y}px) rotate(${rotate}deg) scale(${zoom})`,
                        transformOrigin: 'center center',
                        transition: instant || isDragging ? 'none' : 'transform 0.2s',
                        pointerEvents: 'none',
                      }}
                    >
                      <svg
                        width="100%"
                        height="100%"
                        viewBox={`0 0 ${imgNatural.w} ${imgNatural.h}`}
                        preserveAspectRatio="xMidYMid meet"
                      >
                        {boxes.map((b, i) => {
                          const sw = Math.max(2, imgNatural.w * 0.004)
                          const fs = Math.max(12, imgNatural.w * 0.022)
                          return (
                            <g key={i} className="ticket-box-rect" onClickCapture={() => openTicket(i)}>
                              <rect
                                x={b.x}
                                y={b.y}
                                width={b.w}
                                height={b.h}
                                fill="rgba(64,158,255,0.12)"
                                stroke="#409eff"
                                strokeWidth={sw}
                                style={{ pointerEvents: 'auto', cursor: 'pointer' }}
                              />
                              <text
                                x={b.x}
                                y={Math.max(b.y - fs * 0.4, fs)}
                                fill="#fff"
                                stroke="#409eff"
                                strokeWidth={sw * 0.6}
                                fontSize={fs}
                                fontWeight="bold"
                                style={{ pointerEvents: 'auto', cursor: 'pointer' }}
                              >
                                {`#${i + 1} ${Math.round(b.conf * 100)}%`}
                              </text>
                            </g>
                          )
                        })}
                      </svg>
                    </div>
                  )}
                </div>
                <div className="image-preview-hint">
                  滚轮缩放（以光标为中心）· 左键拖拽平移 · 双击复位
                </div>

                <div className="ai-actions">
                  <button
                    className="btn btn-primary"
                    onClick={() => detectOnly(selected)}
                    disabled={aiLoading || !selected || detectNoRuntime}
                    title="先旋转到正向，再点此：用训练好的 YOLOv8 模型（ONNX 运行时）把每张小票框出来并裁剪。点框可放大，再点「识别此张」读文字"
                  >
                    {aiLoading ? '检测中…' : '检测小票（画框）'}
                  </button>
                </div>
                {detectUnavailable && detectHint && (
                  <div className="detect-hint">
                    <span className="detect-hint-icon">⚠</span>
                    <span>{detectHint}</span>
                  </div>
                )}

                {boxes.length > 0 && (
                  <div className="detect-badge">
                    已标注 {boxes.length} 张小票边界框
                    <button className="btn btn-small btn-link" onClick={() => { setBoxes([]); setTickets([]) }}>
                      清除框
                    </button>
                  </div>
                )}

                {/* 逐张小票缩略图条：点击放大查看单张 */}
                {tickets.length > 0 && (
                  <div className="ticket-strip">
                    <div className="ticket-strip-title">逐张小票（点击放大 / 录入）</div>
                    <div className="ticket-strip-list">
                      {tickets.map((t, i) => (
                        <button
                          key={i}
                          className={'ticket-thumb' + (viewMode === 'single' && active === i ? ' active' : '') + (t.rows.length ? '' : ' unrec')}
                          onClick={() => openTicket(i)}
                          title={`第 ${i + 1} 张${t.rows.length ? '' : '（未识别）'}`}
                        >
                          {t.crop ? (
                            <img src={`data:image/jpeg;base64,${t.crop}`} alt={`第 ${i + 1} 张`} draggable={false} />
                          ) : (
                            <span className="ticket-thumb-no">无图</span>
                          )}
                          <span className="ticket-thumb-idx">#{i + 1}</span>
                          {t.rows.length === 0 && <span className="ticket-thumb-flag">待识别</span>}
                        </button>
                      ))}
                    </div>
                  </div>
                )}

                {/* 总览结果区：检测流程显示汇总，非检测流程显示可编辑明细表 */}
                {usingTickets ? (
                  <div className="ai-results">
                    <div className="ai-results-header">
                      <span>
                        共拆分 {tickets.length} 张小票，识别出 {displayRows.length} 条记录。
                        {recognized.length > 0 && ` 已识别人名：${recognized.join('、')}`}
                      </span>
                      <div className="ai-result-actions">
                        <button
                          className="btn btn-small btn-primary"
                          disabled={!displayRows.length}
                          onClick={() => applyRows(displayRows)}
                        >
                          填入全部
                        </button>
                        <button className="btn btn-small btn-link" onClick={clearRows}>清空</button>
                      </div>
                    </div>
                    <div className="ai-results-hint">点击上方检测框或下方缩略图，可逐张放大查看并复制/录入。</div>
                  </div>
                ) : (
                  rows.length > 0 && (
                    <div className="ai-results">
                      <div className="ai-results-header">
                        <span>识别结果（{rows.length} 条，可直接修改/删除后回填）</span>
                        <div className="ai-result-actions">
                          <button className="btn btn-small btn-primary" onClick={() => applyRows(rows)}>填入当前录入</button>
                          <button className="btn btn-small btn-link" onClick={clearRows}>清空</button>
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
                                  <button className="ai-row-del" onClick={() => deleteRow(i)} title="删除该行">×</button>
                                </td>
                                {RESULT_COLS.map((c) => (
                                  <td key={String(c.field)}>
                                    <input
                                      className="ai-cell-input"
                                      value={String(r[c.field] ?? '')}
                                      onChange={(e) => updateCell(i, c.field, e.target.value)}
                                    />
                                    {c.field === 'person' && r.personCorrected && (
                                      <span className="person-fixed" title="已按人名清单自动修正">已修正</span>
                                    )}
                                  </td>
                                ))}
                                {showSource && <td className="ai-source" title={r.source}>{r.source}</td>}
                              </tr>
                            ))}
                          </tbody>
                        </table>
                      </div>
                    </div>
                  )
                )}

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
              </>
            )}

            {status && (
              <div className={status.includes('失败') ? 'image-status error' : 'image-status'}>
                {status.split('\n').map((line, i) => (
                  <div key={i}>{line}</div>
                ))}
              </div>
            )}
            {copyMsg && <div className="image-status ok">{copyMsg}</div>}
          </div>
        </div>
      </div>
    </div>
  )
}
