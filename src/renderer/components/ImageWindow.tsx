import { forwardRef, useEffect, useImperativeHandle, useLayoutEffect, useRef, useState, type MouseEvent as ReactMouseEvent, type MutableRefObject } from 'react'
import { ElectronAPI, AIRecognizedRow, DetectedBox, RecognizedTicket, ImageSnapshot } from '../types'

interface Props {
  api: ElectronAPI
  detached?: boolean
  onDetach?: (state: ImageSnapshot) => void
  onAttach?: (state: ImageSnapshot) => void
  // 独立窗口启动时由主进程回传的拆分前状态快照（保留结果与进度）
  initialState?: ImageSnapshot
}

// 暴露给父组件的命令式句柄：用于拆分/合并窗口时抓取与恢复整块状态
export interface ImageWindowHandle {
  captureState: () => ImageSnapshot
  restoreState: (s: ImageSnapshot) => void
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

export const ImageWindow = forwardRef<ImageWindowHandle, Props>(function ImageWindow(
  { api, detached, onDetach, onAttach, initialState },
  ref,
) {
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
  const [copyMsg, setCopyMsg] = useState('')
  // 逐张识别的队列与状态（后台顺序识别，不阻塞前台操作）：
  // recogState[i] 标记每张小票的识别状态；bgRunning 表示后台队列正在跑。
  const [recogState, setRecogState] = useState<Record<number, 'pending' | 'busy' | 'done' | 'error'>>({})
  const [bgRunning, setBgRunning] = useState(false)
  // tickets 的实时镜像（供异步队列读取最新裁剪图 / 已识别结果，避免闭包读到旧值）
  const ticketsRef = useRef<RecognizedTicket[]>([])
  // 后台识别队列管理器：queue=待处理下标，running=是否有循环在跑，cancelled=取消
  const recogMgr = useRef<{ queue: number[]; running: boolean; cancelled: boolean }>({
    queue: [],
    running: false,
    cancelled: false,
  })
  // 单张小票图片点击放大：全屏查看（lightbox）
  const [lightbox, setLightbox] = useState(false)
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
  // 单张小票内联视图缩放 / 平移状态（与总览预览共用同一套交互）
  const [singleZoom, setSingleZoom] = useState(1)
  const [singlePan, setSinglePan] = useState<{ x: number; y: number }>({ x: 0, y: 0 })
  const [singleDragging, setSingleDragging] = useState(false)
  const singleZoomRef = useRef(1)
  const singlePanRef = useRef({ x: 0, y: 0 })
  const singleDragRef = useRef<{ startX: number; startY: number; panX: number; panY: number; moved: boolean } | null>(null)
  const singleClickGuard = useRef(false)
  const singleWrapRef = useRef<HTMLDivElement>(null)
  // 放大查看（lightbox）缩放 / 平移状态
  const [lbZoom, setLbZoom] = useState(1)
  const [lbPan, setLbPan] = useState<{ x: number; y: number }>({ x: 0, y: 0 })
  const [lbDragging, setLbDragging] = useState(false)
  const lbZoomRef = useRef(1)
  const lbPanRef = useRef({ x: 0, y: 0 })
  const lbDragRef = useRef<{ startX: number; startY: number; panX: number; panY: number } | null>(null)
  const lbWrapRef = useRef<HTMLDivElement>(null)
  // 图片列表（左）与预览区（右）的可调分隔宽度（px）；面板整体偏窄，默认收窄
  const [listWidth, setListWidth] = useState(104)
  // 左侧图片列表是否折叠：折叠后识别信息区占满整个窗口宽度，便于查看/录入
  const [listCollapsed, setListCollapsed] = useState(false)
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

  // 抓取当前整块状态（用于拆分窗口时把已识别结果与进度带过去）
  function captureState(): ImageSnapshot {
    return {
      imageDir,
      images,
      selected,
      preview,
      rotate,
      rows,
      boxes,
      imgNatural,
      tickets,
      viewMode,
      active,
      singleRotate,
      recogState,
      zoom,
      pan,
      singleZoom,
      singlePan,
      listWidth,
      listCollapsed,
      dateAnchor: dateAnchorRef.current,
    }
  }

  // 从快照恢复整块状态（独立窗口启动时 / 合并回主窗口时复用，保留结果与进度）
  function restoreState(s: ImageSnapshot) {
    setImageDir(s.imageDir ?? '')
    setImages(s.images ?? [])
    setSelected(s.selected ?? '')
    setPreview(s.preview ?? '')
    setRotate(s.rotate ?? 0)
    setRows(s.rows ?? [])
    setBoxes(s.boxes ?? [])
    setImgNatural(s.imgNatural ?? { w: 0, h: 0 })
    setTickets(s.tickets ?? [])
    setViewMode(s.viewMode ?? 'overview')
    setActive(s.active ?? 0)
    setSingleRotate(s.singleRotate ?? 0)
    setRecogState(s.recogState ?? {})
    setZoom(s.zoom ?? 1)
    setPan(s.pan ?? { x: 0, y: 0 })
    setSingleZoom(s.singleZoom ?? 1)
    setSinglePan(s.singlePan ?? { x: 0, y: 0 })
    setListWidth(s.listWidth ?? 124)
    setListCollapsed(s.listCollapsed ?? false)
    dateAnchorRef.current = s.dateAnchor ?? null
    // 复位后台识别队列与瞬时过渡，避免残留动画 / 重复识别
    recogMgr.current.cancelled = true
    recogMgr.current.running = false
    recogMgr.current.queue = []
    setBgRunning(false)
    setInstant(true)
  }

  useImperativeHandle(ref, () => ({ captureState, restoreState }), [captureState, restoreState])

  // 独立窗口：启动时用主进程回传的快照恢复结果与进度。
  // 用 useLayoutEffect 在首帧绘制前完成恢复，避免「先空白再填充」的闪烁，
  // 保证独立窗口打开即是主面板当时的内容（检测框、当前第几张、已识别结果都一致）。
  useLayoutEffect(() => {
    if (initialState) restoreState(initialState)
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [initialState])

  // 独立窗口：X 关闭 / 页面卸载前，把最新状态回传主进程，确保合并回主窗口的是最新进度
  const captureRef = useRef(captureState)
  captureRef.current = captureState
  useEffect(() => {
    if (!detached) return
    const onBeforeUnload = () => {
      try {
        api.detachedStateUpdate(captureRef.current())
      } catch {
        /* 忽略关闭瞬间的回传失败 */
      }
    }
    window.addEventListener('beforeunload', onBeforeUnload)
    return () => window.removeEventListener('beforeunload', onBeforeUnload)
  }, [detached, api])

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

  // 放大查看（lightbox）打开时，按 Esc 关闭
  useEffect(() => {
    if (!lightbox) return
    const onKey = (e: KeyboardEvent) => {
      if (e.key === 'Escape') setLightbox(false)
    }
    window.addEventListener('keydown', onKey)
    return () => window.removeEventListener('keydown', onKey)
  }, [lightbox])

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

  // 同步 tickets 最新值到 ref，供后台识别队列读取（避免异步闭包读到过期快照）
  useEffect(() => {
    ticketsRef.current = tickets
  }, [tickets])

  // 组件卸载（如独立窗口关闭）时取消后台识别队列，避免对已卸载组件 setState
  useEffect(() => {
    return () => {
      recogMgr.current.cancelled = true
    }
  }, [])

  // 同步最新的 zoom/pan 到 ref（供原生事件处理器读取）
  useEffect(() => {
    zoomRef.current = zoom
  }, [zoom])
  useEffect(() => {
    panRef.current = pan
  }, [pan])
  useEffect(() => {
    singleZoomRef.current = singleZoom
  }, [singleZoom])
  useEffect(() => {
    singlePanRef.current = singlePan
  }, [singlePan])
  useEffect(() => {
    lbZoomRef.current = lbZoom
  }, [lbZoom])
  useEffect(() => {
    lbPanRef.current = lbPan
  }, [lbPan])
  // 放大查看（lightbox）每次打开都复位缩放/平移，从原始大小开始
  useEffect(() => {
    if (lightbox) {
      lbZoomRef.current = 1
      lbPanRef.current = { x: 0, y: 0 }
      setLbZoom(1)
      setLbPan({ x: 0, y: 0 })
    }
  }, [lightbox])

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

  // 单张视图：滚轮缩放（以光标为锚点）
  useEffect(() => {
    const el = singleWrapRef.current
    if (!el || viewMode !== 'single') return
    const onWheel = (e: WheelEvent) => {
      if (!tickets[active]) return
      e.preventDefault()
      const rect = el.getBoundingClientRect()
      const cx = e.clientX - (rect.left + rect.width / 2)
      const cy = e.clientY - (rect.top + rect.height / 2)
      zoomAround(singleZoomRef, singlePanRef, setSingleZoom, setSinglePan, e.deltaY < 0 ? 1.12 : 1 / 1.12, cx, cy)
    }
    el.addEventListener('wheel', onWheel, { passive: false })
    return () => el.removeEventListener('wheel', onWheel)
  }, [viewMode, active, tickets])

  // 放大查看（lightbox）：滚轮缩放（以光标为锚点）
  useEffect(() => {
    const el = lbWrapRef.current
    if (!el || !lightbox) return
    const onWheel = (e: WheelEvent) => {
      if (!tickets[active]) return
      e.preventDefault()
      const rect = el.getBoundingClientRect()
      const cx = e.clientX - (rect.left + rect.width / 2)
      const cy = e.clientY - (rect.top + rect.height / 2)
      zoomAround(lbZoomRef, lbPanRef, setLbZoom, setLbPan, e.deltaY < 0 ? 1.12 : 1 / 1.12, cx, cy)
    }
    el.addEventListener('wheel', onWheel, { passive: false })
    return () => el.removeEventListener('wheel', onWheel)
  }, [lightbox, tickets])

  // 单张视图：拖拽平移（监听 window，拖出容器也能继续）
  useEffect(() => {
    const onMove = (e: MouseEvent) => {
      const d = singleDragRef.current
      if (!d) return
      const nx = d.panX + (e.clientX - d.startX)
      const ny = d.panY + (e.clientY - d.startY)
      if (Math.abs(e.clientX - d.startX) > 3 || Math.abs(e.clientY - d.startY) > 3) d.moved = true
      singlePanRef.current = { x: nx, y: ny }
      setSinglePan({ x: nx, y: ny })
    }
    const onUp = () => {
      if (singleDragRef.current) {
        if (singleDragRef.current.moved) singleClickGuard.current = true
        singleDragRef.current = null
        setSingleDragging(false)
      }
    }
    window.addEventListener('mousemove', onMove)
    window.addEventListener('mouseup', onUp)
    return () => {
      window.removeEventListener('mousemove', onMove)
      window.removeEventListener('mouseup', onUp)
    }
  }, [])

  // 放大查看（lightbox）：拖拽平移（监听 window）
  useEffect(() => {
    const onMove = (e: MouseEvent) => {
      const d = lbDragRef.current
      if (!d) return
      const nx = d.panX + (e.clientX - d.startX)
      const ny = d.panY + (e.clientY - d.startY)
      lbPanRef.current = { x: nx, y: ny }
      setLbPan({ x: nx, y: ny })
    }
    const onUp = () => {
      if (lbDragRef.current) {
        lbDragRef.current = null
        setLbDragging(false)
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
          setStatus(res.message || '该图片目录下没有图片')
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

  // 缩放范围限制
  function clampZoom(z: number): number {
    return Math.min(8, Math.max(0.2, z))
  }

  // 以锚点（相对 wrap 中心的偏移 cx/cy）缩放，保持该点处内容不动
  function zoomAround(
    zRef: MutableRefObject<number>,
    pRef: MutableRefObject<{ x: number; y: number }>,
    setZ: (n: number) => void,
    setP: (n: { x: number; y: number }) => void,
    factor: number,
    cx: number,
    cy: number,
  ) {
    const z = zRef.current
    const nz = clampZoom(z * factor)
    if (nz === z) return
    const p = pRef.current
    const np = { x: cx - (nz / z) * (cx - p.x), y: cy - (nz / z) * (cy - p.y) }
    zRef.current = nz
    pRef.current = np
    // 缩回到 100% 及以下时，把平移归零，避免留白偏移
    if (nz <= 1.001) {
      pRef.current = { x: 0, y: 0 }
      setP({ x: 0, y: 0 })
      setZ(nz)
    } else {
      setZ(nz)
      setP(np)
    }
  }

  // 总览预览：以 wrap 中心为锚点的按钮缩放
  function zoomBy(factor: number) {
    const el = wrapRef.current
    if (!el) return
    const rect = el.getBoundingClientRect()
    zoomAround(zoomRef, panRef, setZoom, setPan, factor, rect.width / 2, rect.height / 2)
  }

  // 单张内联视图：以 wrap 中心为锚点的按钮缩放
  function singleZoomBy(factor: number) {
    const el = singleWrapRef.current
    if (!el) return
    const rect = el.getBoundingClientRect()
    zoomAround(singleZoomRef, singlePanRef, setSingleZoom, setSinglePan, factor, rect.width / 2, rect.height / 2)
  }

  // 放大查看（lightbox）：以自身中心为锚点的按钮缩放
  function lbZoomBy(factor: number) {
    zoomAround(lbZoomRef, lbPanRef, setLbZoom, setLbPan, factor, 0, 0)
  }

  function resetSingleView() {
    singleZoomRef.current = 1
    singlePanRef.current = { x: 0, y: 0 }
    setSingleZoom(1)
    setSinglePan({ x: 0, y: 0 })
  }
  function resetLbView() {
    lbZoomRef.current = 1
    lbPanRef.current = { x: 0, y: 0 }
    setLbZoom(1)
    setLbPan({ x: 0, y: 0 })
  }

  // 单张内联视图：左键拖拽平移 / 点击放大（拖拽后抑制点击，避免误触放大）
  function onSingleMouseDown(e: ReactMouseEvent) {
    if (e.button !== 0 || !activeTicket) return
    const p = singlePanRef.current
    singleDragRef.current = { startX: e.clientX, startY: e.clientY, panX: p.x, panY: p.y, moved: false }
    setSingleDragging(true)
  }
  function onSingleWrapClick() {
    if (singleClickGuard.current) {
      singleClickGuard.current = false
      return
    }
    setLightbox(true)
  }
  // 放大查看（lightbox）：左键拖拽平移
  function onLbMouseDown(e: ReactMouseEvent) {
    if (e.button !== 0) return
    const p = lbPanRef.current
    lbDragRef.current = { startX: e.clientX, startY: e.clientY, panX: p.x, panY: p.y }
    setLbDragging(true)
  }

  // 拖动分隔条调整「图片列表 / 预览」左右宽度
  function onSplitterDown(e: ReactMouseEvent) {
    e.preventDefault()
    const startX = e.clientX
    const startW = listWidth
    const onMove = (ev: MouseEvent) => {
      const nw = Math.min(200, Math.max(70, startW + (ev.clientX - startX)))
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
    // 检测会重建整组小票：清空上一批的识别状态与后台队列，避免旧任务串到新小票上
    recogMgr.current.cancelled = true
    recogMgr.current.queue = []
    recogMgr.current.running = false
    setBgRunning(false)
    setRecogState({})
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

  // 打开单张小票视图，并触发识别：本张优先识别，其余小票在后台顺序识别（不阻塞前台）。
  function openTicket(i: number) {
    if (i < 0 || i >= tickets.length) return
    setActive(i)
    setSingleRotate(0)
    resetSingleView()
    setViewMode('single')
    // 本张优先入队；其余顺序加入后台队列（已识别的会跳过）
    enqueueRecognition([i], true)
    const others = tickets.map((_, idx) => idx).filter((idx) => idx !== i)
    enqueueRecognition(others, false)
  }

  // 把若干小票加入后台识别队列（去重：已识别 / 已在队列中的跳过）。
  // front=true 表示优先识别（如用户刚点击的那张）。全部顺序处理，不阻塞前台操作。
  function enqueueRecognition(indices: number[], front = false) {
    if (!ticketsRef.current.length) return
    const toAdd: number[] = []
    for (const i of indices) {
      const hasRows = (ticketsRef.current[i]?.rows.length ?? 0) > 0
      if (hasRows) continue
      if (recogMgr.current.queue.includes(i)) continue
      toAdd.push(i)
    }
    if (toAdd.length === 0) return
    setRecogState((p) => {
      const np = { ...p }
      for (const i of toAdd) if (np[i] !== 'busy') np[i] = 'pending'
      return np
    })
    if (front) recogMgr.current.queue.unshift(...toAdd)
    else recogMgr.current.queue.push(...toAdd)
    if (!recogMgr.current.running) {
      recogMgr.current.running = true
      recogMgr.current.cancelled = false
      setBgRunning(true)
      void drainRecognition()
    }
  }

  // 后台顺序识别循环：一次只识别一张，逐个 await，期间不阻塞 UI；
  // 用户在队列跑动时仍可随意点击其它小票（会被重新插到队首优先识别）。
  async function drainRecognition() {
    const mgr = recogMgr.current
    let processed = 0
    while (mgr.queue.length > 0 && !mgr.cancelled) {
      const i = mgr.queue.shift()!
      // 再次确认：若已被识别（例如用户刚手动识别过 / 已缓存），直接跳过
      if ((ticketsRef.current[i]?.rows.length ?? 0) > 0) {
        setRecogState((p) => ({ ...p, [i]: 'done' }))
        continue
      }
      setStatus(`正在后台识别第 ${i + 1} 张小票…（队列剩余 ${mgr.queue.length} 张）`)
      await recognizeTicketByIdx(i)
      processed++
    }
    mgr.running = false
    setBgRunning(false)
    if (!mgr.cancelled && processed > 0) {
      setStatus(`后台识别完成：本组共识别 ${processed} 张小票，结果已缓存，可直接录入`)
    }
  }

  // 取消后台识别（当前这张识别完即停，不再处理后续队列）
  function cancelRecognition() {
    recogMgr.current.cancelled = true
    recogMgr.current.queue = []
    recogMgr.current.running = false
    setBgRunning(false)
    setStatus('已取消后台识别队列')
  }

  // 核心：识别第 i 张小票（前台 / 后台共用）。
  // 日期规则：仅「首张被识别」的小票贡献日期作为全局基准；之后所有小票的日期都套用
  // 该基准（不各自再识别日期），保证一整组小票使用同一日期。
  async function recognizeTicketByIdx(i: number) {
    const t = ticketsRef.current[i]
    if (!t || !t.crop) {
      setRecogState((p) => ({ ...p, [i]: 'error' }))
      setStatus(`第 ${i + 1} 张没有可用的裁剪图`)
      return
    }
    setRecogState((p) => ({ ...p, [i]: 'busy' }))
    try {
      const res = await api.aiRecognizeCrop(t.crop)
      if (res.error) {
        setStatus(res.message || `第 ${i + 1} 张识别失败`)
        setRecogState((p) => ({ ...p, [i]: 'error' }))
        return
      }
      const src = baseName(selected)
      let rs = (res.rows || []).map((r) => ({ ...r, source: src }))
      // 日期基准：仅首张贡献日期；其余套用基准，丢弃各自识别出的日期
      const anchor = dateAnchorRef.current
      if (anchor == null) {
        const firstDate = rs.find((r) => (r.date || '').trim())?.date?.trim()
        if (firstDate) dateAnchorRef.current = firstDate
        rs = rs.map((r) => ({ ...r, date: dateAnchorRef.current ?? r.date }))
      } else {
        rs = rs.map((r) => ({ ...r, date: anchor }))
      }
      setTickets((prev) => {
        const next = prev.slice()
        next[i] = { ...next[i], rows: rs }
        return next
      })
      setRecogState((p) => ({ ...p, [i]: 'done' }))
    } catch (err) {
      setStatus('识别失败：' + (err instanceof Error ? err.message : '未知错误'))
      setRecogState((p) => ({ ...p, [i]: 'error' }))
    }
  }

  // 逐张识别：对检测出来但还没识别的小票，调用 AI 识别这一张裁剪图（前台手动触发，走队列）
  function recognizeThisCrop(i: number) {
    enqueueRecognition([i], true)
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
    // 独立窗口（带 initialState）的状态已由 useLayoutEffect 的 restoreState 恢复，
    // 不能再自动 loadImages，否则会把 selected 重置到第一张、并清空已识别的 tickets/rows
    if (initialState) return
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

  // 是否展示"来源"列：仅当结果来自多张不同图片时（单图识别不显示，避免冗余）
  const distinctSources = new Set(rows.map((r) => r.source).filter(Boolean))
  const showSource = distinctSources.size > 1

  const activeTicket = tickets[active]
  const isLast = active >= tickets.length - 1
  const isFirst = active <= 0
  // 当前单张是否正在识别（前台/后台共用此状态，取代原 singleLoading）
  const singleBusy = recogState[active] === 'busy'

  return (
    <div className="image-window">
      <div className="image-panel">
        <div className="image-panel-header">
          <div className="image-panel-head-row">
            <span className="image-panel-title">小票识图</span>
            <button
              className="btn btn-small btn-outline"
              onClick={chooseDir}
              title="选择小票图片所在目录"
            >
              选择目录
            </button>
            <div className="head-actions">
              <button
                className="btn btn-small btn-outline"
                onClick={() => setListCollapsed((c) => !c)}
                title="折叠/展开左侧图片列表：折叠后识别信息占满整个窗口宽度，便于查看与录入"
              >
                {listCollapsed ? '☰ 显示列表' : '☰ 隐藏列表'}
              </button>
              {detached ? (
                <button
                  className="btn btn-small btn-outline"
                  onClick={() => onAttach?.(captureState())}
                  title="把本窗口合并回主窗口右侧面板（保留此处的结果与进度）"
                >
                  ⎘ 合并回主窗口
                </button>
              ) : (
                <button
                  className="btn btn-small btn-outline"
                  onClick={() => onDetach?.(captureState())}
                  title="把本面板拆成独立窗口，方便在大屏上单独查看 / 对照录入（保留当前结果与进度）"
                >
                  ⧉ 拆分窗口
                </button>
              )}
            </div>
          </div>
          {imageDir && (
            <span className="image-panel-dir" title={imageDir}>
              {imageDir}
            </span>
          )}
        </div>

        {/* 顶部工具栏：旋转 / 缩放 / 当前图 / 单张导航 —— 统一上移到面板最顶部 */}
        <div className="image-topbar">
          {viewMode === 'single' && activeTicket ? (
            <>
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
              <span className="topbar-spacer" />
              <button className="btn btn-small btn-outline" onClick={() => setSingleRotate((r) => r - 90)} title="向左旋转 90°">
                ↺ 左转90°
              </button>
              <button className="btn btn-small btn-outline" onClick={() => setSingleRotate((r) => r + 90)} title="向右旋转 90°">
                ↻ 右转90°
              </button>
              <button className="btn btn-small btn-link" onClick={() => setSingleRotate(0)} title="复位旋转">
                复位
              </button>
              <span className="image-zoom-sep" />
              <button className="btn btn-small btn-outline" onClick={() => singleZoomBy(1 / 1.2)} title="缩小">－</button>
              <button className="btn btn-small btn-outline" onClick={() => singleZoomBy(1.2)} title="放大">＋</button>
              <span className="image-zoom">{Math.round(singleZoom * 100)}%</span>
            </>
          ) : (
            <>
              <button className="btn btn-small btn-outline" onClick={() => setRotate((r) => (r + 270) % 360)} disabled={!preview} title="向左旋转 90°">
                ↺ 左转
              </button>
              <button className="btn btn-small btn-outline" onClick={() => setRotate((r) => (r + 90) % 360)} disabled={!preview} title="向右旋转 90°">
                ↻ 右转
              </button>
              <button className="btn btn-small btn-link" onClick={() => { setRotate(0); resetView() }} disabled={!preview} title="复位旋转与缩放/位置">
                复位
              </button>
              <button className="btn btn-small btn-outline" onClick={() => zoomBy(1 / 1.2)} disabled={!preview} title="缩小">－</button>
              <button className="btn btn-small btn-outline" onClick={() => zoomBy(1.2)} disabled={!preview} title="放大">＋</button>
              <span className="image-rotate-badge">旋转 {rotate}°</span>
              <span className="image-zoom">{Math.round(zoom * 100)}%</span>
              {selected && (
                <span className="image-current-name" title={selected}>
                  {baseName(selected)}
                </span>
              )}
            </>
          )}
        </div>

        <div className="image-content">
          {/* 左侧：可见的图片列表（可折叠，折叠后识别信息占满窗口宽度） */}
          {!listCollapsed && (
            <>
              <div className="image-list-side" style={{ width: listWidth, flex: '0 1 auto' }}>
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
            </>
          )}
          <div className="image-main-area">
            {viewMode === 'single' && activeTicket ? (
              /* ====== 单张小票放大视图 ====== */
              <div className="single-ticket-view">
                <div
                  className="single-image-wrap"
                  ref={singleWrapRef}
                  onMouseDown={onSingleMouseDown}
                  onClick={onSingleWrapClick}
                  style={{ cursor: singleDragging ? 'grabbing' : 'zoom-in' }}
                  title="滚轮缩放 · 拖拽平移 · 点击放大查看"
                >
                  <img
                    src={`data:image/jpeg;base64,${activeTicket.crop}`}
                    alt={`第 ${active + 1} 张小票`}
                    className="single-image"
                    draggable={false}
                    style={{
                      transform: `translate(${singlePan.x}px, ${singlePan.y}px) rotate(${singleRotate}deg) scale(${singleZoom})`,
                      transition: singleDragging ? 'none' : 'transform 0.2s',
                    }}
                  />
                  <span className="single-zoom-hint">🔍 滚轮缩放 · 点击放大</span>
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
                          disabled={singleBusy || !activeTicket.crop}
                        >
                          {singleBusy ? '识别中…' : '识别此张'}
                        </button>
                      ) : (
                        <button
                          className="btn btn-small btn-outline"
                          onClick={() => recognizeThisCrop(active)}
                          disabled={singleBusy || !activeTicket.crop}
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
                      {singleBusy ? '识别中…' : '本张小票尚未识别，点击右上角「识别此张」进行 AI 识别。'}
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
                  {bgRunning && (
                    <div className="bg-indicator" title="正在后台顺序识别各张小票，不阻塞前台；可取消">
                      <span className="bg-spinner" />
                      <span>后台识别中…</span>
                      <button className="btn btn-small btn-link" onClick={cancelRecognition}>
                        取消
                      </button>
                    </div>
                  )}
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
                          className={'ticket-thumb' + (viewMode === 'single' && active === i ? ' active' : '') + (t.rows.length ? '' : ' unrec') + (recogState[i] === 'busy' ? ' busy' : '') + (recogState[i] === 'error' ? ' err' : '')}
                          onClick={() => openTicket(i)}
                          title={`第 ${i + 1} 张${t.rows.length ? '（已识别）' : recogState[i] === 'busy' ? '（识别中…）' : '（未识别）'}`}
                        >
                          {t.crop ? (
                            <img src={`data:image/jpeg;base64,${t.crop}`} alt={`第 ${i + 1} 张`} draggable={false} />
                          ) : (
                            <span className="ticket-thumb-no">无图</span>
                          )}
                          <span className="ticket-thumb-idx">#{i + 1}</span>
                          {recogState[i] === 'busy' && <span className="ticket-thumb-flag busy">识别中…</span>}
                          {recogState[i] === 'error' && <span className="ticket-thumb-flag err">失败</span>}
                          {recogState[i] !== 'busy' && recogState[i] !== 'error' && t.rows.length === 0 && <span className="ticket-thumb-flag">待识别</span>}
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

        {/* 单张小票放大查看（lightbox）：点击图片弹出居中大图，Esc / 点击空白关闭 */}
        {lightbox && activeTicket && (
          <div className="image-lightbox" onClick={() => setLightbox(false)}>
            <div className="image-lightbox-box" onClick={(e) => e.stopPropagation()}>
              <div className="image-lightbox-bar">
                <span className="image-lightbox-title">第 {active + 1} 张小票 · 滚轮缩放 / 拖拽平移 · 点击空白或按 Esc 关闭</span>
                <div className="image-lightbox-actions">
                  <button className="btn btn-small btn-outline" onClick={() => lbZoomBy(1 / 1.2)} title="缩小">－</button>
                  <button className="btn btn-small btn-outline" onClick={() => lbZoomBy(1.2)} title="放大">＋</button>
                  <button className="btn btn-small btn-outline" onClick={resetLbView} title="复位缩放与位置">复位</button>
                  <button
                    className="btn btn-small btn-outline"
                    onClick={() => setSingleRotate((r) => r - 90)}
                    title="向左旋转 90°"
                  >
                    ↺ 左转90°
                  </button>
                  <button
                    className="btn btn-small btn-outline"
                    onClick={() => setSingleRotate((r) => r + 90)}
                    title="向右旋转 90°"
                  >
                    ↻ 右转90°
                  </button>
                  <button
                    className="btn btn-small btn-link"
                    onClick={() => setLightbox(false)}
                  >
                    关闭
                  </button>
                </div>
              </div>
              <div className="image-lightbox-body" ref={lbWrapRef}>
                <img
                  src={`data:image/jpeg;base64,${activeTicket.crop}`}
                  alt={`第 ${active + 1} 张小票`}
                  className="image-lightbox-img"
                  draggable={false}
                  onMouseDown={onLbMouseDown}
                  style={{
                    transform: `translate(${lbPan.x}px, ${lbPan.y}px) rotate(${singleRotate}deg) scale(${lbZoom})`,
                    transition: lbDragging ? 'none' : 'transform 0.15s',
                    cursor: lbDragging ? 'grabbing' : lbZoom > 1 ? 'grab' : 'zoom-out',
                  }}
                />
              </div>
            </div>
          </div>
        )}
      </div>
    </div>
  )
})
