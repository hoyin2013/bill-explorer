import { useCallback, useEffect, useRef, useState } from 'react'
import { FileEntry, ElectronAPI, AIRecognizedRow } from '../types'

/* 固定 9 列（列数不动，就这么多列），type 决定编辑控件与写回格式 */
interface ColDef {
  field: string
  label: string
  type: 'text' | 'number' | 'date'
  width: number
}
const COLUMNS: ColDef[] = [
  { field: 'no', label: '序号', type: 'text', width: 40 },
  { field: 'date', label: '日期', type: 'text', width: 90 },
  { field: 'name', label: '货品名称', type: 'text', width: 172 },
  { field: 'unit', label: '单位', type: 'text', width: 60 },
  { field: 'qty', label: '数量', type: 'number', width: 56 },
  { field: 'price', label: '单价', type: 'number', width: 72 },
  { field: 'amount', label: '金额', type: 'number', width: 84 },
  { field: 'person', label: '调货人', type: 'text', width: 92 },
  { field: 'remark', label: '备注', type: 'text', width: 150 },
]
const COL_COUNT = COLUMNS.length
const qtyIdx = COLUMNS.findIndex((c) => c.field === 'qty')
const priceIdx = COLUMNS.findIndex((c) => c.field === 'price')
const amountIdx = COLUMNS.findIndex((c) => c.field === 'amount')
const noIdx = COLUMNS.findIndex((c) => c.field === 'no')
const MIN_COL_WIDTH = 50
const PREPARED_EMPTY_ROWS = 12 // 打开时在已有数据下方预备的空行，方便连续快捷录入

function emptyRow(): string[] {
  return Array.from({ length: COL_COUNT }, () => '')
}
function normalizeRow(r?: string[]): string[] {
  const row = r ? r.slice(0, COL_COUNT) : []
  while (row.length < COL_COUNT) row.push('')
  return row
}
// 数字保留最多 2 位小数并去掉无意义尾随 0
function fmtNum(n: number): string {
  if (!isFinite(n)) return ''
  const r = Math.round(n * 100) / 100
  return String(r)
}
// 整行是否为空（用于判定录入行）
function isRowEmpty(row?: string[]): boolean {
  return !row || row.every((c) => !String(c).trim())
}
// A3 新行预填：导航到空行时，从上一行带走这些列的默认值（同批录单据常相同的列）。
// 数量/单价/金额/货品名/序号 不预填（每笔不同）。
const SEED_FIELDS = new Set(['date', 'person', 'remark'])
// 真正构成一笔记录的列：用于保存时剥离"只有预填默认值、无实质内容"的空行
const SUBSTANCE_FIELDS = new Set(['no', 'name', 'qty', 'price', 'amount'])
function rowHasSubstance(row?: string[]): boolean {
  if (!row) return false
  return COLUMNS.some(
    (col, i) => SUBSTANCE_FIELDS.has(col.field) && String(row[i] ?? '').trim() !== '',
  )
}
// 从上一行构造"预填默认值行"（只填 SEED_FIELDS 列）
function buildSeedRow(src?: string[]): string[] {
  const row = emptyRow()
  if (src) COLUMNS.forEach((col, i) => { if (SEED_FIELDS.has(col.field)) row[i] = String(src[i] ?? '') })
  return row
}
// 把各种来源的日期文本规范化成 yyyy-mm-dd（无法识别则返回空串）
// 支持：2026-8-11 / 2026/8/11 / 2026.8.11 / 20260811 / 2026年8月11日 / 8-11（补当年）/ Excel 日期序列号
function parseDateText(input: string): string {
  const s = String(input ?? '').trim()
  if (!s) return ''
  const pad = (n: number) => String(n).padStart(2, '0')
  const make = (y: number, m: number, d: number): string => {
    if (y < 100) y += y < 70 ? 2000 : 1900
    if (m < 1 || m > 12 || d < 1 || d > 31) return ''
    const dt = new Date(y, m - 1, d)
    // 校验真实存在（排除 2 月 30 日之类）
    if (dt.getFullYear() !== y || dt.getMonth() !== m - 1 || dt.getDate() !== d) return ''
    return `${y}-${pad(m)}-${pad(d)}`
  }
  // 年月日 / 带分隔符
  let m = s.match(/^(\d{4}|\d{2})\s*[-/.年]\s*(\d{1,2})\s*[-/.月]\s*(\d{1,2})\s*日?$/)
  if (m) return make(Number(m[1]), Number(m[2]), Number(m[3]))
  // 纯 8 位：20260811
  m = s.match(/^(\d{4})(\d{2})(\d{2})$/)
  if (m) return make(Number(m[1]), Number(m[2]), Number(m[3]))
  // 只有月日：补当前年份（不接受点分隔，"12.5" 这类小数与月日无法区分，宁可不认）
  m = s.match(/^(\d{1,2})\s*[-/月]\s*(\d{1,2})\s*日?$/)
  if (m) return make(new Date().getFullYear(), Number(m[1]), Number(m[2]))
  // Excel 日期序列号（1900 日期系统，起点 1899-12-30）
  if (/^\d{4,5}$/.test(s)) {
    const serial = Number(s)
    if (serial >= 1 && serial <= 60000) {
      const dt = new Date(Date.UTC(1899, 11, 30) + serial * 86400000)
      return make(dt.getUTCFullYear(), dt.getUTCMonth() + 1, dt.getUTCDate())
    }
  }
  return ''
}

// 几何命中：根据鼠标坐标定位所在单元格（比 elementFromPoint 更稳，
// 不会被填充柄 / 行号 / 单元格内边距干扰；拖到行号或列间隙也能正确归位）。
// trs 为 tbody 内的数据行（顺序即行号 0..n）。
function rowIndexAtY(trs: HTMLTableRowElement[], clientY: number): number {
  const last = trs.length - 1
  if (last < 0) return 0
  for (let i = 0; i <= last; i += 1) {
    const rect = trs[i].getBoundingClientRect()
    if (clientY <= rect.bottom) return i
  }
  return last
}
function cellAtPoint(
  trs: HTMLTableRowElement[],
  clientX: number,
  clientY: number,
): { r: number; c: number } {
  if (trs.length === 0) return { r: 0, c: 0 }
  const r = rowIndexAtY(trs, clientY)
  const tds = trs[r].cells // [0]=行号, [1+c]=第 c 列
  let c = 0
  for (let i = 0; i < COL_COUNT; i += 1) {
    const rect = tds[i + 1]?.getBoundingClientRect()
    if (!rect) break
    if (clientX <= rect.right) {
      c = i
      break
    }
    c = i
  }
  return { r, c }
}

interface Props {
  file: FileEntry
  api: ElectronAPI
  onClose: () => void
  onSaved: () => void
}

export function SheetGrid({ file, api, onClose, onSaved }: Props) {
  const [grid, setGrid] = useState<string[][]>([])
  const [active, setActive] = useState<{ r: number; c: number }>({ r: 0, c: 0 })
  const [editing, setEditing] = useState(false)
  const [colWidths, setColWidths] = useState<number[]>(COLUMNS.map((c) => c.width))
  const [dirty, setDirty] = useState(false)
  const [saving, setSaving] = useState(false)
  const [status, setStatus] = useState('')
  const [loaded, setLoaded] = useState(false)
  // 序号列拖拽填充（Excel 式填充柄）：startR=种子行，endR=拖动到的行，拖动时高亮该区间
  const [fill, setFill] = useState<{ startR: number; endR: number } | null>(null)
  const fillDragRef = useRef<{ startR: number; seed: number; valid: boolean; endR: number } | null>(null)
  // 矩形框选（鼠标拖拽选 N×M 单元格）：r1/c1=锚点，r2/c2=当前拉伸到的格
  const [selRange, setSelRange] = useState<{ r1: number; c1: number; r2: number; c2: number } | null>(null)

  // 接收独立"小票识图"窗口发来的回填请求（用户主动点击"填入当前录入"才触发，不影响录入）
  const applyRowsRef = useRef(applyRecognizedRows)
  applyRowsRef.current = applyRecognizedRows
  useEffect(() => {
    const off = api.on('apply-recognized-rows', (rows) => {
      if (Array.isArray(rows)) applyRowsRef.current(rows as AIRecognizedRow[], true)
    })
    return off
  }, [api])
  // 由 AI 图片识别自动预录的行（绝对行号集合）：这些行用不同底色高亮，便于和手输区分
  const [autoRows, setAutoRows] = useState<Set<number>>(new Set())
  // 当前"录入行"：单击已有内容单元格时把值复制到这一行；键盘落到空行时自动跟随
  const [inputRow, setInputRow] = useState(0)
  // 点行号选中的整行（可 Ctrl 多选 / Shift 范围选），右键菜单的"删除行"作用于此
  const [selRows, setSelRows] = useState<Set<number>>(new Set())
  // 右键菜单：屏幕坐标 + 触发处的行列
  const [menu, setMenu] = useState<{ x: number; y: number; r: number; c: number } | null>(null)
  // 仅用于让菜单里的撤销/重做项能正确禁用
  const [histLen, setHistLen] = useState({ undo: 0, redo: 0 })
  // A1 记忆式自动补全：编辑文本列时，按前缀匹配本表同列历史值的下拉建议
  const [suggest, setSuggest] = useState<{
    r: number
    c: number
    items: string[]
    hi: number
    rect: { left: number; top: number; width: number } | null
  } | null>(null)
  // A5 文件导入
  const fileInputRef = useRef<HTMLInputElement>(null)

  const containerRef = useRef<HTMLDivElement>(null)
  const editRef = useRef<HTMLTextAreaElement | HTMLInputElement | null>(null)

  // 供事件回调读取最新值（避免闭包拿到旧 state）
  const activeRef = useRef(active)
  activeRef.current = active
  const editingRef = useRef(editing)
  editingRef.current = editing
  const gridRef = useRef(grid)
  gridRef.current = grid
  const inputRowRef = useRef(inputRow)
  inputRowRef.current = inputRow
  // 指向当前激活单元格的 <td>，用于滚动到可视区域
  const activeTdRef = useRef<HTMLTableCellElement | null>(null)
  const selRowsRef = useRef(selRows)
  selRowsRef.current = selRows
  // 矩形框选（供事件回调读最新值）
  const selRangeRef = useRef(selRange)
  selRangeRef.current = selRange
  // 框选拖拽过程的状态
  const rangeDragRef = useRef<{
    r1: number
    c1: number
    x: number
    y: number
    moved: boolean
    trs: HTMLTableRowElement[]
    rowMode: boolean
  } | null>(null)
  // 框选拖拽结束的尾随 click 不处理（否则会把刚框好的选区又清掉）
  const suppressClickRef = useRef(false)
  // 撤销 / 重做栈：存整表快照（string[][]），账本行数量级下开销可忽略
  const undoRef = useRef<string[][][]>([])
  const redoRef = useRef<string[][][]>([])
  // 同一个单元格连续键入只记一步撤销（否则每个字符都成为一步）
  const lastCommitKeyRef = useRef<string | null>(null)
  // 自己复制过的文本，作为读不到系统剪贴板时的兜底
  const lastCopyRef = useRef('')
  // A3 新行预填：行号 → 从上一行带过来的默认值行（只在"首次录入该行"时合并进 grid，
  // 之后即清除，因此不会污染、也不会被保存成空行）。键为行号。
  const rowDefaultsRef = useRef<Map<number, string[]>>(new Map())

  /* ===================== 撤销 / 重做 ===================== */
  const MAX_UNDO = 200
  const syncHist = useCallback(() => {
    setHistLen({ undo: undoRef.current.length, redo: redoRef.current.length })
  }, [])
  // 变更前调用：把当前整表压入撤销栈（并清空重做栈）
  const pushUndo = useCallback(() => {
    undoRef.current.push(gridRef.current.map((row) => row.slice()))
    if (undoRef.current.length > MAX_UNDO) undoRef.current.shift()
    redoRef.current = []
    syncHist()
  }, [syncHist])

  const undo = useCallback(() => {
    const snap = undoRef.current.pop()
    if (!snap) {
      setStatus('没有可撤销的操作')
      return
    }
    redoRef.current.push(gridRef.current.map((row) => row.slice()))
    setGrid(snap)
    setEditing(false)
    setDirty(true)
    setSelRows(new Set())
    setSelRange(null)
    setAutoRows(new Set())
    lastCommitKeyRef.current = null
    setActive((a) => ({ r: Math.min(a.r, Math.max(0, snap.length - 1)), c: a.c }))
    syncHist()
    setStatus('已撤销')
    containerRef.current?.focus()
  }, [syncHist])

  const redo = useCallback(() => {
    const snap = redoRef.current.pop()
    if (!snap) {
      setStatus('没有可重做的操作')
      return
    }
    undoRef.current.push(gridRef.current.map((row) => row.slice()))
    setGrid(snap)
    setEditing(false)
    setDirty(true)
    setAutoRows(new Set())
    setSelRange(null)
    lastCommitKeyRef.current = null
    setActive((a) => ({ r: Math.min(a.r, Math.max(0, snap.length - 1)), c: a.c }))
    syncHist()
    setStatus('已重做')
    containerRef.current?.focus()
  }, [syncHist])

  // 最底下可录入的位置 = 最后一个有内容的行之后
  // （不用"第一个空行"，避免表中间存在分隔空行时光标错误落在中间）
  const bottomEntryRow = useCallback((rows: string[][]): number => {
    for (let i = rows.length - 1; i >= 0; i -= 1) {
      if (!isRowEmpty(rows[i])) return i + 1
    }
    return 0
  }, [])

  /* ---- 加载整张表：光标定位到数据末尾第一列，并在下方预备一批空行 ---- */
  useEffect(() => {
    let cancelled = false
    setLoaded(false)
    setStatus('读取中…')
    // 换文件即丢弃历史，避免把撤销应用到另一张表
    undoRef.current = []
    redoRef.current = []
    lastCommitKeyRef.current = null
    setHistLen({ undo: 0, redo: 0 })
    setSelRows(new Set())
    setSelRange(null)
    setMenu(null)
    ;(async () => {
      const res = await api.loadSheet(file.filePath)
      if (cancelled) return
      if (res.error) {
        setStatus('读取失败：' + (res.message || '未知错误'))
        setLoaded(true)
        return
      }
      const rows = (res.rows || []).map((r) => normalizeRow(r))
      // 录入起点 = 最后一条有内容的记录之后（最底下）
      const target = bottomEntryRow(rows)
      // 在下方预备一批空行，方便连续快捷录入
      while (rows.length < target + PREPARED_EMPTY_ROWS) rows.push(emptyRow())
      setGrid(rows)
      setAutoRows(new Set())
      setInputRow(target)
      setActive({ r: target, c: 0 }) // 光标定位到最底下（数据末尾）第一列（序号列），便于用填充柄批量生成序号
      setLoaded(true)
      setStatus('')
      // 不自动进入编辑态：打开时不要弹出序号编辑框（需要录入时单击/双击或按字符即可进入）
    })()
    return () => {
      cancelled = true
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [file.filePath])

  // 编辑态切换 / 焦点变化时，把光标聚焦到编辑控件并移到末尾
  // 日期列额外主动弹出原生日期选择器：单纯 focus() 不会展开日历，必须调 showPicker()
  useEffect(() => {
    if (!editing || !editRef.current) return
    const el = editRef.current
    el.focus()
    // 文本框/文本域：光标移到末尾，接着打字即可
    const len = el.value.length
    try {
      el.setSelectionRange(len, len)
    } catch {
      /* 某些 input 类型不支持 setSelectionRange */
    }
  }, [editing, active])

  // 激活单元格滚动到可视区域：首次（打开时）居中，便于同时看到下方预备的空行；
  // 之后按最小幅度滚动，避免导航时画面跳动
  const didInitialScroll = useRef(false)
  useEffect(() => {
    if (!loaded || !activeTdRef.current) return
    const first = !didInitialScroll.current
    activeTdRef.current.scrollIntoView({
      block: first ? 'center' : 'nearest',
      inline: 'nearest',
    })
    if (first) didInitialScroll.current = true
  }, [active, loaded])

  // 切换文件时重置首次滚动标记
  useEffect(() => {
    didInitialScroll.current = false
  }, [file.filePath])

  // 换格即断开撤销合并：下一格的第一次输入会成为独立的一步
  // （只依赖 active，不依赖 editing —— 否则"打字直接进编辑"会被拆成两步撤销）
  useEffect(() => {
    lastCommitKeyRef.current = null
  }, [active.r, active.c])

  /* ===================== 数据变更 ===================== */
  const commitCell = useCallback((r: number, c: number, value: string) => {
    // 同一格连续键入合并成一步撤销；换格/重进编辑会重置 key（见下方 effect）
    const key = `${r}:${c}`
    if (lastCommitKeyRef.current !== key) {
      pushUndo()
      lastCommitKeyRef.current = key
    }
    setGrid((prev) => {
      const next = prev.map((row) => row.slice())
      if (!next[r]) next[r] = emptyRow()
      // A3 新行预填：首次在该空行录入时，把"从上一行带过来的默认值"合并进来（只此一次）
      if (rowDefaultsRef.current.has(r) && isRowEmpty(next[r])) {
        const seed = rowDefaultsRef.current.get(r)!
        COLUMNS.forEach((col, i) => {
          if (SEED_FIELDS.has(col.field) && seed[i]) next[r][i] = seed[i]
        })
        rowDefaultsRef.current.delete(r)
      }
      next[r][c] = value
      const f = COLUMNS[c].field
      // 数量/单价变化时自动重算金额（两者都有值时）
      if (f === 'qty' || f === 'price') {
        const q = Number(next[r][qtyIdx] || 0)
        const p = Number(next[r][priceIdx] || 0)
        if (q > 0 && p > 0) next[r][amountIdx] = fmtNum(q * p)
      }
      return next
    })
    setDirty(true)
    // 用户手动改过这一格 → 该行不再是"纯 AI 预录"，去掉高亮色
    setAutoRows((prev) => {
      if (!prev.has(r)) return prev
      const n = new Set(prev)
      n.delete(r)
      return n
    })
  }, [pushUndo])

  const clearCell = useCallback((r: number, c: number) => {
    pushUndo()
    lastCommitKeyRef.current = null
    setGrid((prev) => {
      const next = prev.map((row) => row.slice())
      if (next[r]) next[r][c] = ''
      return next
    })
    setDirty(true)
  }, [pushUndo])

  /* ===================== 焦点 / 导航 ===================== */
  const focusContainer = () => containerRef.current?.focus()

  const selectCell = useCallback((r: number, c: number) => {
    setActive({ r, c })
    setEditing(false)
    focusContainer()
  }, [])

  const startEdit = useCallback((r: number, c: number, initial?: string) => {
    if (initial !== undefined) commitCell(r, c, initial)
    setActive({ r, c })
    setEditing(true)
  }, [commitCell])

  // 复制文本到系统剪贴板（隐藏 textarea + execCommand，Electron 渲染进程可用）
  const copyText = useCallback((text: string) => {
    lastCopyRef.current = text // 兜底：读不到系统剪贴板时用它
    try {
      const ta = document.createElement('textarea')
      ta.value = text
      ta.style.position = 'fixed'
      ta.style.opacity = '0'
      document.body.appendChild(ta)
      ta.select()
      document.execCommand('copy')
      document.body.removeChild(ta)
    } catch {
      /* ignore */
    }
  }, [])

  // 单击单元格：选中并激活该格（复制请用 Ctrl+C / 右键菜单 / Ctrl+X 剪切）
  const handleCellClick = useCallback((r: number, c: number) => {
    // 框选拖拽结束的尾随 click 不处理，否则会把刚框好的选区又清掉
    if (suppressClickRef.current) return
    // 点普通单元格即取消整行选择（Excel 习惯），避免 Delete 误清掉之前选的行
    if (selRowsRef.current.size > 0) setSelRows(new Set())
    if (isRowEmpty(gridRef.current[r])) setInputRow(r)
    selectCell(r, c)
  }, [selectCell])

  // 矩形框选：在单元格上按下左键并拖动，框选 N×M 区域（Excel 习惯）。
  // 仅左键、非编辑态触发；移动超过阈值(4px)才算拖拽，否则视作普通单击（交给 onClick）。
  // 与序号列填充柄、行号整行选择互不影响：填充柄 onMouseDown 已 stopPropagation，行号是独立元素。
  const onCellMouseDown = useCallback((e: React.MouseEvent, r: number, c: number) => {
    if (editingRef.current || e.button !== 0) return
    e.stopPropagation()
    if (selRowsRef.current.size > 0) setSelRows(new Set())
    const tbody = containerRef.current?.querySelector('tbody')
    const trs = tbody ? (Array.from(tbody.rows) as HTMLTableRowElement[]) : []
    rangeDragRef.current = { r1: r, c1: c, x: e.clientX, y: e.clientY, moved: false, trs, rowMode: false }

    const onMove = (me: MouseEvent) => {
      const d = rangeDragRef.current
      if (!d) return
      if (!d.moved && Math.hypot(me.clientX - d.x, me.clientY - d.y) < 4) return
      d.moved = true
      // 几何定位光标所在单元格（不依赖 elementFromPoint，避免命中填充柄/行号/内边距）
      const { r: rr, c: cc } = cellAtPoint(d.trs, me.clientX, me.clientY)
      setSelRange({ r1: d.r1, c1: d.c1, r2: rr, c2: cc })
      setActive({ r: rr, c: cc })
    }
    const onUp = () => {
      window.removeEventListener('mousemove', onMove)
      window.removeEventListener('mouseup', onUp)
      const d = rangeDragRef.current
      rangeDragRef.current = null
      if (!d || !d.moved) {
        // 没移动 → 普通单击，清除框选（单选交给 onClick 处理）
        setSelRange(null)
      } else {
        // 拖拽结束：激活格回到拖拽起点（Excel 习惯：活动单元格=按下处），
        // 并吞掉尾随 click，避免刚框好的选区被单击清掉
        setActive({ r: d.r1, c: d.c1 })
        suppressClickRef.current = true
        window.setTimeout(() => {
          suppressClickRef.current = false
        }, 0)
      }
      focusContainer()
    }
    window.addEventListener('mousemove', onMove)
    window.addEventListener('mouseup', onUp)
  }, [focusContainer])

  // 序号列填充柄：按住右下角小方块向下拖动，按 Excel 习惯生成 1,2,3… 递增序列。
  // 种子为该序号格当前数值；拖动区间 [min,max] 内除种子外 = seed ± offset。
  // 仅作用于序号列（noIdx），避免数量/单价/金额等数值列被误填充。
  const onFillHandleDown = useCallback((e: React.MouseEvent) => {
    e.preventDefault()
    e.stopPropagation()
    const sr = activeRef.current.r
    const seedStr = String(gridRef.current[sr]?.[noIdx] ?? '').trim()
    const seedNum = Number(seedStr)
    // 仅当种子为纯整数时才做递增序列；否则（空/非数字）退化为整列复制
    const valid = seedStr !== '' && Number.isFinite(seedNum) && !/\D/.test(seedStr)
    const info: { startR: number; seed: number; valid: boolean; endR: number } = {
      startR: sr,
      seed: seedNum,
      valid,
      endR: sr,
    }
    fillDragRef.current = info
    setFill({ startR: sr, endR: sr })
    document.body.style.cursor = 'crosshair'

    const onMove = (me: MouseEvent) => {
      const td = activeTdRef.current
      if (!td) return
      const rect = td.getBoundingClientRect()
      const rowH = rect.height || 22
      const off = Math.round((me.clientY - rect.top) / rowH)
      let endR = info.startR + off
      endR = Math.max(0, Math.min(gridRef.current.length - 1, endR))
      info.endR = endR
      setFill({ startR: info.startR, endR })
    }
    const onUp = () => {
      window.removeEventListener('mousemove', onMove)
      window.removeEventListener('mouseup', onUp)
      document.body.style.cursor = ''
      const f = fillDragRef.current
      fillDragRef.current = null
      setFill(null)
      if (!f || f.endR === f.startR) return
      pushUndo()
      lastCommitKeyRef.current = null
      const minR = Math.min(f.startR, f.endR)
      const maxR = Math.max(f.startR, f.endR)
      setGrid((prev) => {
        const next = prev.map((row) => row.slice())
        for (let rr = minR; rr <= maxR; rr += 1) {
          if (rr === f.startR) continue
          const offset = rr - f.startR
          const v = f.valid ? String(f.seed + offset) : next[f.startR]?.[noIdx] ?? ''
          if (!next[rr]) next[rr] = emptyRow()
          next[rr][noIdx] = v
        }
        return next
      })
      setDirty(true)
      setStatus(`已填充序号序列 ${Math.abs(f.endR - f.startR)} 行`)
      focusContainer()
    }
    window.addEventListener('mousemove', onMove)
    window.addEventListener('mouseup', onUp)
  }, [pushUndo, focusContainer])

  // 序号自动编号：从 start 行向下，给序号列依次填 1,2,3…，直到末尾数据行。
  // 种子取本行已有整数（可续编），否则从 1 开始；末尾行 = 自底向上第一个非空行，避免给空白预备行编号。
  const autoNumber = useCallback((start: number) => {
    const rows = gridRef.current
    // 末尾数据行：自底向上找第一个非空行（只在该区间内编号）
    let end = -1
    for (let r = rows.length - 1; r >= start; r -= 1) {
      if (!isRowEmpty(rows[r])) {
        end = r
        break
      }
    }
    if (end < start) {
      setStatus('本行以下没有可编号的数据')
      return
    }
    const seedStr = String(rows[start]?.[noIdx] ?? '').trim()
    const seedNum = Number(seedStr)
    const seed = seedStr !== '' && Number.isFinite(seedNum) && !/\D/.test(seedStr) ? seedNum : 1
    pushUndo()
    lastCommitKeyRef.current = null
    setGrid((prev) => {
      const next = prev.map((row) => row.slice())
      for (let r = start; r <= end; r += 1) {
        if (!next[r]) next[r] = emptyRow()
        next[r][noIdx] = String(seed + (r - start))
      }
      return next
    })
    setDirty(true)
    setStatus(`已自动编号 ${end - start + 1} 行（${seed}→${seed + (end - start)}）`)
    focusContainer()
  }, [pushUndo, focusContainer])

  /* ===================== 行操作：插入 / 删除 / 清空 ===================== */
  // 删除若干整行（真正移除，行号会重排）；删完保证底部仍有预备空行可继续录入
  const deleteRows = useCallback((rowsIdx: number[]) => {
    const del = new Set(rowsIdx.filter((i) => i >= 0 && i < gridRef.current.length))
    if (del.size === 0) return
    pushUndo()
    lastCommitKeyRef.current = null
    const next = gridRef.current.filter((_, i) => !del.has(i))
    const target = bottomEntryRow(next)
    while (next.length < target + PREPARED_EMPTY_ROWS) next.push(emptyRow())
    setGrid(next)
    setAutoRows(new Set()) // 行列重排，自动着色行号已失效，重置
    setSelRows(new Set())
    setEditing(false)
    setDirty(true)
    const firstDel = Math.min(...Array.from(del))
    setActive((a) => ({ r: Math.min(Math.max(0, firstDel), next.length - 1), c: a.c }))
    setInputRow(Math.min(target, next.length - 1))
    setStatus(`已删除 ${del.size} 行`)
    containerRef.current?.focus()
  }, [bottomEntryRow, pushUndo])

  // 在指定位置插入空行
  const insertRows = useCallback((at: number, count = 1) => {
    pushUndo()
    lastCommitKeyRef.current = null
    const next = gridRef.current.map((row) => row.slice())
    const pos = Math.max(0, Math.min(at, next.length))
    next.splice(pos, 0, ...Array.from({ length: count }, () => emptyRow()))
    setGrid(next)
    setAutoRows(new Set()) // 插入后行号重排，重置自动着色
    setSelRows(new Set())
    setDirty(true)
    setActive({ r: pos, c: 0 })
    setInputRow(pos)
    setEditing(true)
    setStatus(`已插入 ${count} 行`)
  }, [pushUndo])

  // 只清空内容、保留行位置（Excel 里 Delete 的语义）
  const clearRows = useCallback((rowsIdx: number[]) => {
    const set = new Set(rowsIdx)
    if (set.size === 0) return
    pushUndo()
    lastCommitKeyRef.current = null
    setGrid((prev) => prev.map((row, i) => (set.has(i) ? emptyRow() : row)))
    setDirty(true)
    setStatus(`已清空 ${set.size} 行内容`)
  }, [pushUndo])

  const appendRowsIfNeeded = useCallback((nr: number) => {
    setGrid((prev) => {
      let next = prev
      if (nr >= next.length) {
        next = prev.map((row) => row.slice())
        while (next.length <= nr) next.push(emptyRow())
      }
      return next
    })
  }, [])

  // A3 新行预填：落到一个全新的空行时，从上一行带走 SEED_FIELDS 列的默认值，
  // 暂存进 rowDefaultsRef（不写 grid）。真正录入该行时再由 commitCell 合并一次。
  const maybeSeed = useCallback((nr: number) => {
    if (nr <= 0) return
    if (rowDefaultsRef.current.has(nr)) return
    if (!isRowEmpty(gridRef.current[nr])) return
    rowDefaultsRef.current.set(nr, buildSeedRow(gridRef.current[nr - 1]))
  }, [])

  // 编辑态内导航：Enter=下、Tab=右(末列换行)、Shift+Tab=左、Esc=退出编辑
  const moveEdit = useCallback((dir: 'down' | 'up' | 'right' | 'left') => {
    const { r, c } = activeRef.current
    let nr = r
    let nc = c
    if (dir === 'down') nr = r + 1
    else if (dir === 'up') nr = r - 1
    else if (dir === 'right') {
      nc = c + 1
      if (nc >= COL_COUNT) {
        nc = 0
        nr = r + 1
      }
    } else if (dir === 'left') {
      nc = c - 1
      if (nc < 0) {
        nc = COL_COUNT - 1
        nr = r - 1
      }
    }
    if (nr < 0) nr = 0
    if (nr >= gridRef.current.length) appendRowsIfNeeded(nr)
    maybeSeed(nr)
    setActive({ r: nr, c: nc })
    if (isRowEmpty(gridRef.current[nr])) setInputRow(nr)
    setEditing(true)
  }, [appendRowsIfNeeded, maybeSeed])

  // 非编辑态导航
  const nav = useCallback((dr: number, dc: number) => {
    if (selRowsRef.current.size > 0) setSelRows(new Set()) // 方向键移动即取消整行选择
    if (selRangeRef.current) setSelRange(null) // 方向键也取消矩形框选
    const { r, c } = activeRef.current
    let nr = r + dr
    let nc = c + dc
    if (nr < 0) nr = 0
    if (nc < 0) nc = 0
    if (nc >= COL_COUNT) nc = COL_COUNT - 1
    if (nr >= gridRef.current.length) appendRowsIfNeeded(nr)
    maybeSeed(nr)
    setActive({ r: nr, c: nc })
    if (isRowEmpty(gridRef.current[nr])) setInputRow(nr)
    setEditing(false)
    focusContainer()
  }, [appendRowsIfNeeded, maybeSeed])

  /* ===================== 复制 / 粘贴 ===================== */
  // 把当前"选择"转成 TSV 文本 + 尺寸：优先矩形框选 > 整行选择 > 单格。
  // 供复制/剪切使用（剪切后再清空对应区域）。
  const selBlockTSV = useCallback((): { tsv: string; rows: number; cols: number } => {
    const range = selRangeRef.current
    if (range) {
      const rmin = Math.min(range.r1, range.r2)
      const rmax = Math.max(range.r1, range.r2)
      const cmin = Math.min(range.c1, range.c2)
      const cmax = Math.max(range.c1, range.c2)
      const lines: string[] = []
      for (let r = rmin; r <= rmax; r += 1) {
        const row = gridRef.current[r] || emptyRow()
        const cells: string[] = []
        for (let c = cmin; c <= cmax; c += 1) cells.push(row[c] || '')
        lines.push(cells.join('\t'))
      }
      return { tsv: lines.join('\n'), rows: rmax - rmin + 1, cols: cmax - cmin + 1 }
    }
    const selr = selRowsRef.current
    if (selr.size > 0) {
      const idx = Array.from(selr).sort((a, b) => a - b)
      const tsv = idx.map((i) => (gridRef.current[i] || emptyRow()).join('\t')).join('\n')
      return { tsv, rows: idx.length, cols: COL_COUNT }
    }
    const { r, c } = activeRef.current
    return { tsv: gridRef.current[r]?.[c] || '', rows: 1, cols: 1 }
  }, [])

  // 清空当前"选择"的内容（不删行、不破坏结构，与 Excel Delete 语义一致）：
  // 优先矩形框选 > 整行选择 > 单格。
  const clearSelectionBlock = useCallback(() => {
    const range = selRangeRef.current
    if (range) {
      const rmin = Math.min(range.r1, range.r2)
      const rmax = Math.max(range.r1, range.r2)
      const cmin = Math.min(range.c1, range.c2)
      const cmax = Math.max(range.c1, range.c2)
      pushUndo()
      lastCommitKeyRef.current = null
      setGrid((prev) => {
        const next = prev.map((row) => row.slice())
        for (let r = rmin; r <= rmax; r += 1) {
          for (let c = cmin; c <= cmax; c += 1) {
            if (next[r]) next[r][c] = ''
          }
        }
        return next
      })
      setDirty(true)
      setSelRange(null)
      setStatus(`已清空选区 ${rmax - rmin + 1}×${cmax - cmin + 1}`)
      return
    }
    if (selRowsRef.current.size > 0) {
      clearRows(Array.from(selRowsRef.current))
      return
    }
    const { r, c } = activeRef.current
    clearCell(r, c)
  }, [pushUndo, clearRows, clearCell])

  const pasteTSV = useCallback((text: string, start: { r: number; c: number }) => {
    const cleaned = text.replace(/\r/g, '')
    const lines = cleaned.split('\n')
    while (lines.length && lines[lines.length - 1] === '') lines.pop()
    if (lines.length === 0) return
    const matrix = lines.map((l) => l.split('\t'))
    pushUndo()
    lastCommitKeyRef.current = null
    setGrid((prev) => {
      const next = prev.map((row) => row.slice())
      const needRows = start.r + matrix.length
      while (next.length < needRows) next.push(emptyRow())
      matrix.forEach((rowArr, i) => {
        rowArr.forEach((val, j) => {
          const c = start.c + j
          if (c >= COL_COUNT) return
          // 落到日期列的内容统一规范化成 yyyy-mm-dd（识别不了就原样保留）
          if (COLUMNS[c].field === 'date') {
            const iso = parseDateText(val)
            next[start.r + i][c] = iso || val.trim()
          } else {
            next[start.r + i][c] = val
          }
        })
      })
      return next
    })
    setActive({ r: start.r, c: start.c })
    setEditing(false)
    setDirty(true)
    focusContainer()
  }, [pushUndo])

  // 读系统剪贴板并粘贴到指定格（供右键菜单用；Ctrl+V 走原生 paste 事件更可靠）
  const pasteFromClipboard = useCallback(async (r: number, c: number) => {
    let text = ''
    try {
      text = await navigator.clipboard.readText()
    } catch {
      /* 读不到系统剪贴板时用自己复制过的内容兜底 */
    }
    if (!text) text = lastCopyRef.current
    if (!text) {
      setStatus('剪贴板为空，或请改用 Ctrl+V 粘贴')
      return
    }
    // 若有矩形框选，从选区左上角开始粘贴；否则从右键的格开始
    const range = selRangeRef.current
    const start = range ? { r: Math.min(range.r1, range.r2), c: Math.min(range.c1, range.c2) } : { r, c }
    pasteTSV(text, start)
    if (range) setSelRange(null)
    setStatus('已粘贴')
  }, [pasteTSV])

  /* ===================== A1 记忆式自动补全 ===================== */
  // 接受某个建议值：写入该格；moveDir 指定接受后移动到哪个方向继续录入
  const acceptSuggest = useCallback(
    (r: number, c: number, value: string, moveDir?: 'down' | 'right' | 'left') => {
      commitCell(r, c, value)
      setSuggest(null)
      if (moveDir) moveEdit(moveDir)
    },
    [commitCell, moveEdit],
  )

  // A1 记忆式补全：根据当前文本，收集同列（除本行外）的历史值给出前缀匹配建议；
  // 文本为空时展示最近用过的若干值。建议框定位到编辑器下方（视口坐标）。
  const updateSuggest = useCallback((r: number, c: number, value: string) => {
    if (COLUMNS[c].type !== 'text') {
      setSuggest(null)
      return
    }
    const cur = String(value).trim()
    const seen = new Set<string>()
    const hist: string[] = []
    // 从下往上收集，越近录入的越优先
    for (let i = gridRef.current.length - 1; i >= 0; i -= 1) {
      if (i === r) continue
      const v = String(gridRef.current[i]?.[c] ?? '').trim()
      if (!v || seen.has(v)) continue
      seen.add(v)
      hist.push(v)
    }
    const items = cur === ''
      ? hist.slice(0, 8)
      : hist.filter((h) => h.toLowerCase().startsWith(cur.toLowerCase())).slice(0, 8)
    if (items.length === 0) {
      setSuggest(null)
      return
    }
    const el = editRef.current
    const rect = el ? el.getBoundingClientRect() : null
    setSuggest({
      r,
      c,
      items,
      hi: 0,
      rect: rect ? { left: rect.left, top: rect.bottom, width: rect.width } : null,
    })
  }, [])

  // A1：进入文本列编辑态时，按已输入内容给出同列历史补全建议（换格/进入即刷新）
  useEffect(() => {
    if (!editing) {
      setSuggest(null)
      return
    }
    const { r, c } = active
    if (COLUMNS[c].type !== 'text') {
      setSuggest(null)
      return
    }
    updateSuggest(r, c, gridRef.current[r]?.[c] || '')
  }, [editing, active.r, active.c, updateSuggest])

  /* ===================== A2 复制上一行 ===================== */
  // Ctrl+D：把上一行整行复制到当前行（若已选多行，则各自从自己的上一行复制）
  const fillFromAbove = useCallback(
    (r: number) => {
      const sel = selRowsRef.current
      const targets = sel.size > 0 ? Array.from(sel).sort((a, b) => a - b) : [r]
      if (targets[0] <= 0) {
        setStatus('已在首行，没有上一行可复制')
        return
      }
      const base = gridRef.current
      pushUndo()
      lastCommitKeyRef.current = null
      setGrid((prev) => {
        const next = prev.map((row) => row.slice())
        for (const tr of targets) {
          if (tr <= 0) continue
          const src = base[tr - 1]
          if (src) next[tr] = src.slice()
        }
        return next
      })
      setDirty(true)
      setStatus(`已复制上一行到 ${targets.length} 行`)
      focusContainer()
    },
    [pushUndo],
  )

  // 右键"复制此行到下方"：在选中行正下方各插入一份副本
  const duplicateRowsBelow = useCallback(
    (rowsIdx: number[]) => {
      const sorted = rowsIdx.slice().sort((a, b) => a - b)
      if (sorted.length === 0) return
      pushUndo()
      lastCommitKeyRef.current = null
      setGrid((prev) => {
        const next = prev.map((row) => row.slice())
        // 从下往上插，避免索引错位
        for (let k = sorted.length - 1; k >= 0; k -= 1) {
          const srcIdx = sorted[k]
          const copy = (next[srcIdx] || emptyRow()).slice()
          next.splice(srcIdx + 1, 0, copy)
        }
        return next
      })
      rowDefaultsRef.current.clear()
      setDirty(true)
      const firstNew = sorted[0] + 1
      setActive({ r: firstNew, c: 0 })
      setInputRow(firstNew)
      setEditing(true)
      setStatus(`已复制 ${sorted.length} 行到下方`)
    },
    [pushUndo],
  )

  /* ===================== A5 文件导入 ===================== */
  const handleImportFile = useCallback(
    (e: React.ChangeEvent<HTMLInputElement>) => {
      const f = e.target.files && e.target.files[0]
      if (!f) return
      const reader = new FileReader()
      reader.onload = () => {
        const text = String(reader.result || '')
        const lines = text.replace(/\r/g, '').split('\n')
        while (lines.length && lines[lines.length - 1].trim() === '') lines.pop()
        if (lines.length === 0) {
          setStatus('文件为空')
          return
        }
        // 含制表符按 TSV 解析，否则按 CSV（简单去引号）解析
        const matrix = lines.map((ln) =>
          ln.includes('\t')
            ? ln.split('\t')
            : ln.split(',').map((s) => s.replace(/^"|"$/g, '')),
        )
        const tsv = matrix.map((rr) => rr.join('\t')).join('\n')
        // 追加到现有数据末尾
        pasteTSV(tsv, { r: gridRef.current.length, c: 0 })
        setStatus(`已导入 ${matrix.length} 行（含表头请自行删除）`)
      }
      reader.onerror = () => setStatus('文件读取失败')
      reader.readAsText(f)
      e.target.value = '' // 允许重复导入同一文件
    },
    [pasteTSV],
  )

  /* ===================== 行选择（点行号） ===================== */
  const rowAnchorRef = useRef(0) // Shift 范围选的锚点
  const onRowNumClick = useCallback((r: number, e: React.MouseEvent) => {
    setMenu(null)
    if (e.shiftKey) {
      const from = Math.min(rowAnchorRef.current, r)
      const to = Math.max(rowAnchorRef.current, r)
      const set = new Set<number>()
      for (let i = from; i <= to; i += 1) set.add(i)
      setSelRows(set)
    } else if (e.ctrlKey || e.metaKey) {
      setSelRows((prev) => {
        const set = new Set(prev)
        if (set.has(r)) set.delete(r)
        else set.add(r)
        return set
      })
      rowAnchorRef.current = r
    } else {
      setSelRows(new Set([r]))
      rowAnchorRef.current = r
    }
    setActive({ r, c: 0 })
    setEditing(false)
    containerRef.current?.focus()
  }, [])

  // 在行号上按住左键拖动 → 像 Excel 一样整段选中多行（Ctrl/Shift 交给 onClick 处理）
  const onRowNumMouseDown = useCallback((e: React.MouseEvent, r: number) => {
    if (editingRef.current || e.button !== 0) return
    if (e.ctrlKey || e.metaKey || e.shiftKey) return // 修饰键走 onClick 的多选/连选
    e.stopPropagation()
    const tbody = containerRef.current?.querySelector('tbody')
    const trs = tbody ? (Array.from(tbody.rows) as HTMLTableRowElement[]) : []
    rangeDragRef.current = { r1: r, c1: 0, x: e.clientX, y: e.clientY, moved: false, trs, rowMode: true }

    const onMove = (me: MouseEvent) => {
      const d = rangeDragRef.current
      if (!d) return
      if (!d.moved && Math.hypot(me.clientX - d.x, me.clientY - d.y) < 4) return
      d.moved = true
      const rr = rowIndexAtY(d.trs, me.clientY)
      const a = Math.min(r, rr)
      const b = Math.max(r, rr)
      const set = new Set<number>()
      for (let i = a; i <= b; i += 1) set.add(i)
      setSelRows(set)
    }
    const onUp = () => {
      window.removeEventListener('mousemove', onMove)
      window.removeEventListener('mouseup', onUp)
      const d = rangeDragRef.current
      rangeDragRef.current = null
      if (!d || !d.moved) return // 没拖动 → 交给 onClick 单选整行
      // 拖动选行结束：活动格回到起点行，吞掉尾随 click
      setActive({ r: Math.min(r, d.r1), c: 0 })
      suppressClickRef.current = true
      window.setTimeout(() => {
        suppressClickRef.current = false
      }, 0)
      focusContainer()
    }
    window.addEventListener('mousemove', onMove)
    window.addEventListener('mouseup', onUp)
  }, [focusContainer])

  /* ===================== 右键菜单 ===================== */
  // 菜单作用的行集合：右键处在已选中的行里 → 整批；否则只作用于右键那一行
  const menuRows = useCallback((r: number): number[] => {
    const sel = selRowsRef.current
    if (sel.size > 0 && sel.has(r)) return Array.from(sel).sort((a, b) => a - b)
    return [r]
  }, [])

  const openMenu = useCallback((e: React.MouseEvent, r: number, c: number) => {
    e.preventDefault()
    e.stopPropagation()
    // 右键落在未选中的行 → 清掉旧的行选择（Excel 习惯）
    if (!selRowsRef.current.has(r)) setSelRows(new Set())
    // 右键落在矩形框选区域外 → 清除框选，改为单格菜单（Excel 习惯）
    const sr = selRangeRef.current
    if (sr) {
      const inR = r >= Math.min(sr.r1, sr.r2) && r <= Math.max(sr.r1, sr.r2)
      const inC = c >= Math.min(sr.c1, sr.c2) && c <= Math.max(sr.c1, sr.c2)
      if (!inR || !inC) setSelRange(null)
    }
    setActive({ r, c })
    setEditing(false)
    // 防止菜单溢出窗口
    const MW = 190
    const MH = 288
    const x = Math.min(e.clientX, window.innerWidth - MW - 8)
    const y = Math.min(e.clientY, window.innerHeight - MH - 8)
    setMenu({ x: Math.max(4, x), y: Math.max(4, y), r, c })
  }, [])

  // 点击别处 / 滚动 / Esc 关闭菜单
  useEffect(() => {
    if (!menu) return
    const close = () => setMenu(null)
    const onKey = (ev: KeyboardEvent) => {
      if (ev.key === 'Escape') close()
    }
    window.addEventListener('mousedown', close)
    window.addEventListener('resize', close)
    window.addEventListener('blur', close)
    window.addEventListener('keydown', onKey)
    const scroller = containerRef.current
    scroller?.addEventListener('scroll', close)
    return () => {
      window.removeEventListener('mousedown', close)
      window.removeEventListener('resize', close)
      window.removeEventListener('blur', close)
      window.removeEventListener('keydown', onKey)
      scroller?.removeEventListener('scroll', close)
    }
  }, [menu])

  /* ===================== 键盘（容器层，非编辑态） ===================== */
  const onContainerKey = useCallback((e: React.KeyboardEvent<HTMLDivElement>) => {
    const { r, c } = activeRef.current
    // 编辑态：只放行 Ctrl/Cmd+S 保存与撤销/重做，其余交给单元格内编辑器
    // （单元格是受控组件，浏览器原生的输入框撤销本就失效，统一走网格自己的撤销栈）
    if (editingRef.current) {
      if (e.ctrlKey || e.metaKey) {
        const k = e.key.toLowerCase()
        if (k === 's') {
          e.preventDefault()
          void handleSave()
        } else if (k === 'z') {
          e.preventDefault()
          if (e.shiftKey) redo()
          else undo()
        } else if (k === 'y') {
          e.preventDefault()
          redo()
        } else if (k === 'd') {
          e.preventDefault()
          fillFromAbove(r)
        }
      }
      return
    }
    switch (e.key) {
      case 'ArrowDown': e.preventDefault(); nav(1, 0); break
      case 'ArrowUp': e.preventDefault(); nav(-1, 0); break
      case 'ArrowLeft': e.preventDefault(); nav(0, -1); break
      case 'ArrowRight': e.preventDefault(); nav(0, 1); break
      case 'Tab': e.preventDefault(); nav(0, e.shiftKey ? -1 : 1); break
      case 'Enter':
      case 'F2': e.preventDefault(); startEdit(r, c); break
      case 'Delete':
      case 'Backspace':
        e.preventDefault()
        // 清空当前选择的内容：矩形框选 > 整行选择 > 单格（与 Excel 一致：Delete 只清内容不删行）
        clearSelectionBlock()
        break
      default:
        if (e.ctrlKey || e.metaKey) {
          const k = e.key.toLowerCase()
          if (k === 'x') {
            e.preventDefault()
            // 剪切：把当前选择（矩形框选 > 整行 > 单格）复制到剪贴板并清空
            const info = selBlockTSV()
            copyText(info.tsv)
            clearSelectionBlock()
          } else if (k === 's') {
            e.preventDefault()
            void handleSave()
          } else if (k === 'z') {
            e.preventDefault()
            if (e.shiftKey) redo()
            else undo()
          } else if (k === 'y') {
            e.preventDefault()
            redo()
          } else if (e.key === '-' || e.key === 'Subtract') {
            // Excel 的删除行快捷键 Ctrl+-
            e.preventDefault()
            deleteRows(selRowsRef.current.size > 0 ? Array.from(selRowsRef.current) : [r])
          } else if (k === 'd') {
            // Excel 的"向下填充"：把上一行整行复制到当前行
            e.preventDefault()
            fillFromAbove(r)
          }
          // c/v 由原生 copy/paste 事件处理
          return
        }
        // 可打印字符：直接替换进入编辑
        if (e.key.length === 1 && !e.altKey) {
          e.preventDefault()
          // 有矩形框选时：清空整个选区，然后在选区左上角进入编辑（相当于"键入替换选区"）
          const range = selRangeRef.current
          if (range) {
            const r1 = Math.min(range.r1, range.r2)
            const c1 = Math.min(range.c1, range.c2)
            clearSelectionBlock()
            setActive({ r: r1, c: c1 })
            setEditing(true)
            focusContainer()
          } else {
            startEdit(r, c, e.key)
          }
        }
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [nav, startEdit, clearCell, copyText, undo, redo, clearRows, deleteRows, selBlockTSV, clearSelectionBlock, focusContainer])

  /* ===================== 原生复制 / 粘贴 / 剪切 ===================== */
  const onCopy = (e: React.ClipboardEvent) => {
    if (editingRef.current) return // 编辑态让文本框自己复制
    e.preventDefault()
    const info = selBlockTSV()
    e.clipboardData.setData('text/plain', info.tsv)
    lastCopyRef.current = info.tsv
    if (info.rows > 1 || info.cols > 1) setStatus(`已复制选区 ${info.rows}×${info.cols}`)
    else setStatus('已复制')
  }
  const onCut = (e: React.ClipboardEvent) => {
    if (editingRef.current) return
    e.preventDefault()
    const info = selBlockTSV()
    e.clipboardData.setData('text/plain', info.tsv)
    lastCopyRef.current = info.tsv
    clearSelectionBlock()
  }
  const onPaste = (e: React.ClipboardEvent) => {
    if (editingRef.current) return // 编辑态让文本框自己粘贴
    e.preventDefault()
    const text = e.clipboardData.getData('text')
    if (!text) return
    const range = selRangeRef.current
    const start = range
      ? { r: Math.min(range.r1, range.r2), c: Math.min(range.c1, range.c2) }
      : activeRef.current
    pasteTSV(text, start)
    if (range) setSelRange(null)
  }

  /* ===================== 保存（覆盖更新，剥离尾部空行） ===================== */
  async function handleSave() {
    if (saving) return
    setSaving(true)
    setStatus('')
    try {
      // 先去掉尾部无实质内容的行（全空、或只有"预填默认值"但没录任何有效字段的行）
      const rows = gridRef.current.slice()
      while (rows.length && !rowHasSubstance(rows[rows.length - 1])) {
        rows.pop()
      }
      const res = await api.saveSheet(file.filePath, rows)
      if (res.error) {
        setStatus('保存失败：' + (res.message || '未知错误'))
      } else {
        setDirty(false)
        setStatus(res.message || '已保存')
        onSaved()
      }
    } catch (err) {
      setStatus('保存失败：' + (err instanceof Error ? err.message : '未知错误'))
    } finally {
      setSaving(false)
    }
  }

  /* ===================== 关闭（有改动先确认） ===================== */
  function handleClose() {
    if (dirty && !window.confirm('有未保存的修改，确定关闭吗？')) return
    onClose()
  }

  // 按 ESC 关闭录入面板（与右上角「关闭」按钮 title 提示一致）。
  // 正在编辑单元格（焦点在 input/textarea）时，由单元格编辑器自己处理 ESC（退出编辑），此处不关闭面板；
  // 右键菜单 / 补全建议 / 设置等模态浮层打开时也不抢占它们的 ESC。
  const handleCloseRef = useRef(handleClose)
  useEffect(() => {
    handleCloseRef.current = handleClose
  })
  useEffect(() => {
    const onKey = (e: KeyboardEvent) => {
      if (e.key !== 'Escape') return
      const el = document.activeElement as HTMLElement | null
      if (el && (el.tagName === 'INPUT' || el.tagName === 'TEXTAREA' || el.isContentEditable)) return
      if (document.querySelector('.modal-overlay, .ctx-menu, .suggest-menu')) return
      handleCloseRef.current()
    }
    window.addEventListener('keydown', onKey)
    return () => window.removeEventListener('keydown', onKey)
  }, [])

  /* ===================== 恢复上一个版本（回滚到保存前的备份） ===================== */
  async function handleRestore() {
    if (saving) return
    setStatus('')
    try {
      const res = await api.listBackups(file.filePath)
      if (res.error) {
        setStatus('读取备份失败：' + (res.message || '未知错误'))
        return
      }
      const backups = res.backups || []
      if (backups.length === 0) {
        setStatus('没有可恢复的版本（保存后才会自动生成备份）')
        return
      }
      const t = backups[0].time
      if (!window.confirm(`确定恢复上一个版本吗？\n版本时间：${t}\n（当前内容会先自动备份，可再次恢复）`)) {
        return
      }
      setSaving(true)
      const r = await api.restoreBackup(file.filePath)
      if (r.error) {
        setStatus('恢复失败：' + (r.message || '未知错误'))
        return
      }
      // 重新加载该文件内容到网格
      const res2 = await api.loadSheet(file.filePath)
      if (res2.error) {
        setStatus('恢复成功，但重新读取失败：' + (res2.message || ''))
        return
      }
      const rows = (res2.rows || []).map((rr) => normalizeRow(rr))
      const target = bottomEntryRow(rows)
      while (rows.length < target + PREPARED_EMPTY_ROWS) rows.push(emptyRow())
      setGrid(rows)
      setInputRow(target)
      setActive({ r: target, c: 0 })
      setDirty(false)
      setStatus(r.message || '已恢复上一个版本')
      onSaved()
    } catch (err) {
      setStatus('恢复失败：' + (err instanceof Error ? err.message : '未知错误'))
    } finally {
      setSaving(false)
    }
  }

  /* ===================== AI 识别结果回填 ===================== */
  function applyRecognizedRows(rows: AIRecognizedRow[], append = false) {
    if (!rows.length) return
    pushUndo()
    lastCommitKeyRef.current = null
    const start = append ? bottomEntryRow(gridRef.current) : activeRef.current.r
    const filled: number[] = []
    for (let i = 0; i < rows.length; i += 1) filled.push(start + i)
    setGrid((prev) => {
      const next = prev.map((row) => row.slice())
      let pos = start
      for (const r of rows) {
        if (pos >= next.length) next.push(emptyRow())
        const q = Number(String(r.qty ?? '').replace(/,/g, '')) || 0
        const p = Number(String(r.price ?? '').replace(/,/g, '')) || 0
        const a = Number(String(r.amount ?? '').replace(/,/g, '')) || 0
        next[pos] = [
          String(r.no ?? ''),
          parseDateText(String(r.date ?? '')),
          String(r.name ?? ''),
          String(r.unit ?? ''),
          q > 0 ? fmtNum(q) : '',
          p > 0 ? fmtNum(p) : '',
          a > 0 ? fmtNum(a) : '',
          // 调货人列：AI 识别出的 person 仅用于"账单定位"（匹配并置顶对应 Excel 文件），
          // 不应自动写入调货人列（识别结果表下方也未展示该字段），故此处留空，由用户按需手填。
          '',
          String(r.remark ?? ''),
        ]
        pos += 1
      }
      // 保证底部仍有预备空行
      const target = bottomEntryRow(next)
      while (next.length < target + PREPARED_EMPTY_ROWS) next.push(emptyRow())
      return next
    })
    // 标记这些行为"AI 自动预录"，渲染时高亮区分
    setAutoRows((prev) => {
      const n = new Set(prev)
      filled.forEach((p) => n.add(p))
      return n
    })
    setDirty(true)
    setInputRow(start)
    setActive({ r: start, c: 0 })
    setEditing(true)
    setStatus(`已填入 ${rows.length} 行识别结果（黄色高亮为 AI 预录，修改后自动转白）`)
  }

  /* ===================== 列宽拖拽 ===================== */
  const resizeState = useRef<{ colIdx: number; startX: number; startWidth: number } | null>(null)
  const onStartResize = (colIdx: number, e: React.MouseEvent) => {
    e.stopPropagation()
    e.preventDefault()
    const startX = e.clientX
    const startWidth = colWidths[colIdx]
    resizeState.current = { colIdx, startX, startWidth }
    const onMove = (me: MouseEvent) => {
      if (!resizeState.current) return
      const diff = me.clientX - startX
      const w = Math.max(MIN_COL_WIDTH, startWidth + diff)
      setColWidths((prev) => {
        const next = prev.slice()
        next[colIdx] = w
        return next
      })
    }
    const onUp = () => {
      resizeState.current = null
      window.removeEventListener('mousemove', onMove)
      window.removeEventListener('mouseup', onUp)
      document.body.style.cursor = ''
    }
    window.addEventListener('mousemove', onMove)
    window.addEventListener('mouseup', onUp)
    document.body.style.cursor = 'col-resize'
  }

  /* ===================== 渲染 ===================== */
  const totalAmount = grid.reduce((s, row) => {
    const v = Number(row[amountIdx])
    return s + (isFinite(v) ? v : 0)
  }, 0)
  const entryRow = loaded ? inputRow : -1

  function renderEditor(r: number, c: number, val: string) {
    const col = COLUMNS[c]
    const common = {
      ref: editRef as React.Ref<HTMLTextAreaElement & HTMLInputElement>,
      className: 'cell-editor',
      style: { width: '100%' },
      onKeyDown: (e: React.KeyboardEvent) => {
        if (e.key === 'Enter' && !e.shiftKey) {
          e.preventDefault()
          moveEdit('down')
        } else if (e.key === 'Tab') {
          e.preventDefault()
          moveEdit(e.shiftKey ? 'left' : 'right')
        } else if (e.key === 'Escape') {
          e.preventDefault()
          setEditing(false)
          focusContainer()
        }
      },
      onBlur: () => setEditing(false),
    }
    const rows = Math.min(6, Math.max(1, val.split('\n').length))
    return (
      <textarea
        {...common}
        rows={rows}
        value={val}
        onChange={(e) => {
          commitCell(r, c, e.target.value)
          updateSuggest(r, c, e.target.value)
        }}
        onKeyDown={(e) => {
          // 建议框打开时，方向键/回车/Tab/Esc 优先操作建议，而非移动光标
          if (suggest && suggest.r === r && suggest.c === c && suggest.items.length > 0) {
            if (e.key === 'ArrowDown') {
              e.preventDefault()
              setSuggest({ ...suggest, hi: (suggest.hi + 1) % suggest.items.length })
              return
            }
            if (e.key === 'ArrowUp') {
              e.preventDefault()
              setSuggest({
                ...suggest,
                hi: (suggest.hi - 1 + suggest.items.length) % suggest.items.length,
              })
              return
            }
            if (e.key === 'Enter' && !e.shiftKey) {
              e.preventDefault()
              acceptSuggest(r, c, suggest.items[suggest.hi], 'down')
              return
            }
            if (e.key === 'Tab') {
              e.preventDefault()
              acceptSuggest(r, c, suggest.items[suggest.hi], e.shiftKey ? 'left' : 'right')
              return
            }
            if (e.key === 'Escape') {
              e.preventDefault()
              setSuggest(null)
              return
            }
          }
          common.onKeyDown(e)
        }}
      />
    )
  }

  // 存在多选区域（矩形框选或整行选择）时，不再显示"整行/整列"高亮带，
  // 只保留选区本身与活动格边框——与 Excel 一致，避免误以为整行被选中
  const bandOff =
    (!!selRange && (selRange.r1 !== selRange.r2 || selRange.c1 !== selRange.c2)) ||
    selRows.size > 0

  return (
    <div className="sheet-panel">
      <div className="memo-header">
        <div className="memo-file">
          <span className="memo-file-name">{file.fileName}</span>
          {dirty && <span className="dirty-dot" title="有未保存修改">●</span>}
        </div>
        <div className="memo-header-actions">
          <button className="btn btn-primary" disabled={saving} onClick={handleSave}>
            {saving ? '保存中…' : '保存 Ctrl+S'}
          </button>
          <button className="btn btn-outline" onClick={() => { const n = gridRef.current.length; appendRowsIfNeeded(n); setInputRow(n); setActive({ r: n, c: 0 }); setEditing(true) }}>
            + 行
          </button>
          <button
            className="btn btn-outline"
            disabled={histLen.undo === 0}
            onClick={undo}
            title="撤销上一步（Ctrl+Z）"
          >
            撤销
          </button>
          <button
            className="btn btn-outline"
            onClick={() => fileInputRef.current?.click()}
            title="导入 .csv/.tsv/.txt（追加到末尾，含表头请自行删除）"
          >
            导入
          </button>
          <button
            className="btn btn-outline"
            disabled={selRows.size === 0}
            onClick={() => deleteRows(Array.from(selRows))}
            title="删除选中的行（先点左侧行号选行，或右键菜单删除）"
          >
            删除行{selRows.size > 1 ? `(${selRows.size})` : ''}
          </button>
          <button className="btn btn-outline" onClick={() => api.openFile(file.filePath)} title="用 Excel 程序打开">
            Excel 打开
          </button>
          <button
            className="btn btn-outline"
            disabled={saving}
            onClick={handleRestore}
            title="恢复保存前的上一个版本（每次保存都会自动备份，可多次回滚）"
          >
            恢复上一版本
          </button>
          <button className="btn btn-outline" onClick={handleClose} title="Esc 关闭">
            关闭
          </button>
        </div>
      </div>

      <div className="sheet-body">
        <div
          className="sheet-scroll"
          ref={containerRef}
          tabIndex={0}
          onKeyDown={onContainerKey}
          onCopy={onCopy}
          onCut={onCut}
          onPaste={onPaste}
        >
          <table className="sheet-table">
          <colgroup>
            <col style={{ width: 36 }} />
            {colWidths.map((w, i) => (
              <col key={i} style={{ width: w }} />
            ))}
          </colgroup>
          <thead>
            <tr>
              <th className="rownum-h" />
              {COLUMNS.map((col, i) => (
                <th key={col.field} style={{ width: colWidths[i] }}>
                  <span>{col.label}</span>
                  <div
                    className="col-resize-handle"
                    onMouseDown={(e) => onStartResize(i, e)}
                    title="拖动调整列宽"
                  />
                </th>
              ))}
            </tr>
          </thead>
          <tbody>
            {grid.map((row, r) => (
              <tr
                key={r}
                className={
                  (r === entryRow ? 'entry-row' : '') +
                  (r === active.r && !bandOff ? ' row-active' : '') +
                  (selRows.has(r) ? ' row-selected' : '') +
                  (autoRows.has(r) ? ' row-auto' : '') +
                  (fill && r >= Math.min(fill.startR, fill.endR) && r <= Math.max(fill.startR, fill.endR)
                    ? ' fill-range'
                    : '')
                }
              >
                <td
                  className="rownum"
                  title="点击选中整行（Ctrl 多选 / Shift 连选）；按住拖动可连选多行"
                  onMouseDown={(e) => onRowNumMouseDown(e, r)}
                  onClick={(e) => onRowNumClick(r, e)}
                  onContextMenu={(e) => {
                    if (!selRowsRef.current.has(r)) {
                      const s = new Set([r])
                      selRowsRef.current = s // 同步给 openMenu 立即可见，否则会被它当成"未选中"清掉
                      setSelRows(s)
                      rowAnchorRef.current = r
                    }
                    openMenu(e, r, 0)
                  }}
                >
                  {r + 1}
                </td>
                {row.map((val, c) => {
                  const isActive = active.r === r && active.c === c
                  const isEdit = isActive && editing
                  const sr = selRange
                  const inSel = sr
                    ? r >= Math.min(sr.r1, sr.r2) &&
                      r <= Math.max(sr.r1, sr.r2) &&
                      c >= Math.min(sr.c1, sr.c2) &&
                      c <= Math.max(sr.c1, sr.c2)
                    : false
                  const cls =
                    'cell' +
                    (COLUMNS[c].type === 'number' ? ' cell-num' : '') +
                    (isActive ? ' cell-active' : '') +
                    (inSel ? ' cell-range' : '')
                  return (
                    <td
                      key={c}
                      data-r={r}
                      data-c={c}
                      className={isActive ? 'td-active' : c === active.c && !bandOff ? 'col-active' : ''}
                      ref={isActive ? activeTdRef : undefined}
                      onContextMenu={(e) => openMenu(e, r, c)}
                    >
                      {isEdit ? (
                        renderEditor(r, c, val)
                      ) : (
                        <>
                          <div
                            className={cls}
                            title={val ? val : undefined}
                            onMouseDown={(e) => onCellMouseDown(e, r, c)}
                            onClick={() => handleCellClick(r, c)}
                            onDoubleClick={() => {
                              startEdit(r, c)
                            }}
                          >
                            {val}
                          </div>
                          {loaded && isActive && c === noIdx && (
                            <div
                              className="fill-handle"
                              title="拖动向下填充序号序列（1,2,3…）"
                              onMouseDown={onFillHandleDown}
                            />
                          )}
                        </>
                      )}
                    </td>
                  )
                })}
              </tr>
            ))}
          </tbody>
        </table>
      </div>

    </div>

      {suggest && suggest.rect && (
        <ul
          className="suggest-menu"
          style={{
            position: 'fixed',
            left: suggest.rect.left,
            top: suggest.rect.top,
            width: Math.max(suggest.rect.width, 120),
          }}
        >
          {suggest.items.map((it, i) => (
            <li
              key={it + '|' + i}
              className={'suggest-item' + (i === suggest.hi ? ' suggest-hi' : '')}
              onMouseEnter={() => setSuggest((s) => (s ? { ...s, hi: i } : s))}
              onMouseDown={(e) => {
                e.preventDefault() // 别让文本框失焦
                acceptSuggest(suggest.r, suggest.c, it)
              }}
            >
              {it}
            </li>
          ))}
        </ul>
      )}

      {menu && (() => {
        const rows = menuRows(menu.r)
        const many = rows.length > 1
        const item = (label: string, hint: string, fn: () => void, disabled = false) => (
          <button
            className="ctx-item"
            disabled={disabled}
            onMouseDown={(e) => e.stopPropagation()}
            onClick={() => {
              setMenu(null)
              fn()
            }}
          >
            <span>{label}</span>
            {hint && <span className="ctx-hint">{hint}</span>}
          </button>
        )
        return (
          <div className="ctx-menu" style={{ left: menu.x, top: menu.y }}>
            {item('复制', 'Ctrl+C', () => {
              if (selRange) {
                const info = selBlockTSV()
                copyText(info.tsv)
                setStatus(`已复制选区 ${info.rows}×${info.cols}`)
              } else if (many) {
                const tsv = rows.map((i) => (gridRef.current[i] || emptyRow()).join('\t')).join('\n')
                copyText(tsv)
                setStatus(`已复制 ${rows.length} 行`)
              } else {
                const v = gridRef.current[menu.r]?.[menu.c] || ''
                copyText(v)
                setStatus('已复制：' + v.slice(0, 20))
              }
            })}
            {item('剪切', 'Ctrl+X', () => {
              if (selRange) {
                const info = selBlockTSV()
                copyText(info.tsv)
                clearSelectionBlock()
              } else {
                copyText(gridRef.current[menu.r]?.[menu.c] || '')
                clearCell(menu.r, menu.c)
              }
            })}
            {item('粘贴', 'Ctrl+V', () => void pasteFromClipboard(menu.r, menu.c))}
            <div className="ctx-sep" />
            {item('清空内容', 'Delete', () => {
              if (selRange) clearSelectionBlock()
              else if (many || selRowsRef.current.size > 0) clearRows(rows)
              else clearCell(menu.r, menu.c)
            })}
            <div className="ctx-sep" />
            {item('上方插入行', '', () => insertRows(menu.r, rows.length))}
            {item('下方插入行', '', () => insertRows(menu.r + 1, rows.length))}
            {item(many ? '复制选中行到下方' : '复制此行到下方', 'Ctrl+D', () => duplicateRowsBelow(rows))}
            <div className="ctx-sep" />
            {item('序号自动编号', '从本行向下到末尾', () => autoNumber(menu.r))}
            {item(many ? `删除选中的 ${rows.length} 行` : '删除本行', 'Ctrl+-', () => deleteRows(rows))}
            <div className="ctx-sep" />
            {item('撤销', 'Ctrl+Z', () => undo(), histLen.undo === 0)}
            {item('重做', 'Ctrl+Y', () => redo(), histLen.redo === 0)}
          </div>
        )
      })()}

      <div className="sheet-footer">
        <span className="sheet-meta">共 {grid.length} 行 · 合计金额 <b>{fmtNum(totalAmount)}</b></span>
        <span className="memo-hint">
          点单元格选中，双击或打字即进入编辑 · 在单元格上按住左键拖动可<b>框选</b>一片区域，用于复制 / 粘贴 / 删除 ·
          序号格右下角小方块可向下拖动，自动生成 1,2,3… 序列 · 右键「序号自动编号」可从本行一键编号到底 ·
          日期列直接输入或 Ctrl+V 粘贴，支持 2026/8/11、8-11、20260811 等格式自动转换 ·
          文本列输入时按前缀弹出同列历史补全（↑↓ 选择 · Enter/Tab 确认 · Esc 关闭）·
          <b>Ctrl+D</b> 复制上一行（右键"复制此行到下方"亦可直接加副本）·
          <b>导入</b> 按钮可把 .csv/.tsv/.txt 追加到末尾 ·
          点左侧行号选整行（Ctrl/Shift 多选）· 右键菜单可删除/插入行 ·
          Ctrl+C/V/X 复制粘贴剪切 · Ctrl+Z 撤销 · ↑↓←→ 移动 · Enter 下行 · Tab 右移 · Ctrl+S 保存
        </span>
        {status && (
          <span className={status.includes('失败') ? 'memo-status memo-error' : 'memo-status'}>
            {status}
          </span>
        )}
      </div>

      <input
        ref={fileInputRef}
        type="file"
        accept=".csv,.tsv,.txt"
        style={{ display: 'none' }}
        onChange={handleImportFile}
      />
    </div>
  )
}
