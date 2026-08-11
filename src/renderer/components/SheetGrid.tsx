import { useCallback, useEffect, useRef, useState } from 'react'
import { FileEntry, ElectronAPI } from '../types'

/* 固定 9 列（列数不动，就这么多列），type 决定编辑控件与写回格式 */
interface ColDef {
  field: string
  label: string
  type: 'text' | 'number' | 'date'
  width: number
}
const COLUMNS: ColDef[] = [
  { field: 'no', label: '序号', type: 'text', width: 40 },
  { field: 'date', label: '日期', type: 'date', width: 90 },
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
const dateIdx = COLUMNS.findIndex((c) => c.field === 'date')
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
// 主动展开原生日历：单纯 focus() 不会弹出，必须调用 showPicker()（Chrome 99+）。
// 关键：**不要** focus 这个隐藏 input —— 焦点必须留在旁边的文本框上，
// 否则日历被 Esc 关掉后焦点卡在隐藏 input 上，光标消失且打字无效。
// showPicker 不要求元素持有焦点；只有抛错时才退回"先聚焦再弹"。
function openPicker(el: HTMLInputElement) {
  const withPicker = el as HTMLInputElement & { showPicker?: () => void }
  if (typeof withPicker.showPicker !== 'function') return
  try {
    withPicker.showPicker()
  } catch {
    try {
      el.focus()
      withPicker.showPicker()
    } catch {
      /* showPicker 需要用户手势，失败就放弃弹日历，文本框仍可手输 */
    }
  }
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
  // 是否需要在进入日期编辑态后主动弹出日历：
  // 仅双击日期格 / 点日历按钮时置 true；单击与键盘导航不弹，
  // 因为原生日历弹层会接管键盘，页面收不到 Ctrl+V，会让粘贴失效
  const wantPickerRef = useRef(false)
  // 日期格里那个不可见的 <input type="date">，只用来调 showPicker() 弹原生日历
  const datePickerRef = useRef<HTMLInputElement | null>(null)
  // 焦点在日期格内部转移到隐藏 date input 时，不要把它当成"失焦退出编辑"
  const suppressBlurRef = useRef(false)
  // 正在编辑的日期格：退出编辑时把自由文本统一规范化成 yyyy-mm-dd
  const dateEditRef = useRef<{ r: number; c: number } | null>(null)
  const selRowsRef = useRef(selRows)
  selRowsRef.current = selRows
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
      setInputRow(target)
      setActive({ r: target, c: 0 }) // 光标定位到最底下（数据末尾）第一列
      setLoaded(true)
      setStatus('')
      // 自动进入编辑，打开即可直接打字录入
      setEditing(true)
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
    // 只有双击日期格 / 点日历按钮时才弹原生日历
    if (wantPickerRef.current && datePickerRef.current) {
      wantPickerRef.current = false
      suppressBlurRef.current = true
      openPicker(datePickerRef.current)
      window.setTimeout(() => {
        suppressBlurRef.current = false
      }, 400)
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

  // 退出日期格编辑时，把自由输入/粘贴进来的文本统一规范化成 yyyy-mm-dd
  // （编辑期间保留原文，否则打到一半就被改写，没法正常输入）
  useEffect(() => {
    if (editing) return
    const t = dateEditRef.current
    if (!t) return
    dateEditRef.current = null
    const raw = String(gridRef.current[t.r]?.[t.c] ?? '')
    if (!raw.trim() || /^\d{4}-\d{2}-\d{2}$/.test(raw)) return
    const iso = parseDateText(raw)
    // 静默写入：不再入撤销栈（这一格的输入本身已经入过）
    if (iso && iso !== raw) {
      setGrid((prev) => {
        const next = prev.map((row) => row.slice())
        if (next[t.r]) next[t.r][t.c] = iso
        return next
      })
    }
  }, [editing])

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

  // 单击单元格：
  // · 点其它行"已有内容"的单元格 → 复制该值到系统剪贴板（含日期；不弹选择器、不就地改）
  //   之后到目标空格 Ctrl+V 粘贴即可
  // · 点空单元格 / 录入行本身 → 设为录入行；空日期格弹出选择器，其它列选中待录入
  const handleCellClick = useCallback((r: number, c: number) => {
    const val = String(gridRef.current[r]?.[c] ?? '')
    const ir = inputRowRef.current
    // 点普通单元格即取消整行选择（Excel 习惯），避免 Delete 误清掉之前选的行
    if (selRowsRef.current.size > 0) setSelRows(new Set())
    if (val.trim() !== '' && r !== ir) {
      // 复制到剪贴板（不塞进下方单元格）
      copyText(val)
      setActive({ r, c })
      setEditing(false)
      focusContainer()
      const brief = val.replace(/\s+/g, ' ').trim()
      setStatus('已复制：' + (brief.length > 20 ? brief.slice(0, 20) + '…' : brief))
      return
    }
    // 空单元格 / 录入行本身
    if (isRowEmpty(gridRef.current[r])) setInputRow(r)
    if (COLUMNS[c].type === 'date' && val.trim() === '') {
      // 空日期格：进入文本编辑态（光标 I 形，可直接 Ctrl+V 或打字），
      // 不自动弹日历 —— 原生日历会接管键盘导致粘贴失效；要日历点格内按钮或双击
      startEdit(r, c)
    } else {
      selectCell(r, c)
    }
  }, [copyText, selectCell, startEdit])

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
          if (COLUMNS[c].type === 'date') {
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
    pasteTSV(text, { r, c })
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
        // 选中了整行 → 清空这些行的内容（与 Excel 一致：Delete 只清内容不删行）
        if (selRowsRef.current.size > 0) clearRows(Array.from(selRowsRef.current))
        else clearCell(r, c)
        break
      default:
        if (e.ctrlKey || e.metaKey) {
          const k = e.key.toLowerCase()
          if (k === 'x') {
            e.preventDefault()
            copyText(gridRef.current[r][c] || '')
            clearCell(r, c)
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
          startEdit(r, c, e.key)
        }
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [nav, startEdit, clearCell, copyText, undo, redo, clearRows, deleteRows])

  /* ===================== 原生复制 / 粘贴 / 剪切 ===================== */
  const onCopy = (e: React.ClipboardEvent) => {
    if (editingRef.current) return // 编辑态让文本框自己复制
    e.preventDefault()
    const { r, c } = activeRef.current
    // 选中了整行 → 整行按 TSV 复制（可直接粘回 Excel 或本网格其它位置）
    if (selRowsRef.current.size > 0) {
      const idx = Array.from(selRowsRef.current).sort((a, b) => a - b)
      const tsv = idx.map((i) => (gridRef.current[i] || emptyRow()).join('\t')).join('\n')
      e.clipboardData.setData('text/plain', tsv)
      lastCopyRef.current = tsv
      setStatus(`已复制 ${idx.length} 行`)
      return
    }
    const val = gridRef.current[r]?.[c] || ''
    e.clipboardData.setData('text/plain', val)
    lastCopyRef.current = val
  }
  const onCut = (e: React.ClipboardEvent) => {
    if (editingRef.current) return
    e.preventDefault()
    const { r, c } = activeRef.current
    copyText(gridRef.current[r][c] || '')
    clearCell(r, c)
  }
  const onPaste = (e: React.ClipboardEvent) => {
    const { r, c } = activeRef.current
    if (editingRef.current) {
      // 日期编辑态：只取第一段（避免从 Excel 复制的整块塞进一格），并尽量规范化成 ISO；
      // 识别不了就原样填入，让用户看到并手改，而不是静默丢弃
      if (COLUMNS[c].type === 'date') {
        e.preventDefault()
        const text = e.clipboardData.getData('text')
        const first = (text.split(/[\t\n]/)[0] || '').trim()
        if (!first) return
        const iso = parseDateText(first)
        commitCell(r, c, iso || first)
        setStatus(iso ? '已粘贴日期 ' + iso : '无法识别的日期：' + first.slice(0, 20))
        return
      }
      return // 其它列编辑态让文本框自己粘贴
    }
    e.preventDefault()
    const text = e.clipboardData.getData('text')
    if (text) pasteTSV(text, { r, c })
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
      onBlur: () => {
        // 日期格内部把焦点让给隐藏的 date input 时不算失焦，否则日历刚弹出就被关掉
        if (suppressBlurRef.current) return
        setEditing(false)
      },
    }
    if (col.type === 'date') {
      // 编辑器用 text 而非 date：光标是 I 形、Ctrl+V 原生可用、也不会有"年/月/日"占位。
      // 日历改由右侧按钮里的隐藏 date input 弹出（原生日历一旦弹出就会接管键盘，
      // 页面收不到 Ctrl+V，所以不能让它默认占着编辑态）。
      dateEditRef.current = { r, c }
      const iso = /^\d{4}-\d{2}-\d{2}$/.test(val) ? val : ''
      const openCalendar = () => {
        if (!datePickerRef.current) return
        suppressBlurRef.current = true
        openPicker(datePickerRef.current)
        window.setTimeout(() => {
          suppressBlurRef.current = false
        }, 400)
      }
      return (
        <div className="date-edit" style={{ width: '100%' }}>
          <input
            {...common}
            type="text"
            className="cell-editor cell-editor-datetext"
            style={{ flex: 1, minWidth: 0 }}
            value={val}
            onChange={(e) => commitCell(r, c, e.target.value)}
            onKeyDown={(e) => {
              // Alt+↓ / F4：键盘也能唤出日历（Excel 下拉习惯）
              if ((e.altKey && e.key === 'ArrowDown') || e.key === 'F4') {
                e.preventDefault()
                openCalendar()
                return
              }
              common.onKeyDown(e)
            }}
          />
          {/* 只用于弹原生日历；视觉上藏在按钮下面 */}
          <input
            ref={datePickerRef}
            type="date"
            className="date-pick-native"
            tabIndex={-1}
            value={iso}
            onChange={(e) => {
              const v = e.target.value
              if (!v) return
              commitCell(r, c, v)
              // 正常情况下焦点从未离开文本框（openPicker 不抢焦点）；
              // 只有走了 focus 兜底路径时才需要还回去，便于继续 Tab 往右录入
              if (document.activeElement !== editRef.current) {
                suppressBlurRef.current = true
                window.setTimeout(() => {
                  editRef.current?.focus()
                  suppressBlurRef.current = false
                }, 0)
              }
            }}
          />
          <button
            type="button"
            className="date-pick-btn"
            tabIndex={-1}
            title="打开日历（Alt+↓）"
            onMouseDown={(e) => {
              e.preventDefault() // 别让文本框失焦
              openCalendar()
            }}
          >
            <svg viewBox="0 0 16 16" width="12" height="12" aria-hidden="true">
              <path
                fill="currentColor"
                d="M5 1v2H3.5A1.5 1.5 0 0 0 2 4.5v9A1.5 1.5 0 0 0 3.5 15h9a1.5 1.5 0 0 0 1.5-1.5v-9A1.5 1.5 0 0 0 12.5 3H11V1H9.5v2h-3V1H5Zm7.5 5.5v7h-9v-7h9Z"
              />
            </svg>
          </button>
        </div>
      )
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
          <button className="btn btn-small btn-close" onClick={handleClose} title="Esc 关闭">
            关闭
          </button>
        </div>
      </div>

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
                  (r === active.r ? ' row-active' : '') +
                  (selRows.has(r) ? ' row-selected' : '')
                }
              >
                <td
                  className="rownum"
                  title="点击选中整行（Ctrl 多选 / Shift 连选），右键可删除"
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
                  // 已有内容且不在录入行 → 单击即复制到剪贴板（给出复制光标提示）
                  const isCopySrc = String(val).trim() !== '' && r !== inputRow
                  const cls =
                    'cell' +
                    (COLUMNS[c].type === 'number' ? ' cell-num' : '') +
                    (isActive ? ' cell-active' : '') +
                    (isCopySrc ? ' cell-copy-src' : '')
                  return (
                    <td
                      key={c}
                      className={isActive ? 'td-active' : c === active.c ? 'col-active' : ''}
                      ref={isActive ? activeTdRef : undefined}
                      onContextMenu={(e) => openMenu(e, r, c)}
                    >
                      {isEdit ? (
                        renderEditor(r, c, val)
                      ) : (
                        <div
                          className={cls}
                          title={isCopySrc ? '点一下复制到剪贴板，再到目标格 Ctrl+V（双击可就地修改）' : (val ? val : undefined)}
                          onClick={() => handleCellClick(r, c)}
                          onDoubleClick={() => {
                            // 双击就地编辑；日期列同时弹出日历，便于改旧日期
                            if (COLUMNS[c].type === 'date') wantPickerRef.current = true
                            startEdit(r, c)
                          }}
                        >
                          {val}
                        </div>
                      )}
                    </td>
                  )
                })}
              </tr>
            ))}
          </tbody>
        </table>
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
              if (many) {
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
              copyText(gridRef.current[menu.r]?.[menu.c] || '')
              clearCell(menu.r, menu.c)
            })}
            {item('粘贴', 'Ctrl+V', () => void pasteFromClipboard(menu.r, menu.c))}
            <div className="ctx-sep" />
            {item('清空内容', 'Delete', () => {
              if (many || selRowsRef.current.size > 0) clearRows(rows)
              else clearCell(menu.r, menu.c)
            })}
            <div className="ctx-sep" />
            {item('上方插入行', '', () => insertRows(menu.r, rows.length))}
            {item('下方插入行', '', () => insertRows(menu.r + 1, rows.length))}
            {item(many ? '复制选中行到下方' : '复制此行到下方', 'Ctrl+D', () => duplicateRowsBelow(rows))}
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
          点已有内容的格 → 复制到剪贴板，到目标格 Ctrl+V 粘贴 ·
          日期格点一下即可 Ctrl+V 粘贴或直接打字（2026/8/11、8-11 等格式自动转换），
          要日历点格内 <b>日历按钮</b> 或双击（Alt+↓）·
          文本列输入时按前缀弹出同列历史补全（↑↓ 选择 · Enter/Tab 确认 · Esc 关闭）·
          <b>Ctrl+D</b> 复制上一行（右键"复制此行到下方"亦可直接加副本）·
          <b>导入</b> 按钮可把 .csv/.tsv/.txt 追加到末尾 ·
          点左侧行号选整行（Ctrl/Shift 多选）· 右键菜单可删除/插入行 ·
          Ctrl+Z 撤销 · ↑↓←→ 移动 · Enter 下行 · Tab 右移 · Ctrl+S 保存
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
