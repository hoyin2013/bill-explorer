import { useEffect, useRef, useState } from 'react'
import { createUniver, LocaleType } from '@univerjs/presets'
import { UniverSheetsCorePreset } from '@univerjs/preset-sheets-core'
import UniverPresetSheetsCoreZhCN from '@univerjs/preset-sheets-core/locales/zh-CN'
import '@univerjs/preset-sheets-core/lib/index.css'
import type { ElectronAPI, FileEntry, AIRecognizedRow } from '../types'
import {
  rowsToWorkbookData,
  workbookDataToRows,
  mapRecognizedToRow,
  buildAutoAmountRows,
  classifyAmount,
  recomputeAmount,
  onAmountChanged,
  dateStrToSerial,
  COL_COUNT,
  BASE_COL_WIDTHS,
  HEADER,
} from '../univerAdapter'

type UniverAPI = ReturnType<typeof createUniver>['univerAPI']

interface Props {
  file: FileEntry
  api: ElectronAPI
  onClose: () => void
  onSaved: () => void
}

// 设置激活单元格，并按需确保视图滚动到该单元格。
// 根因：Univer 在 createWorkbook 之后骨架（render）尚未创建完成，此时任何滚动命令都找不到
// 渲染层而静默/抛错失败（光标数据层已跳到目标行，视图却停在原地）。
// 所以用 `sheet.command.scroll-view` 显式滚动（按行索引，走 SheetScrollManagerService，
// 不依赖缺失的 SheetsScrollRenderController），并做重试：render 一就绪即滚动成功。
// 打开文件时 scroll=true 并带 leadRows，使「末尾多留几行可见」；
// OCR 填入时 scroll=false，保持当前视窗不动（数据落在光标处，不跳走）。
const END_VISIBLE_LEAD = 8

interface ScrollOpts {
  leadRows?: number
  scroll?: boolean
}

// 结构性变更命令关键字（删/插行、列、区域移动等）：这类命令改的是表结构，
// 不一定触发 SheetValueChanged，但会改变数据。若不标记“有修改”，关闭时就不会自动保存
// （典型表现：删了行，关闭重开内容还在）。故在 CommandExecuted 里用子串匹配捕获，
// 以覆盖 remove-row / remove-row-by-range 等变体命令 id。
const SHEET_STRUCT_MUTATION_KEYWORDS = [
  'remove-row',
  'remove-col',
  'insert-row',
  'insert-col',
  'delete-range-move',
  'insert-range-move',
  'move-range',
]

// 列宽等比自适应时，从容器宽度中扣除的左侧行号表头 + 右侧滚动条占位的估算值（px）。
// 略大于实际占位，宁可右侧留一点缝隙，也不让列总宽超过可视区导致出现横向滚动条。
const COL_WIDTH_GUTTER = 64

// 连续自动保存：任何改动（单元格编辑 / 删插行 / OCR 填入）后，延时 AUTOSAVE_DELAY 毫秒
// （防抖，用户停顿后再写盘）自动整表存盘，无需手动保存或关闭文件。
// 这样边录边存，输完一张即可继续下一张，且大幅降低意外丢失风险。
const AUTOSAVE_DELAY = 1000

// 单元格输入记忆（自动补全）：编辑单元格时，根据「本列已有值 + 历史记录」给出提示，辅助快速输入。
// 行为类似 Excel 的按列自动完成：你打字时，下方浮现本列曾出现过的、以及你以前输入过的相近内容。
// - AC_MAX_SUGGESTIONS：弹出候选最多条数
// - AC_HISTORY_CAP：每个「文件+列」持久化保留的历史条数上限（localStorage，跨会话记忆）
const AC_MAX_SUGGESTIONS = 12
const AC_HISTORY_CAP = 200
const AC_HISTORY_PREFIX = 'bill-ac-history::'
// 临时调试开关：开启后左下角 HUD 实时显示自动补全的触发诊断（编辑器选择器/读到的文本/激活格/本列值数/候选数）。
// 用于定位“完全不提示”的环境问题；确认修复后改为 false。
const AC_DEBUG = true
// 日期列（第 2 列，0 基索引 1）：候选需把 Excel 序列号格式化为可读日期展示；
// 选中时再写回序列号，确保单元格仍是“真日期”（带日期样式），而不是变成文本。
const DATE_COL_INDEX = 1
// Excel 序列号 → yyyy-mm-dd（纪元 1899-12-30，与 ExcelJS dateToSerial 一致；用 UTC 取值避免时区偏移）
const excelSerialToDateStr = (serial: number): string => {
  const ms = (serial - 25569) * 86400 * 1000
  const d = new Date(ms)
  const y = d.getUTCFullYear()
  const m = String(d.getUTCMonth() + 1).padStart(2, '0')
  const day = String(d.getUTCDate()).padStart(2, '0')
  return `${y}-${m}-${day}`
}
// 自动补全候选：display=展示文本（日期列已格式化为日期），raw=写入单元格的真实值
// （日期列=Excel 序列号，保证仍为真日期；其余列=原文本）。
type AcItem = { display: string; raw: string | number }

function setActiveAndScroll(
  univerAPI: UniverAPI,
  row: number,
  col: number,
  opts: ScrollOpts = {},
) {
  const { leadRows = 0, scroll = true } = opts
  const active = univerAPI.getActiveSheet()
  if (!active) return
  const ws = active.worksheet
  try {
    ws.setActiveRange(ws.getRange(row, col))
  } catch {
    /* 激活失败不影响其余功能 */
  }
  if (!scroll) return
  const api = univerAPI as unknown as {
    executeCommand: (id: string, params?: object) => Promise<unknown> | unknown
  }
  const scrollRow = Math.max(1, row - leadRows)
  const params = { sheetViewStartRow: scrollRow, sheetViewStartColumn: col, offsetX: 0, offsetY: 0 }
  const tryScroll = (attempt: number) => {
    try {
      Promise.resolve(api.executeCommand('sheet.command.scroll-view', params)).then((ok) => {
        if (!ok && attempt < 12) setTimeout(() => tryScroll(attempt + 1), 100)
      }).catch(() => {
        if (attempt < 12) setTimeout(() => tryScroll(attempt + 1), 100)
      })
    } catch {
      if (attempt < 12) setTimeout(() => tryScroll(attempt + 1), 100)
    }
  }
  if (typeof requestAnimationFrame === 'function') requestAnimationFrame(() => tryScroll(0))
  else tryScroll(0)
}

// 金额列 = 数量(列4,Excel=E)×单价(列5,Excel=F)。
// 行为（见 univerAdapter.ts 的 recomputeAmount / classifyAmount / buildAutoAmountRows）：
// - 自动金额行（金额==数量×单价，或金额待填）→ 数量/单价变化时金额同步重算。
// - 手工金额行（金额≠数量×单价，如整单总价/运费/折扣）→ 不被自动覆盖；
//   改数量/单价时仍保留用户手填的金额。
// - 数量或单价缺失 → 不动金额，留待用户手填。
// 写的是静态数值（而非 =E*F 活公式）：不依赖 Univer 公式引擎是否计算，避免“公式未算而显示 0”。
// `autoAmountRows` 记录哪些行是“自动金额行”，加载文件 / OCR 填入时初始化，
// 编辑金额列时被重分类（手填≠乘积 → 移出集合；填成==乘积 → 重新纳入）。

export function UniverSheet({ file, api, onClose, onSaved }: Props) {
  const containerRef = useRef<HTMLDivElement>(null)
  const univerRef = useRef<UniverAPI | null>(null)
  const dirtyRef = useRef(false)
  const fileRef = useRef(file)
  fileRef.current = file
  // 记录“自动金额行”（金额==数量×单价或待填）。手工金额行（≠乘积）不在此集合，
  // 加载 / OCR 填入时初始化，编辑金额列时按值重分类。
  const autoAmountRows = useRef<Set<number>>(new Set())

  const [status, setStatus] = useState('')
  const [saving, setSaving] = useState(false)
  const [dirty, setDirty] = useState(false)
  const [reloadToken, setReloadToken] = useState(0)
  // 窗口 resize 防抖用的 requestAnimationFrame 句柄
  const resizeRafRef = useRef<number | null>(null)
  // 连续自动保存相关句柄/锁：
  // - autoSaveTimerRef：防抖定时器，改动后延时 AUTOSAVE_DELAY 再写盘
  // - savingRef：是否正在保存（用 ref 而非 state，避免并发闭包读到旧值）
  // - pendingSaveRef：保存中又来了新改动，则保存完再补一次
  const autoSaveTimerRef = useRef<ReturnType<typeof setTimeout> | null>(null)
  const savingRef = useRef(false)
  const pendingSaveRef = useRef(false)
  // 单元格输入记忆（自动补全）相关：
  // - acEditorRef：当前正在编辑的单元格编辑器 DOM（contentEditable[data-u-comp='editor']）
  // - acStateRef：弹出状态机的可变数据（避免打字时频繁 setState 引起重渲染）；items 为候选
  // - acAcceptingRef：正在“选中并填充”的过程中，避免程序化 input 事件被自身监听器误处理
  // - acUi：仅供渲染弹层的快照（open/items/index/定位）
  const acEditorRef = useRef<HTMLElement | null>(null)
  const acAcceptingRef = useRef(false)
  const acRafRef = useRef<number | null>(null)
  const acGhostRef = useRef<HTMLDivElement | null>(null)
  const acMeasureRef = useRef<HTMLSpanElement | null>(null)
  const acDebugRef = useRef<HTMLDivElement | null>(null)
  const acStateRef = useRef<{
    open: boolean
    items: AcItem[]
    index: number
    col: number
    row: number
    partial: string
  }>({ open: false, items: [], index: 0, col: -1, row: -1, partial: '' })
  // 轮询去重键：col:row:text，仅在“文本或所在单元格变化”时才重算候选，避免每帧重复计算
  const acLastKeyRef = useRef('')
  // 编辑时 Univer 的 getActiveRange() 经常返回 null（编辑器接管了选区），
  // 故在非编辑态时记住“最近一次激活单元格”，编辑态拿不到 range 就回退用它，确保仍知道在哪列输入。
  const acLastCellRef = useRef<{ col: number; row: number } | null>(null)
  const [acUi, setAcUi] = useState<{
    open: boolean
    items: AcItem[]
    index: number
    pos: { left: number; top: number } | null
  }>({ open: false, items: [], index: 0, pos: null })

  const markDirty = () => {
    if (!dirtyRef.current) {
      dirtyRef.current = true
      setDirty(true)
    }
  }

  // 列宽等比自适应：按容器可用宽度缩放各列，使编辑区在最大化/缩放窗口时铺满，不留右侧空白。
  // 不改变单元格内容，也不写入 xlsx（saveSheet 只回写值、不回写列宽）。
  const applyProportionalColumnWidths = () => {
    const univerAPI = univerRef.current
    const container = containerRef.current
    if (!univerAPI || !container) return
    const fws = univerAPI.getActiveSheet()?.worksheet
    if (!fws) return
    const avail = container.clientWidth - COL_WIDTH_GUTTER
    if (avail <= 0) return
    const sumBase = BASE_COL_WIDTHS.reduce((a, b) => a + b, 0) || 1
    let scale = avail / sumBase
    // 限制缩放范围，避免极小/极大窗口下单元格过窄或过宽
    scale = Math.max(0.5, Math.min(scale, 2.5))
    BASE_COL_WIDTHS.forEach((w, i) => {
      try {
        fws.setColumnWidth(i, Math.max(24, Math.round(w * scale)))
      } catch {
        /* 单行失败不影响其他 */
      }
    })
  }

  // 保存：取回 Univer 数据 → 还原为 string[][] → 交给 ExcelJS 写回（保留全部既有约定）
  // auto=true 表示由“连续自动保存”触发，状态文案与手动保存区分，并支持保存中重叠保护。
  const handleSave = async (auto = false) => {
    // 正在保存中：标记待补一次，立即返回（避免并发写盘）
    if (savingRef.current) {
      pendingSaveRef.current = true
      return
    }
    savingRef.current = true
    setSaving(true)
    setStatus(auto ? '自动保存中…' : '')
    try {
      const univerAPI = univerRef.current
      if (!univerAPI) throw new Error('未获取到工作簿')
      const wb = univerAPI.getActiveWorkbook()?.getSnapshot()
      if (!wb) throw new Error('未获取到工作簿')
      let rows = workbookDataToRows(wb)
      // 去掉尾部全空行（与旧逻辑一致）
      while (rows.length && rows[rows.length - 1].every((c) => !String(c).trim())) rows.pop()
      const res = await api.saveSheet(fileRef.current.filePath, rows)
      if (res.error) {
        setStatus('保存失败：' + (res.message || '未知错误'))
      } else {
        dirtyRef.current = false
        setDirty(false)
        setStatus(auto ? '已自动保存' : (res.message || '已保存'))
        onSaved()
      }
    } catch (err) {
      setStatus('保存失败：' + (err instanceof Error ? err.message : '未知错误'))
    } finally {
      savingRef.current = false
      setSaving(false)
      // 保存期间又发生了改动 → 再补一次存盘，确保最终一致
      if (pendingSaveRef.current) {
        pendingSaveRef.current = false
        void handleSave(auto)
      }
    }
  }
  const saveRef = useRef(handleSave)
  saveRef.current = handleSave

  // 连续自动保存（防抖）：任意改动后，等 AUTOSAVE_DELAY 毫秒用户停手再写盘。
  // 这样既“边录边存”，又不会每次按键都触发一次全表存盘（大文件代价高）。
  const scheduleAutoSave = () => {
    if (autoSaveTimerRef.current) clearTimeout(autoSaveTimerRef.current)
    autoSaveTimerRef.current = setTimeout(() => {
      autoSaveTimerRef.current = null
      void saveRef.current(true)
    }, AUTOSAVE_DELAY)
  }

  // ===== 单元格输入记忆（自动补全） =====
  // 历史存于 localStorage：键 = 文件+列，值为该列曾输入过的去重字符串数组（跨会话保留）。
  const historyKeyOf = (col: number) => `${AC_HISTORY_PREFIX}${fileRef.current.filePath}::${col}`
  const loadHistory = (key: string): string[] => {
    try {
      const s = localStorage.getItem(key)
      const arr = s ? JSON.parse(s) : []
      return Array.isArray(arr) ? arr.filter((x) => typeof x === 'string') : []
    } catch {
      return []
    }
  }
  // 把某次输入值追加进该列历史（去重 + 截断到上限）
  const pushHistory = (col: number, value: string) => {
    const v = (value ?? '').toString().trim()
    if (!v) return
    const key = historyKeyOf(col)
    const set = new Set(loadHistory(key))
    set.add(v)
    const next = [...set].slice(-AC_HISTORY_CAP)
    try {
      localStorage.setItem(key, JSON.stringify(next))
    } catch {
      /* 配额/隐私模式下静默忽略 */
    }
  }
  // 读取当前工作表某一列里出现过的所有（去重、非空）值，附带来源行号与出现次数，
  // 作为“本列曾经内容”的候选源。行号用于「就近（上方最近）优先」排序；次数用于频率排序。
  // 值保留原始类型（日期列存的是 Excel 序列号 → number，其它列是 string），便于候选格式化。
  const getColumnValuesWithRows = (col: number): Array<{ value: string | number; row: number; freq: number }> => {
    const univerAPI = univerRef.current
    const wb = univerAPI?.getActiveWorkbook()?.getSnapshot()
    const sheets = wb?.sheets as Record<string, { cellData?: Record<number, Record<number, { v?: unknown }>> }> | undefined
    const sheet = sheets ? (sheets['bill'] ?? Object.values(sheets)[0]) : undefined
    const cd = sheet?.cellData
    if (!cd) return []
    const map = new Map<string, { row: number; freq: number }>()
    for (const r of Object.keys(cd)) {
      const cell = cd[Number(r)]?.[col]
      if (cell && cell.v != null) {
        const s = String(cell.v).trim()
        if (s) {
          const prev = map.get(s)
          if (prev) {
            prev.freq += 1
            prev.row = Math.max(prev.row, Number(r))
          } else {
            map.set(s, { row: Number(r), freq: 1 })
          }
        }
      }
    }
    return [...map.entries()].map(([value, info]) => {
      // 注意：cellNumOf 对文本返回 null，而 Number(null) === 0（非 null），
      // 故不能直接 Number(cellNumOf(value))，否则所有文本值都会被错归一成 0，
      // 导致候选匹配全部失败（表现为“什么都不提示”）。只在确为数字时转 number。
      const num = cellNumOf(value)
      return { value: num != null ? num : value, ...info }
    })
  }
  // value 可能是 number 或 string；这里只做“字符串→数字”的轻量归一（用于日期列序列号识别）。
  const cellNumOf = (v: string | number): number | null => {
    if (typeof v === 'number') return v
    const t = (v as string).trim()
    if (/^-?\d+(\.\d+)?$/.test(t)) return Number(t)
    return null
  }
  // 综合「本列已有值（仅当前行上方，与 Excel 自动完成一致）+ 历史记录」给出候选。
  // 匹配规则对齐 Excel/WPS 自动完成：优先「以已输入内容开头」的值（大小写不敏感）；
  // 仅当没有任何开头匹配时，才退化为「包含匹配」（便于输入中间片段也能提示）。
  // 排序：① 上方最近优先（行号大者在前，与 Excel“取上方最近的匹配单元格”一致）；
  //      ② 频率高优先（同一值出现越多 → 越可能是常输项）；③ 字母序兜底。
  // 日期列（DATE_COL_INDEX）：把序列号格式化为 yyyy-mm-dd 展示，raw 仍为序列号（写回即真日期）。
  const computeSuggestions = (partial: string, col: number, currentRow: number): AcItem[] => {
    const p = partial.trim().toLowerCase()
    if (!p) return []
    const isDate = col === DATE_COL_INDEX
    // 把原始值转成候选：display=展示文本，raw=真实写入值。
    // 日期列（DATE_COL_INDEX）：无论来源是序列号（载入/手敲后 Univer 自动转的）还是文本
    // （识图/OCR 填入的 yyyy-mm-dd、手动敲的文本日期），一律解析成**真日期序列号**作为 raw，
    // 这样选中落盘后单元格始终是真日期，不会混入“像日期的文本”；候选展示统一为 yyyy-mm-dd。
    const toItem = (raw: string | number): AcItem => {
      if (isDate) {
        let num: number | null = null
        if (typeof raw === 'number') num = raw
        else {
          const asNum = cellNumOf(raw)
          if (asNum != null) num = asNum
          else num = dateStrToSerial(String(raw)) // 文本日期（2026-08-16 / 2026/8/16 等）
        }
        if (num != null && isFinite(num)) {
          return { display: excelSerialToDateStr(num), raw: num }
        }
        // 极端兜底：解析不出就按原文（理论不会发生）
        return { display: String(raw), raw: String(raw) }
      }
      return { display: String(raw), raw: raw }
    }
    // 本列“上方”已有值（row < 当前行，与 Excel 自动完成范围一致）
    const colVals = getColumnValuesWithRows(col).filter((x) => x.row < currentRow)
    const hist = loadHistory(historyKeyOf(col))
    const cands = new Map<string, AcItem & { freq: number; row: number }>()
    for (const c of colVals) {
      const it = toItem(c.value)
      const prev = cands.get(it.display)
      if (prev) {
        prev.freq += c.freq
        prev.row = Math.max(prev.row, c.row)
      } else {
        cands.set(it.display, { ...it, freq: c.freq, row: c.row })
      }
    }
    // 历史（你以前输入过的）轻微加权，使常输项更容易浮现
    for (const h of new Set(hist)) {
      const it = toItem(h)
      const prev = cands.get(it.display)
      if (prev) prev.freq += 2
      else cands.set(it.display, { ...it, freq: 2, row: -1 })
    }
    const all = [...cands.values()]
    let matched = all.filter((x) => x.display.toLowerCase().startsWith(p))
    if (!matched.length) matched = all.filter((x) => x.display.toLowerCase().includes(p))
    matched.sort((a, b) => {
      const ra = a.row < 0 ? -Infinity : a.row
      const rb = b.row < 0 ? -Infinity : b.row
      if (ra !== rb) return rb - ra // 上方最近优先
      if (a.freq !== b.freq) return b.freq - a.freq // 频率高优先
      return a.display.localeCompare(b.display)
    })
    return matched.slice(0, AC_MAX_SUGGESTIONS)
  }
  // 根据编辑框位置计算弹层坐标：默认在编辑框下方；若下方空间不足（单元格靠近窗口底部），
  // 则翻到编辑框上方，避免弹层被裁掉导致「内容不全 / 只能点到前面几条」。
  const computeAcPos = (editorRect: DOMRect, itemsCount: number): { left: number; top: number } => {
    const estH = Math.min(220, itemsCount * 29 + 10)
    const vh = window.innerHeight || 600
    const belowTop = editorRect.bottom + 2
    const aboveTop = editorRect.top - estH - 2
    let top = belowTop
    if (belowTop + estH > vh - 8 && aboveTop > 8) top = aboveTop
    top = Math.max(8, Math.min(top, vh - estH - 8))
    return { left: editorRect.left, top }
  }
  const showAc = (items: AcItem[], rect: DOMRect) => {
    const pos = computeAcPos(rect, items.length)
    acStateRef.current = { ...acStateRef.current, open: true, items, index: 0 }
    setAcUi({ open: true, items, index: 0, pos })
  }
  const hideAc = () => {
    if (!acStateRef.current.open && !acUi.open) return
    acStateRef.current = { ...acStateRef.current, open: false, items: [], index: 0 }
    setAcUi({ open: false, items: [], index: 0, pos: null })
    // 同时收起“框内灰色续接”提示
    if (acGhostRef.current) acGhostRef.current.style.display = 'none'
  }
  // 渲染 Excel 式“框内灰色续接”：在编辑器内已输入文本之后，用灰色叠加显示
  // 最佳匹配（以已输入内容开头）的剩余后缀，作为视觉提示。纯展示，不改变实际输入；
  // 接受动作仍由下拉候选（Enter/Tab/点击）完成，故不破坏现有下拉行为。
  const updateGhost = (partial: string, items: AcItem[], editor: HTMLElement) => {
    const ghost = acGhostRef.current
    if (!ghost) return
    const p = partial.trim().toLowerCase()
    const match = items.find((it) => it.display.toLowerCase().startsWith(p))
    if (!match || !p) {
      ghost.style.display = 'none'
      return
    }
    const suffix = match.display.slice(p.length)
    if (!suffix) {
      ghost.style.display = 'none'
      return
    }
    const cs = getComputedStyle(editor)
    // 用与编辑器同字体的隐藏量测 span，测“已输入文本”像素宽度，定位灰色后缀起点
    let measure = acMeasureRef.current
    if (!measure) {
      measure = document.createElement('span')
      measure.style.position = 'absolute'
      measure.style.visibility = 'hidden'
      measure.style.whiteSpace = 'nowrap'
      measure.style.top = '-9999px'
      measure.style.left = '-9999px'
      document.body.appendChild(measure)
      acMeasureRef.current = measure
    }
    measure.style.font = cs.font
    measure.style.fontSize = cs.fontSize
    measure.style.fontFamily = cs.fontFamily
    measure.style.fontWeight = cs.fontWeight
    measure.style.letterSpacing = cs.letterSpacing
    measure.textContent = partial
    const w = measure.offsetWidth
    const padLeft = parseFloat(cs.paddingLeft) || 0
    const rect = editor.getBoundingClientRect()
    ghost.textContent = suffix
    ghost.style.display = 'block'
    ghost.style.left = `${rect.left + padLeft + w}px`
    ghost.style.top = `${rect.top}px`
    ghost.style.height = `${rect.height}px`
    ghost.style.lineHeight = `${rect.height}px`
    ghost.style.font = cs.font
    ghost.style.fontSize = cs.fontSize
    ghost.style.fontFamily = cs.fontFamily
    ghost.style.fontWeight = cs.fontWeight
    ghost.style.letterSpacing = cs.letterSpacing
  }
  // 选中某条候选：把完整文本填入当前单元格。
  // 关键：Univer 的单元格编辑器是 Docs 富文本结构，直接改 editor.textContent 不一定进其内部模型，
  // 因此单纯“改文本+派发合成 Enter”经常提交不进去（表现为“点了没填上”）。
  // 可靠做法：① 先尽量把编辑器文本改为所选值并触发 input（让编辑器内部模型同步）；
  // ② 执行 set-cell-edit-visible 关闭编辑器（编辑器会按当前内容提交一次）；
  // ③ 兜底：无论编辑器提交结果如何，延迟用 worksheet.setValue 强制把单元格写成所选值（权威写入）。
  // 这样即使富文本模型同步失败，单元格最终也一定是所选值，不会“点了没反应”。
  // item：候选（display=展示文本，raw=真实写入值；日期列 raw 为序列号，保证写回仍是真日期）。
  const acceptAc = (item: AcItem) => {
    const editor = acEditorRef.current
    const st = acStateRef.current
    const col = st.col
    const row = st.row
    const value = item?.display ?? ''
    const raw = item?.raw
    if (!value && raw == null) {
      hideAc()
      return
    }
    // 立即收起弹层与灰色续接，避免重复触发 / 编辑器关闭后再次弹出
    hideAc()
    const univerAPI = univerRef.current
    if (editor) {
      try {
        editor.focus()
        editor.textContent = value
        editor.dispatchEvent(new InputEvent('input', { bubbles: true }))
      } catch {
        /* noop */
      }
    }
    // 关闭编辑器（按当前内容提交一次）
    try {
      ;(univerAPI as unknown as { executeCommand: (id: string, p?: object) => unknown })
        .executeCommand('sheet.operation.set-cell-edit-visible', { visible: false })
    } catch {
      /* noop */
    }
    // 兜底强制写入：编辑器关闭后再写一次，确保单元格一定是所选值（日期列写序列号 raw）
    if (row > 0 && col >= 0) {
      window.setTimeout(() => {
        try {
          const ws = univerRef.current?.getActiveSheet()?.worksheet
          ws?.getRange(row, col).setValue(raw != null ? raw : value)
        } catch {
          /* noop */
        }
      }, 60)
    }
    // 锁定 300ms：接纳后编辑器尚在关闭过程中，轮询可能在下一帧读到刚写入的 value 而重新弹层；
    // 锁定期内轮询直接跳过，待编辑器被 Univer 移除（轮询检测到无编辑器）后自然收起。
    acAcceptingRef.current = true
    window.setTimeout(() => {
      acAcceptingRef.current = false
    }, 300)
    pushHistory(col, raw != null ? String(raw) : value)
  }

  // 关闭文件：若有未保存修改，先自动保存（不弹确认框），保存成功后再关闭；
  // 若保存失败则保持面板打开，避免静默丢数据。
  const handleClose = async () => {
    // 关闭前先取消待触发的自动保存定时器，避免卸载后再尝试写盘
    if (autoSaveTimerRef.current != null) {
      clearTimeout(autoSaveTimerRef.current)
      autoSaveTimerRef.current = null
    }
    if (dirtyRef.current) {
      setStatus('正在自动保存…')
      await handleSave()
      // 保存失败（dirty 仍为 true）时保持面板打开，让用户看到错误并重试
      if (dirtyRef.current) return
    }
    onClose()
  }

  const handleRestore = async () => {
    const list = await api.listBackups(fileRef.current.filePath)
    if (list.error || !list.backups || !list.backups.length) {
      setStatus('没有可用的备份')
      return
    }
    if (!window.confirm('将恢复到上一个保存前的版本，确定吗？')) return
    const r = await api.restoreBackup(fileRef.current.filePath)
    if (r.error) {
      setStatus('恢复失败：' + (r.message || ''))
      return
    }
    setStatus('已恢复上一版本')
    setReloadToken((t) => t + 1)
  }

  // 初始化 Univer（打开文件时灌数据；reloadToken 变化 = 恢复备份后重新加载）
  useEffect(() => {
    if (!containerRef.current) return
    let disposed = false
    const { univerAPI } = createUniver({
      locale: LocaleType.ZH_CN,
      locales: { [LocaleType.ZH_CN]: UniverPresetSheetsCoreZhCN },
      presets: [UniverSheetsCorePreset({ container: containerRef.current })],
    })
    univerRef.current = univerAPI

    // 结构性变更（删/插行、列、区域移动等）不一定触发 SheetValueChanged，
    // 但会改变数据，必须标记为“有修改”并触发自动保存。
    const dispCmd = univerAPI.addEvent(univerAPI.Event.CommandExecuted, (e: unknown) => {
      const ev = e as { id?: string }
      if (
        ev && typeof ev.id === 'string' &&
        SHEET_STRUCT_MUTATION_KEYWORDS.some((k) => ev.id!.includes(k))
      ) {
        markDirty()
        scheduleAutoSave()
      }
    })

    // 窗口尺寸变化（含最大化）时，列宽等比自适应铺满，不留右侧空白。
    const onResize = () => {
      if (resizeRafRef.current != null) cancelAnimationFrame(resizeRafRef.current)
      resizeRafRef.current = requestAnimationFrame(() => {
        resizeRafRef.current = null
        applyProportionalColumnWidths()
      })
    }
    window.addEventListener('resize', onResize)

    // 编辑即标记脏；数量/单价变化自动重算金额（自动金额行），
    // 金额列被编辑时按“是否==数量×单价”重分类该行（手工金额不被覆盖）。
    const disp = univerAPI.addEvent(univerAPI.Event.SheetValueChanged, (e: unknown) => {
      markDirty()
      scheduleAutoSave()
      const ev = e as { effectedRanges?: Array<{ getRange: () => { startRow: number; endRow: number; startColumn: number; endColumn: number } }> }
      const ranges = ev?.effectedRanges
      if (!Array.isArray(ranges) || !ranges.length) return
      const ws = univerAPI.getActiveSheet()?.worksheet
      if (!ws) return
      for (const fr of ranges) {
        const rg = fr.getRange()
        const startRow = Math.max(1, rg.startRow)
        // 数量(列4)/单价(列5)列变化 → 重算金额；金额(列6)列变化 → 重分类
        const touchesQP = rg.startColumn <= 5 && rg.endColumn >= 4
        const touchesAmt = rg.startColumn <= 6 && rg.endColumn >= 6
        for (let r = startRow; r <= rg.endRow; r++) {
          try {
            if (touchesQP) recomputeAmount(ws, r, autoAmountRows.current)
            // 金额列被改（含一次粘贴覆盖到金额）：按“当前值==数量×单价”重新归类
            if (touchesAmt) onAmountChanged(ws, r, autoAmountRows.current)
          } catch {
            /* 单行失败不影响其他 */
          }
        }
        // 记录输入历史：仅“单格”提交时记录（避免 OCR 批量填入 / 金额整列重算污染历史）
        if (rg.startRow === rg.endRow && rg.startColumn === rg.endColumn) {
          try {
            const v = String(ws.getRange(rg.startRow, rg.startColumn).getValue() ?? '')
            pushHistory(rg.startColumn, v)
          } catch {
            /* 单行失败不影响其他 */
          }
        }
      }
    })

    ;(async () => {
      try {
        const res = await api.loadSheet(fileRef.current.filePath)
        if (disposed) return
        if (res.error) {
          setStatus('打开失败：' + (res.message || ''))
          return
        }
        const data = rowsToWorkbookData(res.rows)
        // 按加载数据初始化“自动金额行”集合（金额==数量×单价或待填 → 自动；否则手工）
        autoAmountRows.current = buildAutoAmountRows(res.rows)
        univerAPI.createWorkbook(data)
        // 激活单元格落在数据末尾空行，便于继续录入 / 默认「填入」位置；
        // 并滚动到该行，同时多留末尾几行真实数据在视野内（END_VISIBLE_LEAD）
        const r = Math.max(1, Math.min(res.rows.length + 1, 100000))
        setActiveAndScroll(univerAPI, r, 0, { leadRows: END_VISIBLE_LEAD })
        // 打开即用当前窗口宽度做列宽等比铺满（render 就绪后再设列宽，避免时机过早）
        requestAnimationFrame(() => applyProportionalColumnWidths())
      } catch (e) {
        if (!disposed) setStatus('打开失败：' + (e instanceof Error ? e.message : '未知错误'))
      }
    })()

    return () => {
      disposed = true
      window.removeEventListener('resize', onResize)
      if (resizeRafRef.current != null) cancelAnimationFrame(resizeRafRef.current)
      // 组件卸载时取消待触发的自动保存定时器
      if (autoSaveTimerRef.current != null) {
        clearTimeout(autoSaveTimerRef.current)
        autoSaveTimerRef.current = null
      }
      try {
        disp.dispose()
      } catch {
        /* noop */
      }
      try {
        dispCmd.dispose()
      } catch {
        /* noop */
      }
      try {
        univerAPI.dispose()
      } catch {
        /* noop */
      }
      // 兜底清理：编辑单元格时 Univer 的 DOCS 单元格编辑器
      // （div#univer-doc-selection-container-* 内的 contentEditable[data-u-comp='editor']）
      // 会抢走键盘焦点，且可能被 reparent 到 document.body 上、dispose 未同步移除。
      // 若残留，这个隐藏输入框会持续“吞掉”键盘事件，导致关闭表格后
      // 左侧搜索框“卡住、无法输入”。这里强制移除残留容器并释放其焦点。
      try {
        const editors = document.querySelectorAll("div[id^='univer-doc-selection-container-']")
        editors.forEach((el) => {
          const ed = el.querySelector("[data-u-comp='editor']") as HTMLElement | null
          if (ed && document.activeElement === ed) ed.blur()
          el.remove()
        })
      } catch {
        /* noop */
      }
      univerRef.current = null
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [file.filePath, reloadToken])

  // 录制窗口「填入」：写入到当前激活单元格所在行（首列起），并推进激活位置
  useEffect(() => {
    const off = api.on('apply-recognized-rows', (rows) => {
      const univerAPI = univerRef.current
      if (!univerAPI) return
      const target = univerAPI.getActiveSheet()
      if (!target) return
      const ws = target.worksheet
      const list = (rows as AIRecognizedRow[]) || []
      if (!list.length) return
      const ar = ws.getActiveRange()
      const startRow = ar ? Math.max(1, ar.getRow()) : 1
      // 数量/单价/金额数值化（写数字而非文本），否则像 =E*F 的公式会因“文本*文本”得到 #VALUE!。
      // 金额：AI 已识别则保留其数值；否则留空，由下方 SheetValueChanged 监听自动写 =E*F 活公式。
      const matrix = list.map((r) => {
        const row = mapRecognizedToRow(r) as (string | number)[]
        // 数量/单价数值化（写数字而非文本），否则像 =E*F 的公式会因"文本*文本"得到 #VALUE!。
        // 金额列已被 mapRecognizedToRow 置空，交给 SheetValueChanged 监听自动写 =E*F 活公式。
        const q = Number(row[4] || 0)
        const p = Number(row[5] || 0)
        row[4] = q > 0 ? q : ''
        row[5] = p > 0 ? p : ''
        return row
      })
      try {
        ws.getRange(startRow, 0, matrix.length, COL_COUNT).setValues(matrix)
        // 填入行金额留空（pending），若数量/单价为正，纳入“自动金额行”集合，
        // 方便后续改数量/单价时金额自动同步（与加载初始化口径一致）。
        matrix.forEach((row, i) => {
          const q = Number(row[4] || 0)
          const p = Number(row[5] || 0)
          if (q > 0 && p > 0) autoAmountRows.current.add(startRow + i)
        })
        // 推进激活格到填入内容之后的空行，但保持当前视窗不动（数据落在光标处，不跳走）
        setActiveAndScroll(univerAPI, startRow + matrix.length, 0, { scroll: false })
      } catch (e) {
        setStatus('填入失败：' + (e instanceof Error ? e.message : '未知错误'))
        return
      }
      markDirty()
      scheduleAutoSave()
      setStatus(`已填入 ${matrix.length} 行识别结果`)
    })
    return () => off()
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [])

  // Ctrl+S 保存
  useEffect(() => {
    const onKey = (e: KeyboardEvent) => {
      if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 's') {
        e.preventDefault()
        void saveRef.current()
      }
    }
    window.addEventListener('keydown', onKey)
    return () => window.removeEventListener('keydown', onKey)
  }, [])

  // 单元格输入记忆（自动补全）：轮询式触发。
  // 为什么用轮询而非“监听原生 input 事件”：
  // ① Univer 的 Docs 单元格编辑器是 contentEditable，敲首字符时 input 事件虽触发，但此刻字符
  //    尚未提交进 DOM（textContent 读到空），导致“输 9 一个提示都没有”；轮询每帧重读 textContent，
  //    首帧可能为空、下一帧文本已落 DOM，自然补齐首字符。
  // ② 直接 document.querySelector('[data-u-comp="editor"]') 定位编辑器，不依赖 document.activeElement
  //    恰好是编辑器、也不依赖 input 事件能否冒泡/被捕获，规避此前多次“完全不提示”的脆弱点。
  // ③ 仅在“文本或所在单元格变化”时才重算候选（acLastKeyRef 去重），空闲时每帧仅一次轻量 querySelector。
  // 弹层出现后，上下键/回车/Tab/Esc 仍用捕获阶段 keydown 拦截，做到“像 Excel 一样边打边提示”。
  useEffect(() => {
    console.log('[AC-DEBUG] v5 mounted; AC_DEBUG=' + AC_DEBUG)
    // 编辑器查找：依次尝试一组选择器；只要其中任何一个存在且包含/自身是 INPUT/TEXTAREA/contenteditable
    // 都算命中。Univer 版本之间包装层有时是 [data-u-comp='editor']、有时是 slate 编辑区
    // （带 [data-slate-editor]），多选择器兜底更稳。
    const EDITOR_SELECTORS = [
      "[data-u-comp='editor']",
      "[data-u-comp='cell-editor']",
      '.univer-cell-editor',
      '[data-slate-editor]',
      '[contenteditable]:not([data-u-comp])',
    ]
    // 从匹配到的根向下找出真正承担输入的"叶子"：INPUT/TEXTAREA/contenteditable
    const pickInnerInput = (el: HTMLElement): { node: HTMLElement; via: string } | null => {
      const ownTag = (el.tagName || '').toUpperCase()
      if (ownTag === 'INPUT' || ownTag === 'TEXTAREA') return { node: el, via: 'self:' + ownTag }
      const ceAttr = el.getAttribute('contenteditable')
      if (ceAttr === 'true' || ceAttr === '') return { node: el, via: 'self:CE' }
      const innerInp = el.querySelector('input, textarea') as HTMLElement | null
      if (innerInp) return { node: innerInp, via: 'inner:' + innerInp.tagName }
      const innerCE = el.querySelector('[contenteditable="true"], [contenteditable=""]') as HTMLElement | null
      if (innerCE) return { node: innerCE, via: 'inner:CE' }
      return null
    }
    const getCellEditor = (): { el: HTMLElement | null; sel: string; inner: { node: HTMLElement; via: string } | null } => {
      for (const sel of EDITOR_SELECTORS) {
        try {
          const el = document.querySelector(sel) as HTMLElement | null
          if (!el) continue
          // 这个候选是不是真"在编辑"：自己或子孙能找到 INPUT/TEXTAREA/contenteditable 才算
          const inner = pickInnerInput(el)
          if (inner) return { el, sel, inner }
          // 自己不是也没子节点：记录但暂时不用，继续找下一个选择器
        } catch {
          /* noop */
        }
      }
      // 兜底：如果所有带优先选择器的都没有 inner，但首选择器命中了（说明它只是个永远存在的占位）
      // 还是返回第一个命中节点 + sel，但 inner=null，让 HUD 把这种 case 直接暴露出来
      for (const sel of EDITOR_SELECTORS) {
        try {
          const el = document.querySelector(sel) as HTMLElement | null
          if (el) return { el, sel, inner: null }
        } catch {
          /* noop */
        }
      }
      return { el: null, sel: '', inner: null }
    }
    // 读"输入区"真实文本：INPUT/TEXTAREA 读 .value；contenteditable 读 textContent
    const readInputText = (inner: { node: HTMLElement; via: string } | null, fallback: HTMLElement | null): string => {
      const node = inner?.node || fallback
      if (!node) return ''
      const tag = (node.tagName || '').toUpperCase()
      if (tag === 'INPUT' || tag === 'TEXTAREA') {
        return ((node as HTMLInputElement | HTMLTextAreaElement).value || '').replace(/\s+/g, ' ').trim()
      }
      // contenteditable：textContent
      return ((node.textContent || '').replace(/\s+/g, ' ')).trim()
    }

    const dbg = (s: string) => {
      if (!AC_DEBUG) return
      const d = acDebugRef.current
      if (d) d.textContent = s
    }
    const tick = () => {
      // 自调度，保证持续轮询
      acRafRef.current = requestAnimationFrame(tick)
      // 接纳候选过程中（编辑器正关闭）跳过，避免读到刚写入的 value 而重新弹层
      if (acAcceptingRef.current) {
        dbg('[accepting 锁定中]')
        return
      }
      try {
        const univerAPI = univerRef.current
        const { el: editor, sel, inner } = getCellEditor()
        // 把编辑器根和 inner 的标签摘要拼出来，无论命中与否都暴露 DOM 真相
        const tagSummary = editor
          ? '<' + (editor.tagName || '').toLowerCase() + '>' +
            (inner ? '/' + inner.via + ':<' + (inner.node.tagName || '').toLowerCase() + '>' :
              (() => {
                // 命中但没 inner：暴露外层首 160 字 innerHTML，下一轮就能直接定真选择器
                const snippet = (editor.innerHTML || '').slice(0, 160).replace(/\s+/g, ' ').trim()
                return '/no-inner html=' + snippet
              })())
          : ''
        if (!editor) {
          // 编辑器一个都没选中：把整页“像编辑器”的 DOM 全部探测出来，直接告诉我们真选择器。
          acLastKeyRef.current = ''
          if (acStateRef.current.open) hideAc()
          const probe: string[] = []
          const ucomps = new Set<string>()
          document.querySelectorAll('[data-u-comp]').forEach((e) => {
            const v = e.getAttribute('data-u-comp')
            if (v) ucomps.add(v)
          })
          probe.push('ucomps=' + [...ucomps].join('|'))
          const editableEls = Array.from(document.querySelectorAll('[contenteditable]')) as HTMLElement[]
          probe.push('editable=' + editableEls.length)
          editableEls.slice(0, 3).forEach((e, i) => {
            probe.push('ce' + i + '=<' + e.tagName.toLowerCase() + (e.className ? '.' + String(e.className).split(/\s+/).slice(0, 2).join('.') : '') + '>')
          })
          const inputs = Array.from(document.querySelectorAll('input,textarea')) as HTMLElement[]
          probe.push('inputs=' + inputs.length)
          inputs.slice(0, 3).forEach((e, i) => {
            probe.push('inp' + i + '=<' + e.tagName.toLowerCase() + (e.id ? '#' + e.id : '') + (e.className ? '.' + String(e.className).split(/\s+/).slice(0, 2).join('.') : '') + '>')
          })
          const containers = Array.from(document.querySelectorAll("div[id^='univer-doc-selection-container-']")) as HTMLElement[]
          probe.push('containers=' + containers.length)
          if (containers[0]) {
            const c = containers[0]
            probe.push('cont=<' + c.tagName.toLowerCase() + (c.className ? '.' + String(c.className).split(/\s+/).slice(0, 2).join('.') : '') + '> kids=' + c.children.length)
            const k = c.children[0] as HTMLElement | undefined
            if (k) probe.push('kid0=<' + k.tagName.toLowerCase() + (k.className ? '.' + String(k.className).split(/\s+/).slice(0, 2).join('.') : '') + '> html=' + (k.innerHTML || '').slice(0, 160))
          }
          dbg('ed=none | ' + probe.join(' | '))
          return
        }
        const ar = univerAPI?.getActiveSheet()?.worksheet?.getActiveRange?.()
        // 编辑时 getActiveRange() 可能返回 null（编辑器接管选区）→ 回退到最近一次激活单元格
        let col = -1
        let row = -1
        let usingFallback = false
        if (ar) {
          col = ar.getColumn()
          row = ar.getRow()
          acLastCellRef.current = { col, row }
        } else if (acLastCellRef.current) {
          col = acLastCellRef.current.col
          row = acLastCellRef.current.row
          usingFallback = true
        } else {
          acLastKeyRef.current = ''
          if (acStateRef.current.open) hideAc()
          dbg('ed=' + sel + ' ' + tagSummary + ' txt=' + JSON.stringify(readInputText(inner, editor).slice(0, 24)) + ' ar=null(no-fallback)')
          return
        }
        const text = readInputText(inner, editor)
        const key = col + ':' + row + ':' + text
        if (key === acLastKeyRef.current) return // 未变化：保持当前弹层，不重复计算
        acLastKeyRef.current = key
        acEditorRef.current = inner?.node || editor
        // 列范围硬约束：只在账单表 9 列内才弹候选（先判断，避免在账外点时永远停在 txt=空 而看不到真正原因）
        const colOut = col < 0 || col >= COL_COUNT
        if (colOut) {
          hideAc()
          dbg('ed=' + sel + ' ' + tagSummary + ' ar=' + (usingFallback ? 'fb:' : '') + col + ',' + row + ' txt=' + JSON.stringify(text) + ' SKIP:col-out[0..' + (COL_COUNT - 1) + ']')
          return
        }
        if (!text) {
          hideAc()
          dbg('ed=' + sel + ' ' + tagSummary + ' ar=' + (usingFallback ? 'fb:' : '') + col + ',' + row + ' txt=空 (col ok)')
          return
        }
        const colVals = getColumnValuesWithRows(col)
        const items = computeSuggestions(text, col, row)
        if (!items.length) {
          hideAc()
          dbg('ed=' + sel + ' ' + tagSummary + ' ar=' + (usingFallback ? 'fb:' : '') + col + ',' + row + ' txt=' + JSON.stringify(text) + ' colVals=' + colVals.length + ' sample=' + JSON.stringify(colVals.slice(0, 3).map((v) => v.value)) + ' sug=0')
          return
        }
        acStateRef.current = { ...acStateRef.current, open: true, items, index: 0, col, row, partial: text }
        showAc(items, editor.getBoundingClientRect())
        // Excel 式“框内灰色续接”提示
        updateGhost(text, items, acEditorRef.current)
        dbg('ed=' + sel + ' ' + tagSummary + ' ar=' + (usingFallback ? 'fb:' : '') + col + ',' + row + ' txt=' + JSON.stringify(text) + ' colVals=' + colVals.length + ' sample=' + JSON.stringify(colVals.slice(0, 3).map((v) => v.value)) + ' sug=' + items.length + ' OPEN')
      } catch (err) {
        dbg('ERR ' + (err instanceof Error ? err.message : String(err)))
      }
    }

    const onEditorKeydown = (e: KeyboardEvent) => {
      if (!acStateRef.current.open) return
      const st = acStateRef.current
      const items = st.items
      if (e.key === 'ArrowDown') {
        e.preventDefault()
        e.stopPropagation()
        const idx = (st.index + 1) % items.length
        acStateRef.current = { ...st, index: idx }
        setAcUi((u) => ({ ...u, index: idx }))
      } else if (e.key === 'ArrowUp') {
        e.preventDefault()
        e.stopPropagation()
        const idx = (st.index - 1 + items.length) % items.length
        acStateRef.current = { ...st, index: idx }
        setAcUi((u) => ({ ...u, index: idx }))
      } else if (e.key === 'Enter' || e.key === 'Tab') {
        if (items.length) {
          e.preventDefault()
          e.stopPropagation()
          acceptAc(items[st.index])
        }
      } else if (e.key === 'Escape') {
        e.preventDefault()
        e.stopPropagation()
        hideAc()
      }
    }

    acRafRef.current = requestAnimationFrame(tick)
    document.addEventListener('keydown', onEditorKeydown, true)
    return () => {
      if (acRafRef.current != null) cancelAnimationFrame(acRafRef.current)
      document.removeEventListener('keydown', onEditorKeydown, true)
      if (acMeasureRef.current && acMeasureRef.current.parentElement) {
        acMeasureRef.current.parentElement.removeChild(acMeasureRef.current)
        acMeasureRef.current = null
      }
    }
    // 这些回调只依赖 refs / 稳定 setter，故仅挂载一次
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [])

  return (
    <div className="univer-sheet">
      <div className="memo-header">
        <div className="memo-file">
          <span className="memo-file-name">{file.fileName}</span>
          {dirty && <span className="dirty-dot" title="有未保存修改">●</span>}
        </div>
        <div className="memo-header-actions">
          <button className="btn btn-primary" disabled={saving} onClick={() => void handleSave()}>
            {saving ? '保存中…' : '保存 Ctrl+S'}
          </button>
          <button className="btn btn-outline" onClick={() => api.openFile(file.filePath)} title="用 Excel 程序打开">
            Excel 打开
          </button>
          <button className="btn btn-outline" disabled={saving} onClick={() => void handleRestore()} title="恢复保存前的上一个版本">
            恢复上一版本
          </button>
          <button className="btn btn-outline" onClick={handleClose} title="关闭">
            关闭
          </button>
        </div>
      </div>
      <div ref={containerRef} className="univer-container" />
      {/* Excel 式“框内灰色续接”：绝对定位叠加在编辑器已输入文本之后，仅展示，不拦截事件 */}
      <div className="ac-ghost" ref={acGhostRef} />
      {AC_DEBUG && <div className="ac-debug" ref={acDebugRef} />}
      {acUi.open && acUi.pos && (
        <div className="ac-popup" style={{ left: acUi.pos.left, top: acUi.pos.top }}>
          {acUi.items.map((it, i) => (
            <div
              key={i}
              className={'ac-item' + (i === acUi.index ? ' ac-active' : '')}
              onMouseEnter={() => {
                acStateRef.current = { ...acStateRef.current, index: i }
                setAcUi((u) => ({ ...u, index: i }))
              }}
              onMouseDown={(e) => {
                e.preventDefault()
                acceptAc(it)
              }}
            >
              {it.display}
            </div>
          ))}
        </div>
      )}
      {status && <div className="memo-status">{status}</div>}
    </div>
  )
}
