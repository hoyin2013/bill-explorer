import { useEffect, useRef, useState } from 'react'
import { createUniver, LocaleType } from '@univerjs/presets'
import { UniverSheetsCorePreset } from '@univerjs/preset-sheets-core'
import UniverPresetSheetsCoreZhCN from '@univerjs/preset-sheets-core/locales/zh-CN'
import { IEditorBridgeService } from '@univerjs/sheets-ui'
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
import { extractColumnValues, type AcItem } from '../lib/autocomplete'
// 单元格输入记忆（自动补全）控制器：基于 Univer 官方「编辑桥接服务」事件驱动，
// 彻底替代此前每帧 querySelector 轮询 DOM 编辑器（被公式栏占位壳反复坑的脆弱方案）。
// 控制器逻辑见 src/renderer/lib/acController.ts，本组件只提供「候选算法 / 写入 / 渲染」三件套。
import { attachAutocomplete, type AcController } from '../lib/acController'

type UniverAPI = ReturnType<typeof createUniver>['univerAPI']

interface Props {
  file: FileEntry
  api: ElectronAPI
  onClose: () => void
  onSaved: () => void
  /** 组件挂载时回调父组件，传入 switchFile 供切换前 await */
  onRegisterSwitchFile?: (fn: () => Promise<boolean>) => void
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

// 多选区域按退格/Delete 应清空整个选区（Excel/WPS 语义）。
// 根因：Univer 在"选择"模式下对 Backspace 有两条快捷键——ClearSelectionContentCommand（清选区）
// 与 EditorDeleteLeftShortcutInActive（binding: BACKSPACE，进入左上角一格编辑）——后者在多选时优先匹配，
// 导致用户按 Backspace 时只进入左上角一格编辑、而非清空整个选区。
// 修复：在 window 捕获阶段拦截，当"未进入编辑"且"选区为多格"时，主动执行 sheet.command.clear-selection-content
// 清空整个选区；单格选区（此时 Backspace 应进入编辑）与已处于编辑状态时，一律放行由 Univer 原生处理。
const CLEAR_SELECTION_CMD_ID = 'sheet.command.clear-selection-content'

// WPS 风格自动保存：停顿 AUTOSAVE_DELAY 毫秒后保存（比原 1s 更长，减少大文件存盘频率，
// 避免每次按键停顿都触发一次整表存盘导致的卡顿）。
const AUTOSAVE_DELAY = 5000
// 周期性保存：即使持续编辑不暂停，也每隔 IDLE_SAVE_INTERVAL 检测并保存一次，
// 作为"意外丢失"的安全网（WPS 默认约 10 分钟，此处取 2 分钟以兼顾安全与性能）。
const IDLE_SAVE_INTERVAL = 600_000

// 单元格输入记忆（自动补全）：编辑单元格时，根据「本列已有值 + 历史记录」给出提示，辅助快速输入。
// 行为类似 Excel 的按列自动完成：你打字时，下方浮现本列曾出现过的、以及你以前输入过的相近内容。
// - AC_MAX_SUGGESTIONS：弹出候选最多条数
// - AC_HISTORY_CAP：每个「文件+列」持久化保留的历史条数上限（localStorage，跨会话记忆）
const AC_MAX_SUGGESTIONS = 12
const AC_HISTORY_CAP = 200
const AC_HISTORY_PREFIX = 'bill-ac-history::'
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
// （日期列=Excel 序列号，保证仍为真日期；其余列=原文本）。类型统一复用 lib/autocomplete.ts 的 AcItem。

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

export function UniverSheet({ file, api, onClose, onSaved, onRegisterSwitchFile }: Props) {
  const containerRef = useRef<HTMLDivElement>(null)
  const univerRef = useRef<UniverAPI | null>(null)
  // createUniver 同时返回 Univer 实例；自动补全控制器经其注入器取出官方编辑服务（见 acController.ts）。
  const univerInstRef = useRef<any>(null)
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
  // - autoSaveIntervalRef：周期性保存定时器（IDLE_SAVE_INTERVAL 间隔检测脏状态，触发一次防抖保存）
  const autoSaveIntervalRef = useRef<ReturnType<typeof setInterval> | null>(null)
  const clearKeyListenerRef = useRef<() => void>(() => {})
  // 单元格输入记忆（自动补全）相关状态：
  // - acRef：attachAutocomplete 返回的控制器实例（暴露 setIndex / accept 给鼠标交互，dispose 在卸载时调用）
  // - acUi：仅供渲染弹层的快照（open/items/index/定位 rect）；状态机在控制器内维护，避免打字时频繁重渲染
  const acRef = useRef<AcController | null>(null)
  const [acUi, setAcUi] = useState<{
    open: boolean
    items: AcItem[]
    index: number
    pos: { left: number; top: number; width: number } | null
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
        // 保存失败：保持 dirty 状态，错误持久显示在状态栏（不会自动消失），
        // 提示用户文件可能被 Excel 打开锁定（修复 U3）
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
  // 读取当前工作表某一列里出现过的所有值。同名值在多处出现时每个 row 都贡献一条候选，
  // 避免聚合时 max(row) 把同名值的 row 推高、然后被 row<currentRow filter 误过滤掉
  // （这正是历史上 "敲字但不弹" 的真实根因之一 —— 算法 bug，e2e 测试已复现并修复）。
  // 见 src/renderer/lib/autocomplete.ts 的 extractColumnValues 单元测试。
  const getColumnValuesWithRows = (col: number): Array<{ value: string | number; row: number; freq: number }> => {
    const univerAPI = univerRef.current
    const wb = univerAPI?.getActiveWorkbook()?.getSnapshot()
    if (!wb) return []
    // 直接复用 lib 模块的纯函数（已通过 10 个单元测试 + e2e 验证）
    return extractColumnValues(wb, col, 'bill')
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
  // 选中候选后的“写入”：直接把所选值写回单元格（日期列写序列号 raw，保证仍是真日期），
  // 并记录到该列历史。编辑器关闭由控制器统一处理，这里只负责最终落盘。
  // 写入选中候选：实际落盘由 Univer 原生提交完成（acController 在 SetRangeValuesCommand 执行前
  // 把提交值替换为候选，从而绕过 currentEditCell$.row 的 +1 偏移）。此处仅负责把该值记入历史，
  // 供后续打字联想。故不再直接 setValue（直接写会用偏移后的 row，反而写错行）。
  const commitAc = (item: AcItem, col: number) => {
    const raw = item.raw
    const display = item.display
    pushHistory(col, raw != null ? String(raw) : display)
  }

  // 切换文件前：若有未保存修改，先等保存完成再放行（避免静默丢数据）
  // App.tsx 通过 sheetSaveBeforeSwitchRef 持有此函数，在 onOpen 里 await
  const switchFile = async (): Promise<boolean> => {
    if (autoSaveTimerRef.current != null) {
      clearTimeout(autoSaveTimerRef.current)
      autoSaveTimerRef.current = null
    }
    if (!dirtyRef.current) return true
    await handleSave()
    return !dirtyRef.current
  }

  // 组件挂载后把 switchFile 暴露给父组件（供切换文件前 await）
  useEffect(() => {
    onRegisterSwitchFile?.(switchFile)
  }, [onRegisterSwitchFile])

  // 关闭文件：若有未保存修改，先自动保存（不弹确认框），保存成功后再关闭；
  // 若保存失败则保持面板打开，避免静默丢数据。
  const handleClose = async () => {
    // 关闭前先取消待触发的自动保存定时器与周期性保存定时器，避免卸载后再尝试写盘
    if (autoSaveTimerRef.current != null) {
      clearTimeout(autoSaveTimerRef.current)
      autoSaveTimerRef.current = null
    }
    if (autoSaveIntervalRef.current != null) {
      clearInterval(autoSaveIntervalRef.current)
      autoSaveIntervalRef.current = null
    }
    clearKeyListenerRef.current?.()
    // 复用 switchFile 逻辑：有改动就保存，保存失败则保持面板打开
    const ok = await switchFile()
    if (!ok) {
      setStatus('关闭前保存失败，请先解决保存问题后再关闭（文件可能被 Excel 打开锁定）。如有需要可点「关闭」前手动另存。')
      return
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
    const { univer, univerAPI } = createUniver({
      locale: LocaleType.ZH_CN,
      locales: { [LocaleType.ZH_CN]: UniverPresetSheetsCoreZhCN },
      // 关键：禁用 Univer 的两个"文本格式数字"告警/标记——
      //   1) disableForceStringAlert/Mark（sheets-ui）：当单元格为文本格式但值为数字时，
      //      会弹出一个"以文本形式存储的数字"错误气泡（forceStringInfo），该气泡的 overlay 会阻断
      //      交互约 1 分钟后自动消失，表现为"输入时报错然后卡死 1 分钟自行恢复"。
      //   2) disableTextFormatAlert/Mark（sheets-numfmt）：同一机制在 numfmt 插件中也会触发，
      //      并且其拦截器会强制把数字值改为 STRING 类型，导致"7.8 变成 80"之类的显示错乱。
      // 本应用的数据格式由我们自己控制（数量/单价/金额存为数字、日期存为序列号），不存在"文本格式数字"
      // 的语义需求，故禁用这两个告警既消除报错/卡死，也彻底杜绝数值被误改为字符串。
      presets: [UniverSheetsCorePreset({
        container: containerRef.current,
        disableTextFormatAlert: true,
        disableTextFormatMark: true,
        sheets: {
          disableForceStringAlert: true,
          disableForceStringMark: true,
        },
      })],
    })
    univerRef.current = univerAPI
    univerInstRef.current = univer

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
    // RAF 节流 + 150ms 防抖：拖拽 resize 时 16ms/帧触发 RAF，但列宽只在 150ms 停顿后应用，
    // 避免拖拽过程中每帧都触发 Univer 9 次 setColumnWidth + layout 重算（修复 P3）。
    let resizeTimer: ReturnType<typeof setTimeout> | null = null
    const onResize = () => {
      if (resizeRafRef.current != null) cancelAnimationFrame(resizeRafRef.current)
      resizeRafRef.current = requestAnimationFrame(() => {
        resizeRafRef.current = null
        if (resizeTimer != null) clearTimeout(resizeTimer)
        resizeTimer = setTimeout(() => {
          resizeTimer = null
          applyProportionalColumnWidths()
        }, 150)
      })
    }
    window.addEventListener('resize', onResize)

    // 周期性自动保存（WPS 风格）：每隔 IDLE_SAVE_INTERVAL 检测是否有未保存修改，有则触发一次防抖保存。
    // 与"停顿后保存"互补：前者覆盖"持续编辑不暂停"的场景（作为意外丢失安全网），
    // 后者覆盖"完成一行后短暂停顿"的场景。二者共用同一套防抖/并发保护，不会叠加写盘。
    autoSaveIntervalRef.current = setInterval(() => {
      if (dirtyRef.current) scheduleAutoSave()
    }, IDLE_SAVE_INTERVAL)

    // 多选退格/Delete 修复：获取 editorBridge 以便拦截时判断"是否正在编辑"。
    // （编辑器未激活 = 处于选择模式，此时退格应清空整个选区；编辑器已激活 = 退格应删除字符，放行。）
    const editorBridge = (() => {
      try {
        return (univerInstRef.current as any).__getInjector?.()?.get?.(IEditorBridgeService)
      } catch {
        return null
      }
    })()

    // window 捕获阶段拦截退格/Delete：多选 + 未编辑 → 清空整个选区；其余放行 Univer 原生处理。
    const onClearKeyCapture = (e: KeyboardEvent) => {
      const key = e.key
      if (key !== 'Backspace' && key !== 'Delete') return
      // 正在编辑单元格（编辑器可见）：放行，由 Univer 处理"退格删字符"
      const vb = editorBridge?.visible$?.getValue?.()
      if (vb?.visible) return
      const ws = univerRef.current?.getActiveSheet()?.worksheet
      if (!ws) return
      const ar = ws.getActiveRange()
      if (!ar) return
      // FRange 的 getRange() 返回内部 Range（带 startRow/startColumn/endRow/endColumn）
      const range = (ar as { getRange?: () => { startRow?: number; endRow?: number; startColumn?: number; endColumn?: number } }).getRange?.()
      if (!range || (typeof range.startRow !== 'number' && typeof (range as any).startRow !== 'number')) return
      const sr = (range as { startRow: number; startColumn: number; endRow: number; endColumn: number }).startRow
      const er = (range as { startRow: number; startColumn: number; endRow: number; endColumn: number }).endRow
      const sc = (range as { startRow: number; startColumn: number; endRow: number; endColumn: number }).startColumn
      const ec = (range as { startRow: number; startColumn: number; endRow: number; endColumn: number }).endColumn
      const isMulti = (er - sr) > 0 || (ec - sc) > 0
      if (!isMulti) return
      // 多选且未编辑：清空整个选区（Excel/WPS 语义）
      e.preventDefault()
      e.stopPropagation()
      try {
        const api = univerRef.current as unknown as { executeCommand: (id: string) => unknown }
        void api.executeCommand(CLEAR_SELECTION_CMD_ID)
      } catch {
        /* noop */
      }
    }
    window.addEventListener('keydown', onClearKeyCapture, true)
    const onRemoveClearKeyCapture = () => window.removeEventListener('keydown', onClearKeyCapture, true)
    clearKeyListenerRef.current = onRemoveClearKeyCapture

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
        // 挂载单元格输入记忆（自动补全）控制器：用 Univer 官方编辑桥接服务驱动。
        // 关键：必须在 createWorkbook 之后挂载 —— Univer 的 IEditorBridgeService（编辑器桥接服务）
        // 要等第一个工作表真正实例化后才完成注册；在此之前 injector.get(IEditorBridgeService)
        // 会抛「Expect 1 dependency item(s) ... but get 0」，attachAutocomplete 只能静默降级为 noop，
        // 表现为「输入过程完全不出现提示」。故把挂载点从 effect 顶层挪到 createWorkbook 之后。
        acRef.current = attachAutocomplete(univer, {
          computeItems: (t, col, row) => computeSuggestions(t, col, row),
          commit: (item, col) => commitAc(item, col),
          render: (s) => setAcUi(s),
          container: containerRef.current,
          univerAPI: univerAPI as unknown as { executeCommand: (id: string, params?: object) => unknown },
        })
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
      // 组件卸载时取消待触发的自动保存定时器与周期性保存定时器
      if (autoSaveTimerRef.current != null) {
        clearTimeout(autoSaveTimerRef.current)
        autoSaveTimerRef.current = null
      }
      if (autoSaveIntervalRef.current != null) {
        clearInterval(autoSaveIntervalRef.current)
        autoSaveIntervalRef.current = null
      }
      clearKeyListenerRef.current?.()
      // 卸载时若还有未保存修改，尝试保存（安全网；不阻塞卸载）
      if (dirtyRef.current) void handleSave()
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
        acRef.current?.dispose()
      } catch {
        /* noop */
      }
      acRef.current = null
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
      univerInstRef.current = null
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

  return (
    <div className="univer-sheet">
      <div className="memo-header">
        <div className="memo-file">
          <span className="memo-file-name">{file.fileName}</span>
          {dirty && <span className="dirty-dot" title="有未保存修改">●</span>}
        </div>
        <div className="memo-header-actions">
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
      {acUi.open && acUi.pos && (
        <div
          className="ac-popup"
          style={{ left: acUi.pos.left, top: acUi.pos.top, minWidth: acUi.pos.width }}
          // 关键：mousedown/click 都拦截，绝不漏到 Univer 画布。
          // 候选项的 accept 放在 onClick 触发（而非 mousedown）：
          // 若 mousedown 上就 accept+关弹层，弹层会在「mousedown→mouseup」整段手势中途被移除，
          // 随后的 mouseup/click 落到下方画布，被 Univer 解读为“点到了某格 → 重新进入该格编辑”，
          // 于是 currentEditCell$ 再次触发、弹层重新弹出（表现为“点第一次能带入但弹层不消失，再点一次才消失”）。
          // 改为 onClick 触发后，弹层在整个手势期间一直存在，click 始终落在弹层内部、被 stopPropagation 吞掉，
          // 画布收不到任何事件，从根本上杜绝重开。
          onMouseDown={(e) => e.stopPropagation()}
          onClick={(e) => e.stopPropagation()}
        >
          {acUi.items.map((it, i) => (
            <div
              key={i}
              className={'ac-item' + (i === acUi.index ? ' ac-active' : '')}
              onMouseEnter={() => acRef.current?.setIndex(i)}
              // mousedown 只 preventDefault：保住编辑器焦点、不丢已输文本，但【先不】accept/关弹层；
              // 真正的 accept 放在 onClick（见下），让弹层陪完整个手势，避免事件漏到画布重开。
              onMouseDown={(e) => {
                e.preventDefault()
              }}
              onClick={(e) => {
                e.preventDefault()
                e.stopPropagation()
                acRef.current?.accept(it)
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
