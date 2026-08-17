import { useEffect, useRef, useState } from 'react'
import { createUniver, LocaleType } from '@univerjs/presets'
import { UniverSheetsCorePreset } from '@univerjs/preset-sheets-core'
import UniverPresetSheetsCoreZhCN from '@univerjs/preset-sheets-core/locales/zh-CN'
import '@univerjs/preset-sheets-core/lib/index.css'
import { ICommandService, CommandType, IConfigService } from '@univerjs/core'
import { KeyCode, IShortcutService } from '@univerjs/ui'
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

export function UniverSheet({ file, api, onClose, onSaved }: Props) {
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
    // 本列“上方”已有值（row < 当前行，与 Excel 自动完成范围一致）。
    // 排除表头行（row 0）：表头是列名（如「货品名称」「日期」），不是数据值，不应作为候选。
    const colVals = getColumnValuesWithRows(col).filter((x) => x.row >= 1 && x.row < currentRow)
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
    // partial 非空时优先「以已输入内容开头」匹配，无开头匹配时退化为「包含匹配」。
    // （注：当前仅在用户真正键入（hasInput）时才调用，partial 必非空；空分支为防御性保留。）
    let matched = p ? all.filter((x) => x.display.toLowerCase().startsWith(p)) : all
    if (p && !matched.length) matched = all.filter((x) => x.display.toLowerCase().includes(p))
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
    // 整行删除：自定义命令 + 高优先级快捷键的注册句柄（createWorkbook 之后注册，卸载时释放）
    const rowDeleteDisposables: { dispose(): void }[] = []
    const { univer, univerAPI } = createUniver({
      locale: LocaleType.ZH_CN,
      locales: { [LocaleType.ZH_CN]: UniverPresetSheetsCoreZhCN },
      presets: [UniverSheetsCorePreset({ container: containerRef.current })],
    })
    // 关闭 Univer「数字以文本形式存储」单元格悬停告警：该告警的 i18n 键
    // (sheets-ui.info.error / forceStringInfo) 在本版语言包中未定义，会渲染成裸键名；
    // 账单录入器无此提示需求，且货品名/备注常含数字文本，留着反而误报。
    univer.__getInjector().get(IConfigService)
         .setConfig('sheets-ui.config', { disableForceStringAlert: true })
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
        // 选中整行后按 Delete/Backspace：删除整行（移除该行、下方整体上移），而非仅清空内容。
        // 做法：注册一个自定义命令 + 高优先级快捷键（DELETE / BACKSPACE 双绑），仅在「整行选中」时
        // 由本快捷键接管；Univer 原生的「清除选区内容」快捷键（优先级更低）对同一按键不再触发，
        // 故干净地只删行、不残留清空。这避开了此前 window keydown 监听抢不过 Univer 原生处理
        // （其快捷键挂在 window 捕获阶段且注册更早）的问题。
        // 判定：非编辑态 + 起始列 0 且覆盖到工作表最右列（点行号选整行即此形态，也兼容拖选 A~I 整行）
        // + 非表头行（row 0）。单格/区域删除内容、编辑态删字符均保持原生行为。
        const injector = (univer as unknown as { __getInjector?: () => { get: (t: unknown) => unknown } }).__getInjector?.()
        const deleteCmdSvc = injector?.get(ICommandService) as
          | { registerCommand: (c: unknown) => { dispose(): void } }
          | undefined
        const shortcutSvc = injector?.get(IShortcutService) as
          | { registerShortcut: (s: unknown) => { dispose(): void } }
          | undefined
        const DeleteSelectedRowsCommand = {
          id: 'bill-explorer.command.delete-selected-rows',
          type: CommandType.COMMAND,
          handler: () => {
            const ws = univerRef.current?.getActiveWorkbook()?.getActiveSheet()
            if (!ws) return false
            const ar = ws.getActiveRange()
            if (!ar) return false
            const startRow = ar.getRow()
            const endRow = ar.getLastRow()
            const startCol = ar.getColumn()
            const endCol = ar.getLastColumn()
            if (startRow < 1) return false // 不删表头行
            if (startCol !== 0 || endCol < COL_COUNT - 1) return false // 仅整行选中才删整行
            const delCount = endRow - startRow + 1
            // 按删除范围调整自动金额行集合（删除区间丢弃、下方整体上移），避免行号错位算错金额
            const newSet = new Set<number>()
            for (const r of autoAmountRows.current) {
              if (r < startRow) newSet.add(r)
              else if (r > endRow) newSet.add(r - delCount)
              // r 落在删除区间 → 丢弃
            }
            autoAmountRows.current = newSet
            // 直接转发到 Univer 内置「删除选中行」命令（与右键菜单完全相同的命令路径），
            // 而不是自己调 FWorksheet.deleteRows，确保选区处理 / 合并单元格 / 撤销栈与右键一致。
            try {
              if (!injector) return false
              const cmdSvc = injector.get(ICommandService) as {
                executeCommand: (id: string, params?: object) => unknown
              }
              cmdSvc.executeCommand('sheet.command.remove-row', {
                range: { startRow, startColumn: 0, endRow, endColumn: COL_COUNT - 1 },
              })
            } catch {
              return false
            }
            markDirty()
            scheduleAutoSave()
            return true
          },
        }
        const isFullRowSelected = (): boolean => {
          // 编辑态（焦点在 Univer 单元格编辑器内）→ 放行，仅删除字符
          const ae = document.activeElement as HTMLElement | null
          if (ae && (ae.isContentEditable || ae.tagName === 'INPUT' || ae.tagName === 'TEXTAREA')) return false
          const ws = univerRef.current?.getActiveWorkbook()?.getActiveSheet()
          if (!ws) return false
          const ar = ws.getActiveRange()
          if (!ar) return false
          const startRow = ar.getRow()
          const startCol = ar.getColumn()
          const endCol = ar.getLastColumn()
          return startRow >= 1 && startCol === 0 && endCol >= COL_COUNT - 1
        }
        if (deleteCmdSvc && shortcutSvc) {
          rowDeleteDisposables.push(deleteCmdSvc.registerCommand(DeleteSelectedRowsCommand))
          // 优先级高于 Univer 原生「清除选区内容」快捷键（默认 0）：整行选中时本快捷键先命中并消费该按键。
          rowDeleteDisposables.push(
            shortcutSvc.registerShortcut({
              id: DeleteSelectedRowsCommand.id,
              binding: KeyCode.DELETE,
              priority: 100,
              preconditions: () => isFullRowSelected(),
            }),
          )
          rowDeleteDisposables.push(
            shortcutSvc.registerShortcut({
              id: DeleteSelectedRowsCommand.id,
              binding: KeyCode.BACKSPACE,
              priority: 100,
              preconditions: () => isFullRowSelected(),
            }),
          )
        }
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
        acRef.current?.dispose()
      } catch {
        /* noop */
      }
      acRef.current = null
      for (const d of rowDeleteDisposables) {
        try {
          d.dispose()
        } catch {
          /* noop */
        }
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
