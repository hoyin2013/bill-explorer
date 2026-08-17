// 单元格输入记忆（自动补全）的纯函数：算法层。可在 Node 单测 / 浏览器渲染层复用。
// 不依赖任何 DOM / Univer API，输入数据 → 输出候选。
import { dateStrToSerial } from '../univerAdapter'

export const AC_MAX_SUGGESTIONS = 12
// 列范围参考（与 univerAdapter 的 HEADER 一致）；任何读 col 的纯函数都用这个做边界。
export const HEADER = ['序号', '日期', '货品名称', '单位', '数量', '单价', '金额', '调货人', '备注']
export const COL_COUNT = HEADER.length
// 日期列 = 第 2 列（0 基索引 1）。仅这一列会被序列号 <-> 日期 双向转换。
export const DATE_COL_INDEX = 1

export type AcItem = { display: string; raw: string | number }

export interface ColVal {
  value: string | number
  row: number
  freq: number
}

// 把任意 cell 值转成候选 (display, raw)：
// - 日期列：raw 必须是 Excel 序列号（写回单元格就是真日期，而非"看起来像日期的文本"）
// - 其他列：原样
export function toItem(raw: string | number, isDate: boolean): AcItem {
  if (isDate) {
    let num: number | null = null
    if (typeof raw === 'number') num = raw
    else {
      // 严格数字串 → 直接当序列号
      if (/^-?\d+(\.\d+)?$/.test(String(raw).trim())) num = Number(String(raw).trim())
      // 文本日期（"2026-08-16" / "2026/8/16" / "2026年8月16日" / "20260816"）→ 序列号
      else num = dateStrToSerial(String(raw))
    }
    if (num != null && isFinite(num)) {
      const d = serialToDateStr(num)
      return { display: d, raw: num }
    }
    return { display: String(raw), raw: String(raw) }
  }
  return { display: String(raw), raw }
}

// Excel 序列号 → yyyy-mm-dd。纪元 1899-12-30，与 ExcelJS dateToSerial 一致；用 UTC 避免时区偏移。
export function serialToDateStr(serial: number): string {
  if (!isFinite(serial)) return ''
  const ms = (serial - 25569) * 86400 * 1000
  const d = new Date(ms)
  const y = d.getUTCFullYear()
  const m = String(d.getUTCMonth() + 1).padStart(2, '0')
  const day = String(d.getUTCDate()).padStart(2, '0')
  return `${y}-${m}-${day}`
}

// 抽取 workbook snapshot 的某列所有值。
// 返回每个出现位置的条目（不聚合 row）：同名值在多处出现时，每个位置都贡献一条候选，
// 排序时由 computeSuggestions 用 "上方最近 (row<currentRow 优先 + row 越大越近)" 来定最终顺序。
// 注意：聚合 row 会出错 —— 比如 '苹果' 在 row=0 和 row=4 都被录入，若聚合 row 取 max=4，
// 用户在 row=2 输入时会被 row<2 过滤掉，根本看不到 '苹果'（真实 bug，已确认）。
export function extractColumnValues(
  wb: unknown,
  col: number,
  sheetName = 'bill',
): ColVal[] {
  const sheets = (wb as { sheets?: Record<string, { cellData?: Record<number, Record<number, { v?: unknown }>> }> })?.sheets
  const sheet = sheets ? (sheets[sheetName] ?? Object.values(sheets)[0]) : undefined
  const cd = sheet?.cellData
  if (!cd) return []
  // 按 col 聚合：value -> { rows: [], freq }
  // 这样同名值每出现一次都记录 row，排序时 row<currentRow 的同名值都参与。
  const agg = new Map<string, { rows: number[]; freq: number }>()
  for (const r of Object.keys(cd)) {
    const cell = cd[Number(r)]?.[col]
    if (cell && cell.v != null) {
      const s = String(cell.v).trim()
      if (s) {
        const prev = agg.get(s)
        if (prev) {
          prev.freq += 1
          prev.rows.push(Number(r))
        } else {
          agg.set(s, { rows: [Number(r)], freq: 1 })
        }
      }
    }
  }
  // 展开为每个 row 一条 ColVal（同名值在多个 row 多次参与）
  const out: ColVal[] = []
  for (const [value, { rows, freq }] of agg.entries()) {
    for (const row of rows) {
      out.push({ value, row, freq })
    }
  }
  return out
}

// 主算法：给 partial + col + currentRow，返回排序好的候选。
// 规则（与 Excel/WPS 自动完成一致）：
//   1. 优先 "以已输入内容开头" (大小写不敏感)
//   2. 没有开头匹配时退化为 "包含匹配"
//   3. 排序：① 上方最近优先  ② 频率高优先  ③ 字母序
//   4. 历史记录轻微加权（freq +2），常输项更突出
export function computeSuggestions(
  partial: string,
  col: number,
  currentRow: number,
  columnVals: ColVal[],
  history: string[],
): AcItem[] {
  const p = partial.trim().toLowerCase()
  if (!p) return []
  const isDate = col === DATE_COL_INDEX
  const colVals = columnVals.filter((x) => x.row < currentRow)
  // 按 display 聚合：同名值在多处出现时取 row 最大的条目代表（与 Excel "上方最近优先 row 越大越近" 一致），
  // 同时累加 freq —— 这样同名值即便出现在多个位置也只占一条候选，但频率权重正确（出现越多越排前）。
  const cands = new Map<string, AcItem & { freq: number; row: number }>()
  for (const c of colVals) {
    const it = toItem(c.value, isDate)
    const prev = cands.get(it.display)
    if (prev) {
      prev.freq += c.freq
      prev.row = Math.max(prev.row, c.row)
    } else {
      cands.set(it.display, { ...it, freq: c.freq, row: c.row })
    }
  }
  for (const h of new Set(history)) {
    const it = toItem(h, isDate)
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
    if (ra !== rb) return rb - ra
    if (a.freq !== b.freq) return b.freq - a.freq
    return a.display.localeCompare(b.display)
  })
  return matched.slice(0, AC_MAX_SUGGESTIONS)
}
