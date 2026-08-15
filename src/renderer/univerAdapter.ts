// Univer 与现有账单数据模型（string[][]）之间的适配器。
// 设计：ExcelJS 仍负责文件读写（保留日期UTC、空行占位、大文件拦截、自动备份等全部约定），
// Univer 只负责渲染与编辑。本文件把 ExcelJS 读出的数据灌进 Univer、把 Univer 的数据取回交给 ExcelJS 写回。
import type { IWorkbookData, IWorksheetData } from '@univerjs/core'
import type { AIRecognizedRow } from './types'

// 固定 9 列表头（与 ExcelJS 端 HEADER 一致）。首行作为 Univer 的第 0 行显示，
// 保存时由 workbookDataToRows 剥离，再由 saveSheet 重新写入真正的表头。
export const HEADER = ['序号', '日期', '货品名称', '单位', '数量', '单价', '金额', '调货人', '备注']
export const COL_COUNT = HEADER.length
const SHEET_ID = 'bill'
// 日期列（第 2 列，0 基索引 1）的样式 id。该列在 Univer 里按 yyyy-mm-dd 显示，
// 与 Excel 中 numFmt='yyyy-mm-dd' 保持一致 —— 这样编辑器里看到的就是打开 Excel 后看到的，所见即所得。
const DATE_STYLE = 'bill-date'

// Excel / Univer 日期序列号纪元：1899-12-30（与 ExcelJS 一致，2026 年这类远期日期无 1900 闰年 bug 干扰）
const EXCEL_EPOCH = Date.UTC(1899, 11, 30)

// 数字保留最多 2 位小数并去掉无意义尾随 0
function fmtNum(n: number): string {
  if (!isFinite(n)) return ''
  const r = Math.round(n * 100) / 100
  return String(r)
}

// 把各种来源的日期文本解析为 {y,m,d}（无法识别返回 null）。
// 支持：2026-8-11 / 2026/8/11 / 2026.8.11 / 20260811 / 2026年8月15日 / 8-11（补当年）/ Excel 日期序列号
function parseDateParts(input: string): { y: number; m: number; d: number } | null {
  const s = String(input ?? '').trim()
  if (!s) return null
  const make = (y: number, m: number, d: number): { y: number; m: number; d: number } | null => {
    if (y < 100) y += y < 70 ? 2000 : 1900
    if (m < 1 || m > 12 || d < 1 || d > 31) return null
    const dt = new Date(y, m - 1, d)
    if (dt.getFullYear() !== y || dt.getMonth() !== m - 1 || dt.getDate() !== d) return null
    return { y, m, d }
  }
  let m = s.match(/^(\d{4}|\d{2})\s*[-/.年]\s*(\d{1,2})\s*[-/.月]\s*(\d{1,2})\s*日?$/)
  if (m) return make(Number(m[1]), Number(m[2]), Number(m[3]))
  m = s.match(/^(\d{4})(\d{2})(\d{2})$/)
  if (m) return make(Number(m[1]), Number(m[2]), Number(m[3]))
  m = s.match(/^(\d{1,2})\s*[-/月]\s*(\d{1,2})\s*日?$/)
  if (m) return make(new Date().getFullYear(), Number(m[1]), Number(m[2]))
  if (/^\d{4,5}$/.test(s)) {
    const serial = Number(s)
    if (serial >= 1 && serial <= 60000) {
      const dt = new Date(EXCEL_EPOCH + Math.round(serial) * 86400000)
      return make(dt.getUTCFullYear(), dt.getUTCMonth() + 1, dt.getUTCDate())
    }
  }
  return null
}

// 把各种来源的日期文本规范化成 yyyy-mm-dd（无法识别则返回空串）
function parseDateText(input: string): string {
  const p = parseDateParts(input)
  if (!p) return ''
  const pad = (n: number) => String(n).padStart(2, '0')
  return `${p.y}-${pad(p.m)}-${pad(p.d)}`
}

// {y,m,d} → Excel/Univer 日期序列号（与 ExcelJS dateToSerial 同纪元，远期日期无 1900 闰年 bug）
function datePartsToSerial(p: { y: number; m: number; d: number }): number {
  return Math.round((Date.UTC(p.y, p.m - 1, p.d) - EXCEL_EPOCH) / 86400000)
}

// Excel/Univer 日期序列号 → yyyy-mm-dd
function serialToDateText(serial: number): string {
  const dt = new Date(EXCEL_EPOCH + Math.round(serial) * 86400000)
  const pad = (n: number) => String(n).padStart(2, '0')
  return `${dt.getUTCFullYear()}-${pad(dt.getUTCMonth() + 1)}-${pad(dt.getUTCDate())}`
}

// 日期列的取值归一化：序列号 → yyyy-mm-dd；文本 → 各种写法归一化；非日期原样保留。
// 这样无论 Univer 把“用户输入的日期”存成序列号还是文本，落盘到 Excel 的都是规整的 yyyy-mm-dd，
// 而 Univer 编辑区也因列 numFmt='yyyy-mm-dd' 显示成日期而非数字。
function normalizeDateValue(v: unknown): string {
  if (v == null) return ''
  if (typeof v === 'number') {
    if (v >= 1 && v <= 60000) return serialToDateText(v)
    return String(v)
  }
  const t = String(v).trim()
  if (!t) return ''
  return parseDateText(t) || t
}

// 一条 AI 识别结果 → 9 列字符串数组（复刻旧 SheetGrid.applyRecognizedRows 的映射）。
// 注意：调货人列（index 7）故意留空，AI 识别出的 person 仅用于「账单定位」，不写入该列。
export function mapRecognizedToRow(r: AIRecognizedRow): string[] {
  const q = Number(String(r.qty ?? '').replace(/,/g, '')) || 0
  const p = Number(String(r.price ?? '').replace(/,/g, '')) || 0
  const a = Number(String(r.amount ?? '').replace(/,/g, '')) || 0
  return [
    String(r.no ?? ''),
    parseDateText(String(r.date ?? '')),
    String(r.name ?? ''),
    String(r.unit ?? ''),
    q > 0 ? fmtNum(q) : '',
    p > 0 ? fmtNum(p) : '',
    a > 0 ? fmtNum(a) : '',
    '', // 调货人留空
    String(r.remark ?? ''),
  ]
}

// string[][]（纯数据行，不含表头）→ Univer IWorkbookData（首行为表头）
export function rowsToWorkbookData(rows: string[][]): Partial<IWorkbookData> {
  const all = [HEADER, ...rows]
  const cellData: Record<number, Record<number, { v: string | number; s?: string }>> = {}
  all.forEach((row, r) => {
    const rowObj: Record<number, { v: string | number; s?: string }> = {}
    for (let c = 0; c < COL_COUNT; c++) {
      const val = row[c] ?? ''
      if (c === 1) {
        // 日期列：能解析成日期就存成“序列号 + yyyy-mm-dd 样式”，
        // Univer 会按该格式显示成 2026-08-15 而非原始序列号（这是之前“输入日期变成数字”的根因：
        // Univer 自动把输入的日期识别成序列号，但列上没有 numFmt，于是显示成裸数字）。
        const parts = parseDateParts(val)
        if (parts) {
          rowObj[c] = { v: datePartsToSerial(parts), s: DATE_STYLE }
          continue
        }
      }
      if (val !== '') rowObj[c] = { v: val }
    }
    cellData[r] = rowObj
  })
  const sheet: Partial<IWorksheetData> = {
    id: SHEET_ID,
    name: '账单',
    // 多预留 100 行缓冲，便于继续录入 / OCR 填入，避免越界
    rowCount: all.length + 100,
    columnCount: COL_COUNT,
    cellData: cellData as unknown as IWorksheetData['cellData'],
    // 列级默认样式：第 2 列（0 基索引 1）整体按 yyyy-mm-dd 显示，
    // 新输入/粘贴的日期也会套用该格式，不会变成数字。
    columnData: { 1: { s: DATE_STYLE } },
  }
  return {
    id: SHEET_ID,
    name: '账单',
    styles: {
      [DATE_STYLE]: { n: { pattern: 'yyyy-mm-dd' } },
    },
    sheetOrder: [SHEET_ID],
    sheets: { [SHEET_ID]: sheet },
  }
}

// Univer IWorkbookData → string[][]（纯数据行；首行表头被剥离）
export function workbookDataToRows(wb: IWorkbookData): string[][] {
  const sheet = wb.sheets ? wb.sheets[SHEET_ID] ?? Object.values(wb.sheets)[0] : undefined
  const cd = sheet?.cellData as unknown as Record<number, Record<number, { v?: unknown }>> | undefined
  if (!cd) return []
  const rowKeys = Object.keys(cd)
    .map(Number)
    .sort((a, b) => a - b)
  let maxC = COL_COUNT - 1
  for (const r of rowKeys) {
    const rowObj = cd[r] || {}
    for (const c of Object.keys(rowObj)) {
      const ci = Number(c)
      if (ci > maxC) maxC = ci
    }
  }
  const out: string[][] = []
  for (const r of rowKeys) {
    if (r === 0) continue // 跳过表头行
    const rowObj = cd[r] || {}
    const arr: string[] = []
    for (let c = 0; c <= maxC; c++) {
      const cell = rowObj[c]
      if (c === 1) {
        // 日期列：序列号 → yyyy-mm-dd，文本写法归一化，保证落盘与编辑区一致
        arr.push(normalizeDateValue(cell && cell.v != null ? cell.v : ''))
      } else {
        arr.push(cell && cell.v != null ? String(cell.v) : '')
      }
    }
    out.push(arr)
  }
  return out
}
