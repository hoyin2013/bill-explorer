import { readFileSync, existsSync, accessSync, readdirSync, constants as FSConstants } from 'fs'
import { copyFile, mkdir, readdir, unlink } from 'fs/promises'
import { dirname, basename, extname, join } from 'path'
import * as ExcelJS from 'exceljs'
import { dialog } from 'electron'

// ============ 超大文件保护 ============
// 异常情况：不规范文件末尾大量空行、或在百万行位置残留一个孤立数字，
// 导致 ExcelJS 的 ws.rowCount 达到 100 万+ 行。打开时若逐行读取会实例化上百万行，
// 直接把程序卡死。超过此阈值即视为“异常超大”，直接拒绝打开并弹窗提示（不读取数据）。
const SAFE_MAX_ROWS = 100_000

// ============ 版本备份：保存前自动备份，便于“恢复上一个版本” ============
// 备份放在与文件同级的隐藏目录 `.billbackups` 下；扫描器跳过 `.` 前缀目录，不会被当成文件列出。
const BACKUP_DIR_NAME = '.billbackups'
// 每个文件最多保留的备份份数（超出删最旧的）
const MAX_BACKUPS = 5

// 备份文件名带时间戳（精确到毫秒，避免同一秒内多次保存时文件名撞车互相覆盖），
// 例如：账单.xlsx.20260812-101530123.bak
function backupFileName(filePath: string, date: Date): string {
  const base = basename(filePath)
  const pad = (n: number, w = 2) => String(n).padStart(w, '0')
  const ts = `${date.getFullYear()}${pad(date.getMonth() + 1)}${pad(date.getDate())}-${pad(date.getHours())}${pad(date.getMinutes())}${pad(date.getSeconds())}${pad(date.getMilliseconds(), 3)}`
  return `${base}.${ts}.bak`
}

// 写入前先备份当前文件（异步，避免大文件阻塞主线程）。备份失败不影响保存。
async function backupBeforeWrite(filePath: string): Promise<string | null> {
  if (!existsSync(filePath)) return null
  try {
    const dir = dirname(filePath)
    const bakDir = join(dir, BACKUP_DIR_NAME)
    await mkdir(bakDir, { recursive: true })
    const target = join(bakDir, backupFileName(filePath, new Date()))
    await copyFile(filePath, target)
    // 清理多余备份：按时间戳升序，删最旧的，保留最近 MAX_BACKUPS 份
    const base = basename(filePath)
    const list = (await readdir(bakDir))
      .filter((n) => n.startsWith(base + '.') && n.endsWith('.bak'))
      .sort() // 时间戳格式可字典序排序，最早在前
    while (list.length > MAX_BACKUPS) {
      const old = list.shift()!
      try { await unlink(join(bakDir, old)) } catch { /* ignore */ }
    }
    return target
  } catch (err) {
    // 备份失败不应阻断保存，仅记录
    console.warn('[backup] 创建备份失败：', err)
    return null
  }
}

// 列出某文件可用的备份（新 → 旧），time 形如 20260812-101530
export function listBackups(filePath: string): Array<{ path: string; time: string }> {
  const dir = dirname(filePath)
  const bakDir = join(dir, BACKUP_DIR_NAME)
  if (!existsSync(bakDir)) return []
  const base = basename(filePath)
  const list = readdirSync(bakDir)
    .filter((n) => n.startsWith(base + '.') && n.endsWith('.bak'))
    .sort()
    .reverse()
  return list.map((n) => {
    const ts = n.slice(base.length + 1, n.length - '.bak'.length)
    return { path: join(bakDir, n), time: ts }
  })
}

// 恢复最近一份备份：先备份“当前（可能已损坏）文件”，再把最新备份复制回去（恢复操作本身可逆）
export async function restoreBackup(filePath: string): Promise<{ error?: boolean; message?: string; backupTime?: string }> {
  const backups = listBackups(filePath)
  if (backups.length === 0) {
    return { error: true, message: '没有可恢复的备份版本（保存后才会生成备份）。' }
  }
  const latest = backups[0]
  try {
    await backupBeforeWrite(filePath)
    await copyFile(latest.path, filePath)
    return { message: `已恢复版本：${latest.time}`, backupTime: latest.time }
  } catch (err) {
    return { error: true, message: '恢复失败：' + (err instanceof Error ? err.message : '未知错误') }
  }
}

// ============ 预览：读取文件最近 N 行数据（不修改文件） ============
export interface PreviewCell {
  col: number
  value: string
  label?: string
}
export interface PreviewRow {
  cells: PreviewCell[]
}
export interface PreviewResult {
  error?: boolean
  message?: string
  sheetName?: string
  totalRows?: number     // 数据行总数（不含表头）
  headerLabels?: string[]
  rows?: PreviewRow[]
}

// ============ 追加写入返回类型 ============
export interface AppendMemoResult {
  error?: boolean
  message?: string
  rowNumber?: number
  count?: number   // 本次写入的行数
}

export async function previewRows(
  filePath: string,
  limit: number = 10,
): Promise<PreviewResult> {
  if (!existsSync(filePath)) {
    return { error: true, message: '文件不存在或已被删除。' }
  }
  const ext = extname(filePath).toLowerCase()
  if (ext !== '.xlsx') {
    return { error: true, message: '仅支持 .xlsx 文件预览。' }
  }

  const workbook = new ExcelJS.Workbook()
  try {
    const buffer = readFileSync(filePath)
    await workbook.xlsx.load(buffer)
  } catch {
    return { error: true, message: '无法读取文件，可能被 Excel 锁定或已损坏。' }
  }

  const ws = pickDataSheet(workbook)
  if (!ws || ws.rowCount === 0) {
    return { sheetName: ws?.name || '', totalRows: 0, headerLabels: [], rows: [] }
  }

  // 防卡死：异常超大文件直接跳过预览，避免从百万行自底向上扫描导致卡顿。
  if (ws.rowCount > SAFE_MAX_ROWS) {
    return {
      error: true,
      message: `文件过大（约 ${ws.rowCount.toLocaleString('zh-CN')} 行），疑似异常超大，已跳过预览以免卡顿。`,
      sheetName: ws.name || '',
      totalRows: 0,
      headerLabels: [],
      rows: [],
    }
  }

  // 表头：取第一行
  let headerLabels: string[] = []
  let dataStartRow = 2
  const headerRow = ws.getRow(1)
  const lastCol = ws.getColumn(ws.columnCount).number
  for (let c = 1; c <= lastCol; c++) {
    const v = headerRow.getCell(c).value
    headerLabels.push(v ? String(v) : '')
  }
  // 若首行也是数据（无表头），则从第 1 行开始
  if (headerLabels.every((h) => h === '')) {
    dataStartRow = 1
    headerLabels = Array.from({ length: lastCol }, (_, i) => `列${i + 1}`)
  }

  // 找最后一行有数据的行
  let lastDataRow = dataStartRow - 1
  for (let r = ws.rowCount; r >= dataStartRow; r -= 1) {
    let hasData = false
    for (let c = 1; c <= lastCol; c++) {
      const v = ws.getRow(r).getCell(c).value
      if (v !== undefined && v !== null && String(v).trim() !== '') {
        hasData = true
        break
      }
    }
    if (hasData) { lastDataRow = r; break }
  }

  const totalRows = lastDataRow - dataStartRow + 1
  // 取最近 N 行
  const startRow = Math.max(dataStartRow, lastDataRow - limit + 1)
  const previewRowsList: PreviewRow[] = []
  for (let r = startRow; r <= lastDataRow; r++) {
    const row = ws.getRow(r)
    const cells: PreviewCell[] = []
    for (let c = 1; c <= headerLabels.length; c++) {
      const v = row.getCell(c).value
      let s = ''
      if (v instanceof Date) s = v.toISOString().slice(0, 10)
      else if (v !== undefined && v !== null) s = String(v)
      cells.push({ col: c, value: s, label: headerLabels[c - 1] || undefined })
    }
    previewRowsList.push({ cells })
  }

  return {
    sheetName: ws.name || '',
    totalRows,
    headerLabels,
    rows: previewRowsList,
  }
}

// 固定账单表头（9 列，顺序即默认列顺序）—— 这是写入与匹配的源头
const HEADER = [
  { key: 'no',     label: '序号' },
  { key: 'date',   label: '日期' },
  { key: 'name',   label: '货品名称' },
  { key: 'unit',   label: '单位' },
  { key: 'qty',    label: '数量' },
  { key: 'price',  label: '单价' },
  { key: 'amount', label: '金额' },
  { key: 'person', label: '调货人' },
  { key: 'remark', label: '备注' },
] as const

export interface BillRecord {
  no: string
  date: string
  name: string
  unit: string
  qty: number
  price: number
  amount: number | ''
  person: string
  remark: string
}

// 在某一行里按标签找列号（ExcelJS 的 Row.findCell 在部分版本不可靠，这里直接遍历）
function findColByLabel(row: ExcelJS.Row, label: string): number {
  const last = Math.max(row.cellCount, HEADER.length)
  for (let c = 1; c <= last; c += 1) {
    const v = row.getCell(c).value
    if (v != null && String(v).trim() === label) return c
  }
  return 0
}

// 一行里至少有一个非空字段，才算有效记录（避免把空行写进 Excel）
function isEmptyRow(r: BillRecord): boolean {
  return !r.date && !r.name && !r.unit && !r.person && !r.remark && r.qty === 0 && r.price === 0 && !r.amount
}

// 选“数据所在工作表”：优先用首行匹配到 9 列账单表头最多的那张；
// 这样即使数据不在第一张表（如多了封面/目录表），也能正确读写，
// 避免只读 worksheets[0] 导致的漏读 / 误以为内容被清空。
function headerMatchScore(ws: ExcelJS.Worksheet): number {
  const headRow = ws.getRow(1)
  const labels = new Set(HEADER.map((h) => h.label))
  let score = 0
  const maxCol = Math.max(ws.columnCount, HEADER.length)
  for (let c = 1; c <= maxCol; c += 1) {
    const v = headRow.getCell(c).value
    if (v != null && labels.has(String(v).trim())) score += 1
  }
  return score
}
function pickDataSheet(workbook: ExcelJS.Workbook): ExcelJS.Worksheet | undefined {
  const sheets = workbook.worksheets
  if (sheets.length === 0) return undefined
  let best = sheets[0]
  let bestScore = headerMatchScore(best)
  for (let i = 1; i < sheets.length; i += 1) {
    const s = headerMatchScore(sheets[i])
    if (s > bestScore) {
      bestScore = s
      best = sheets[i]
    }
  }
  return best
}

// 标准日期字符串转 JS Date（Excel 原生日期），非法/空值原样返回。
// 关键：必须用 Date.UTC 构造"UTC 午夜"，否则用 new Date(y, m-1, d)（本地午夜）会被
// ExcelJS 按 UTC 偏移换算成带小数的序列号 —— 既差一天（如 8-15→8-14）又带时分秒。
// 同时容忍 `2026/8/15`、尾随时间（2026-08-15 14:30:00）、`2026年8月15日`、`20260815` 等写法，
// 一律只取"年月日"部分，保证存进 Excel 的是无时间的纯日期。
function toExcelDate(s: string): Date | string | undefined {
  if (!s) return undefined
  // 去掉尾随时间（14:30:00 / 9:05 等）与空白，只保留日期部分
  const t = String(s).trim().replace(/\s+\d{1,2}[:：]\d{1,2}([:：]\d{1,2})?$/, '')
  let m = /^(\d{4})[-/](\d{1,2})[-/](\d{1,2})$/.exec(t)
  if (!m) m = /^(\d{4})\D+(\d{1,2})\D+(\d{1,2})/.exec(t) // 中文 / 其它分隔符
  if (!m) m = /^(\d{4})(\d{2})(\d{2})$/.exec(t) // 紧凑 yyyymmdd
  if (!m) return s
  const y = +m[1]
  const mo = +m[2]
  const d = +m[3]
  const dt = new Date(Date.UTC(y, mo - 1, d))
  return isNaN(dt.getTime()) ? s : dt
}

// ============ 把多条账单记录一次性追加到 .xlsx 文件第一张表 ============
// 写入逻辑：
//   1. 优先匹配文件首行表头 → 各字段写到列名对应的列
//   2. 文件为空 → 自动补写表头，数据从第 2 行开始
//   3. 文件有数据但未识别表头 → 按默认列序(1..9)写入，并提示
//   4. 一次性 load 一次、writeFile 一次，避免逐行重复读写
// 金额留空(amount === '')时该单元格留空，不写入 0 或空字符串
function fillRecord(row: ExcelJS.Row, colMap: Record<string, number>, r: BillRecord) {
  row.getCell(colMap.no).value = r.no
  row.getCell(colMap.date).value = toExcelDate(r.date)
  row.getCell(colMap.date).numFmt = 'yyyy-mm-dd'
  row.getCell(colMap.name).value = r.name
  row.getCell(colMap.unit).value = r.unit
  row.getCell(colMap.qty).value = r.qty
  row.getCell(colMap.qty).numFmt = '#,##0'
  row.getCell(colMap.price).value = r.price
  row.getCell(colMap.price).numFmt = '#,##0'
  row.getCell(colMap.amount).numFmt = '#,##0'
  row.getCell(colMap.amount).value = (r.amount === '' || r.amount == null) ? null : r.amount
  row.getCell(colMap.person).value = r.person
  row.getCell(colMap.remark).value = r.remark
}
export async function appendRows(
  filePath: string,
  rows: BillRecord[],
): Promise<AppendMemoResult> {
  const ext = extname(filePath).toLowerCase()

  if (ext === '.xls') {
    return { error: true, message: '旧版 .xls 文件不支持写入。请另存为 .xlsx 后重试。' }
  }
  if (ext !== '.xlsx') {
    return { error: true, message: '仅支持 .xlsx 文件写入。' }
  }
  if (!existsSync(filePath)) {
    return { error: true, message: '文件不存在或已被删除。' }
  }

  try {
    accessSync(filePath, FSConstants.W_OK)
  } catch {
    return { error: true, message: '文件写权限不足。' }
  }

  const workbook = new ExcelJS.Workbook()
  try {
    // 必须先用 readFileSync 读成 Buffer 再加载，避开 jszip 路径读取的 EOCD bug
    const buffer = readFileSync(filePath)
    await workbook.xlsx.load(buffer)
  } catch (err) {
    return {
      error: true,
      message: '无法读取 Excel 文件，可能被 Excel 打开锁定或已损坏：' +
        (err instanceof Error ? err.message : '未知错误'),
    }
  }

  const ws = workbook.worksheets[0]
  if (!ws) {
    return { error: true, message: 'Excel 文件中没有找到工作表。' }
  }

  // ---- 表头列映射：把每个字段匹配到文件首行对应列 ----
  const colMap: Record<string, number> = {}
  if (ws.rowCount > 0) {
    const headerRow = ws.getRow(1)
    HEADER.forEach((f) => {
      const cell = headerRow.findCell(
        (c) => String(c.value).trim() === f.label,
      )
      if (cell) colMap[f.key] = cell.col
    })
  }
  HEADER.forEach((f, i) => {
    if (!(f.key in colMap)) colMap[f.key] = i + 1
  })
  const headerMatched = HEADER.some((f) => f.key in colMap)

  // ---- 确定写入起始行：从底往上找最后一行"有真实数据"的行，数据紧跟其后 ----
  // 不能直接用 ws.rowCount，它会把末尾空白行也算进去，导致数据被写到很靠下
  const maxCol = HEADER.length
  let lastDataRow = 0
  for (let r = ws.rowCount; r >= 1; r -= 1) {
    let hasData = false
    for (let c = 1; c <= maxCol; c += 1) {
      const v = ws.getRow(r).getCell(c).value
      if (v !== undefined && v !== null && String(v).trim() !== '') {
        hasData = true
        break
      }
    }
    if (hasData) { lastDataRow = r; break }
  }

  let startRow: number
  if (lastDataRow === 0) {
    // 整表为空：第 1 行写表头，第 2 行写数据
    HEADER.forEach((f, i) => {
      const c = ws.getCell(1, i + 1)
      c.value = f.label
      c.font = { bold: true }
    })
    startRow = 2
  } else {
    startRow = lastDataRow + 1
  }

  // ---- 一次性写入所有非空行 ----
  const validRows = rows.filter((r) => !isEmptyRow(r))
  validRows.forEach((record, i) => {
    fillRecord(ws.getRow(startRow + i), colMap, record)
  })

  try {
    // 写入前先备份当前文件，便于“恢复上一个版本”
    await backupBeforeWrite(filePath)
    await workbook.xlsx.writeFile(filePath)
  } catch (err) {
    return {
      error: true,
      message: '写入文件失败，文件可能被 Excel 打开：' +
        (err instanceof Error ? err.message : '未知错误'),
    }
  }

  const endRow = startRow + validRows.length - 1
  const hint = !headerMatched && ws.rowCount > 1
    ? '（未识别到表头，按默认列序写入）'
    : ''
  return {
    rowNumber: startRow,
    count: validRows.length,
    message: `已保存 ${validRows.length} 行，写入第 ${startRow}${endRow !== startRow ? '–' + endRow : ''} 行${hint}`,
  }
}

// ============ 读取整张表（全量，供类 Excel 网格编辑使用） ============
// 返回固定 9 列表头 + 全部数据行（字符串矩阵）。
// - 末行空行（尾部空行）自动跳过，避免打开时带着一堆空行
// - 日期单元格统一转成 yyyy-mm-dd，方便网格里的日期控件解析
// - 数字去掉无意义的尾随 .0
export interface SheetData {
  error?: boolean
  message?: string
  sheetName?: string
  headerLabels: string[]   // 固定 9 列：序号/日期/货品名称/单位/数量/单价/金额/调货人/备注
  rows: string[][]          // 数据行，每行长度 = headerLabels.length
}

// 把 Date 格式化为 yyyy-mm-dd。
// 写入时用 Date.UTC 午夜，ExcelJS 读回也是 UTC 基准的 Date，故这里用 UTC getter 取年月日，
// 与写入端保持一致，保证"存进去 8-15、读出来也是 8-15"往返无偏差。
function fmtDateLocal(d: Date): string {
  const y = d.getUTCFullYear()
  const m = String(d.getUTCMonth() + 1).padStart(2, '0')
  const day = String(d.getUTCDate()).padStart(2, '0')
  return `${y}-${m}-${day}`
}

// 把 Excel 单元格值转成可编辑的纯文本
function cellToText(value: unknown): string {
  if (value === undefined || value === null) return ''
  if (value instanceof Date) return fmtDateLocal(value)
  if (typeof value === 'object') {
    const v = value as Record<string, unknown>
    if ('result' in v && v.result !== undefined && v.result !== null) {
      return cellToText(v.result)
    }
    if ('text' in v && v.text != null) return String(v.text)
    if ('richText' in v && Array.isArray(v.richText)) {
      return (v.richText as Array<{ text?: string }>)
        .map((t) => t.text || '')
        .join('')
    }
    return ''
  }
  if (typeof value === 'number') {
    // 去掉尾随 .0：1000 -> "1000"，12.5 -> "12.5"
    return String(Number(value))
  }
  return String(value)
}

export async function loadSheet(filePath: string): Promise<SheetData> {
  const ext = extname(filePath).toLowerCase()
  if (ext !== '.xlsx') {
    return { error: true, message: '仅支持 .xlsx 文件。', headerLabels: [], rows: [] }
  }
  if (!existsSync(filePath)) {
    return { error: true, message: '文件不存在或已被删除。', headerLabels: [], rows: [] }
  }

  let workbook: ExcelJS.Workbook
  try {
    const buffer = readFileSync(filePath)
    workbook = new ExcelJS.Workbook()
    await workbook.xlsx.load(buffer)
  } catch (err) {
    return {
      error: true,
      message: '无法读取 Excel 文件，可能被 Excel 打开锁定或已损坏：' +
        (err instanceof Error ? err.message : '未知错误'),
      headerLabels: [],
      rows: [],
    }
  }

  const ws = pickDataSheet(workbook)
  if (!ws) {
    return { error: true, message: 'Excel 文件中没有找到工作表。', headerLabels: [], rows: [] }
  }

  // 防卡死：异常超大文件（多为末尾空行 + 某行一个孤立数据，rowCount 达百万级）直接拒绝打开，
  // 避免逐行读取百万行把程序卡死。仅用 O(1) 的 rowCount 判断，不触碰任何单元格。
  if (ws.rowCount > SAFE_MAX_ROWS) {
    const rc = ws.rowCount
    const msg =
      `文件「${basename(filePath)}」疑似异常超大（约 ${rc.toLocaleString('zh-CN')} 行）。\n\n` +
      `这类文件通常不规范：末尾存在大量空行，或在百万行位置残留孤立数据。若直接打开，程序会逐行读取上百万行而卡死。\n\n` +
      `已阻止打开。请先在 Excel 中清理多余行与格式、删除孤立数据后另存，再重新打开本程序。`
    try { dialog.showErrorBox('文件过大，已阻止打开', msg) } catch { /* 弹窗失败不影响返回错误 */ }
    return {
      error: true,
      message: `文件过大（约 ${rc.toLocaleString('zh-CN')} 行），已阻止打开以免程序卡死。请在 Excel 中清理多余行后重试。`,
      headerLabels: [],
      rows: [],
    }
  }

  const headerLabels = HEADER.map((h) => h.label)

  // 表头列映射：首行匹配标签；未匹配则回退到默认位置 i+1
  const colMap: Record<string, number> = {}
  const headerRow = ws.getRow(1)
  HEADER.forEach((f) => {
    const col = findColByLabel(headerRow, f.label)
    if (col) colMap[f.key] = col
  })
  HEADER.forEach((f, i) => {
    if (!(f.key in colMap)) colMap[f.key] = i + 1
  })

  // 数据起始行：首行若存在表头 → 第 2 行；否则第 1 行就是数据
  const headerMatched = HEADER.some((f) => findColByLabel(headerRow, f.label) !== 0)
  const dataStart = headerMatched ? 2 : 1

  // 单次遍历读取数据：用 eachRow 只迭代「有值」的行（自动跳过尾部 / 中间的空行），
  // 一边读一边记录最后一行有数据的位置，省去原来「自底向上扫描找末行 + 自顶向下再读一遍」两遍遍历。
  // 对末尾存在大量空行的不规范文件，原写法要 O(rowCount × 9) 逐个单元格判断，
  // 现在 eachRow 对空行只做一次 hasValues 判空，几乎零成本，打开速度显著提升。
  // 数据区内空行的不可见占位符（与 saveSheet 对应）
  const PLACEHOLDER = '\u200b'
  const rows: string[][] = []
  let lastDataRow = dataStart - 1
  ws.eachRow((row, rowNumber) => {
    if (rowNumber < dataStart) return // 跳过表头行
    const arr: string[] = []
    HEADER.forEach((f) => {
      arr.push(cellToText(row.getCell(colMap[f.key]).value))
    })
    // 整行无任何非空白内容（例如仅残留格式/旧空格的行）→ 跳过，避免生成空行/脏行
    if (arr.every((t) => t.trim() === '')) return
    // 整行仅由占位符构成（数据区空行被 saveSheet 用 \u00a0 占位）→ 还原为真正的空行，
    // 使网格里表现为用户有意留的空行（且仍是“全空”语义，不影响导出/统计）。
    const onlyPlaceholder = arr.every((t) => t === PLACEHOLDER || t.trim() === '')
    rows.push(onlyPlaceholder ? HEADER.map(() => '') : arr)
    lastDataRow = rowNumber
  })

  return { sheetName: ws.name || '', headerLabels, rows }
}

// ============ 覆盖保存整张表（类 Excel 网格的"保存=更新"） ============
// 写入逻辑：
//   1. 先剥离传入数据里"尾部全空行"
//   2. 重写第 1 行表头（保证 9 列一致）
//   3. 清空原数据区（行 2 ~ 原 rowCount）后，按列类型写回全部数据
//   4. 原有尾部多余行/空行一并被清空（即处理了"Excel 尾部大量空行"问题）
export async function saveSheet(
  filePath: string,
  rowsIn: string[][],
): Promise<AppendMemoResult> {
  const ext = extname(filePath).toLowerCase()
  if (ext === '.xls') {
    return { error: true, message: '旧版 .xls 文件不支持写入。请另存为 .xlsx 后重试。' }
  }
  if (ext !== '.xlsx') {
    return { error: true, message: '仅支持 .xlsx 文件写入。' }
  }
  if (!existsSync(filePath)) {
    return { error: true, message: '文件不存在或已被删除。' }
  }
  try {
    accessSync(filePath, FSConstants.W_OK)
  } catch {
    return { error: true, message: '文件写权限不足。' }
  }

  // 剥离尾部全空行（调用方一般已处理，这里再兜底一次）
  const rows = rowsIn.map((r) => r.map((c) => (c == null ? '' : String(c))))
  while (rows.length && rows[rows.length - 1].every((c) => !c.trim())) {
    rows.pop()
  }

  let workbook: ExcelJS.Workbook
  try {
    const buffer = readFileSync(filePath)
    workbook = new ExcelJS.Workbook()
    await workbook.xlsx.load(buffer)
  } catch (err) {
    return {
      error: true,
      message: '无法读取 Excel 文件，可能被 Excel 打开锁定或已损坏：' +
        (err instanceof Error ? err.message : '未知错误'),
    }
  }

  // 选数据所在的工作表（按表头匹配，兼容数据不在第一张表的情况），
  // 直接在原表覆盖写入——不再 removeWorksheet + addWorksheet。
  // 旧写法会把新建表追加到末尾，导致其它表升到 worksheets[0]，
  // 而本应用只读第一张表，于是下次打开误以为“内容被清空”。
  const ws = pickDataSheet(workbook)
  if (!ws) {
    return { error: true, message: 'Excel 文件中没有找到工作表。' }
  }

  const startRow = 2
  const lastNeeded = startRow + rows.length - 1

  // 最后一个“有实质内容”的行（相对 rows 下标）。用于区分两类空行：
  //  - 数据区内的空行（夹在两条有数据行之间，用户有意留的分隔行）→ 写入不可见占位，使其被 Excel 保留
  //  - 末尾的空行（视为未使用的预分配行）→ 不加占位，按原逻辑剥离
  let lastContent = -1
  rows.forEach((arr, i) => {
    if (arr.some((c) => String(c).trim() !== '')) lastContent = i
  })

  // 写表头（强制 9 列一致）
  HEADER.forEach((f, i) => {
    const c = ws.getCell(1, i + 1)
    c.value = f.label
    c.font = { bold: true }
  })

  // 逐行写回，按列类型处理（日期/数字/文本）。
  // 对“数据区内”的全空行写入不可见占位符（不间断空格 \u00a0）：
  // ExcelJS 不会把“所有单元格为空”的行写入文件，导致用户有意留的空行在保存后丢失；
  // 用占位符让该行被写入，loadSheet 读回时再还原为真正的空行。
  rows.forEach((arr, i) => {
    const row = ws.getRow(startRow + i)
    const isEmpty = !arr.some((c) => String(c).trim() !== '')
    if (isEmpty) {
      if (i <= lastContent) {
        row.values = [] // 先清空（避免残留旧数据），再写占位
        row.getCell(1).value = '\u200b'
      }
      return
    }
    HEADER.forEach((f, ci) => {
      const cell = row.getCell(ci + 1)
      const raw = arr[ci] != null ? String(arr[ci]) : ''
      if (raw.trim() === '') {
        cell.value = null
        return
      }
      if (f.key === 'date') {
        const d = toExcelDate(raw)
        if (d instanceof Date) {
          cell.value = d
          cell.numFmt = 'yyyy-mm-dd'
        } else {
          cell.value = raw
        }
      } else if (f.key === 'qty' || f.key === 'price') {
        const n = Number(raw)
        cell.value = isNaN(n) ? raw : n
        if (f.key !== 'price') cell.numFmt = '#,##0'
      } else if (f.key === 'amount') {
        const n = Number(raw)
        cell.value = isNaN(n) ? raw : n
        cell.numFmt = '#,##0'
      } else {
        cell.value = raw
      }
    })
  })

  // 清空数据区之后的多余行（旧数据残留 / 尾部空行），避免文件尾部堆积旧数据。
  // 注意：ExcelJS 的 spliceRows 对“尾部行”无效（其删除分支要求删除区间之后还有行），
  // 因此这里直接把尾部行的单元格值置空——空行不会被写入文件，读回时 rowCount 自然收敛。
  for (let r = lastNeeded + 1; r <= ws.rowCount; r += 1) {
    ws.getRow(r).values = []
  }

  try {
    // 写入前先备份当前文件，便于“恢复上一个版本”
    await backupBeforeWrite(filePath)
    await workbook.xlsx.writeFile(filePath)
  } catch (err) {
    return {
      error: true,
      message: '写入文件失败，文件可能被 Excel 打开：' +
        (err instanceof Error ? err.message : '未知错误'),
    }
  }

  const endRow = startRow + rows.length - 1
  return {
    rowNumber: startRow,
    count: rows.length,
    message: `已保存 ${rows.length} 行，覆盖第 ${startRow}${endRow !== startRow ? '–' + endRow : ''} 行`,
  }
}

// ============ 覆盖更新：同一文件连续编辑（未切换其他文件）时覆盖上次写入的行 ============
// 从 prevStartRow 开始覆盖，写入新的有效行；若新行数少于上次则把多余旧行清空。
export async function updateRows(
  filePath: string,
  rows: BillRecord[],
  prevStartRow: number,
  prevCount: number,
): Promise<AppendMemoResult> {
  const ext = extname(filePath).toLowerCase()
  if (ext === '.xls') {
    return { error: true, message: '旧版 .xls 文件不支持写入。请另存为 .xlsx 后重试。' }
  }
  if (ext !== '.xlsx') {
    return { error: true, message: '仅支持 .xlsx 文件写入。' }
  }
  if (!existsSync(filePath)) {
    return { error: true, message: '文件不存在或已被删除。' }
  }

  try {
    accessSync(filePath, FSConstants.W_OK)
  } catch {
    return { error: true, message: '文件写权限不足。' }
  }

  const workbook = new ExcelJS.Workbook()
  try {
    const buffer = readFileSync(filePath)
    await workbook.xlsx.load(buffer)
  } catch (err) {
    return {
      error: true,
      message: '无法读取 Excel 文件，可能被 Excel 打开锁定或已损坏：' +
        (err instanceof Error ? err.message : '未知错误'),
    }
  }

  const ws = workbook.worksheets[0]
  if (!ws) {
    return { error: true, message: 'Excel 文件中没有找到工作表。' }
  }

  // ---- 表头列映射 ----
  const colMap: Record<string, number> = {}
  if (ws.rowCount > 0) {
    const headerRow = ws.getRow(1)
    HEADER.forEach((f) => {
      const cell = headerRow.findCell(
        (c) => String(c.value).trim() === f.label,
      )
      if (cell) colMap[f.key] = cell.col
    })
  }
  HEADER.forEach((f, i) => {
    if (!(f.key in colMap)) colMap[f.key] = i + 1
  })

  const validRows = rows.filter((r) => !isEmptyRow(r))
  const span = Math.max(prevCount, validRows.length)
  for (let i = 0; i < span; i += 1) {
    const row = ws.getRow(prevStartRow + i)
    if (i < validRows.length) {
      fillRecord(row, colMap, validRows[i])
    } else {
      // 本次未再写入的旧行 → 整行清空
      Object.values(colMap).forEach((c) => {
        row.getCell(c).value = null
      })
    }
  }

  try {
    // 写入前先备份当前文件，便于“恢复上一个版本”
    await backupBeforeWrite(filePath)
    await workbook.xlsx.writeFile(filePath)
  } catch (err) {
    return {
      error: true,
      message: '写入文件失败，文件可能被 Excel 打开：' +
        (err instanceof Error ? err.message : '未知错误'),
    }
  }

  const endRow = prevStartRow + validRows.length - 1
  return {
    rowNumber: prevStartRow,
    count: validRows.length,
    message: `已更新 ${validRows.length} 行，覆盖第 ${prevStartRow}${endRow !== prevStartRow ? '–' + endRow : ''} 行`,
  }
}
