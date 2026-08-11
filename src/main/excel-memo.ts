import { readFileSync, existsSync, accessSync, constants as FSConstants } from 'fs'
import { extname } from 'path'
import * as ExcelJS from 'exceljs'

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

  const ws = workbook.worksheets[0]
  if (!ws || ws.rowCount === 0) {
    return { sheetName: ws?.name || '', totalRows: 0, headerLabels: [], rows: [] }
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

// 标准日期字符串转 JS Date（Excel 原生日期），非法/空值原样返回
function toExcelDate(s: string): Date | string | undefined {
  if (!s) return undefined
  const m = /^\d{4}-\d{2}-\d{2}$/.exec(s)
  if (!m) return s
  const [y, mo, d] = s.split('-').map(Number)
  const dt = new Date(y, mo - 1, d)
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

// 把 Date 按"本地时区"格式化为 yyyy-mm-dd（避免 toISOString 的 UTC 偏移导致差一天）
function fmtDateLocal(d: Date): string {
  const y = d.getFullYear()
  const m = String(d.getMonth() + 1).padStart(2, '0')
  const day = String(d.getDate()).padStart(2, '0')
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

  const ws = workbook.worksheets[0]
  if (!ws) {
    return { error: true, message: 'Excel 文件中没有找到工作表。', headerLabels: [], rows: [] }
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

  // 末行有数据行（从底向上扫描，跳过尾部空行）
  let lastDataRow = dataStart - 1
  for (let r = ws.rowCount; r >= dataStart; r -= 1) {
    let has = false
    for (let c = 1; c <= HEADER.length; c += 1) {
      const v = ws.getRow(r).getCell(c).value
      if (v !== undefined && v !== null && String(v).trim() !== '') {
        has = true
        break
      }
    }
    if (has) {
      lastDataRow = r
      break
    }
  }

  const rows: string[][] = []
  for (let r = dataStart; r <= lastDataRow; r += 1) {
    const row = ws.getRow(r)
    const arr: string[] = []
    HEADER.forEach((f) => {
      arr.push(cellToText(row.getCell(colMap[f.key]).value))
    })
    rows.push(arr)
  }

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

  const firstWs = workbook.worksheets[0]
  if (!firstWs) {
    return { error: true, message: 'Excel 文件中没有找到工作表。' }
  }

  // 重建第一张工作表：删除后重新添加，可彻底丢弃尾部空行 / 旧数据，
  // 让保存后的文件 rowCount 等于实际行数（Excel 不再显示大片空白）。
  // 其它工作表原样保留。
  const sheetName = firstWs.name || 'Sheet1'
  workbook.removeWorksheet(firstWs.id)
  const ws = workbook.addWorksheet(sheetName)

  const startRow = 2
  const lastNeeded = startRow + rows.length - 1

  // 写表头（强制 9 列一致）
  HEADER.forEach((f, i) => {
    const c = ws.getCell(1, i + 1)
    c.value = f.label
    c.font = { bold: true }
  })

  // 逐行写回，按列类型处理（日期/数字/文本）
  rows.forEach((arr, i) => {
    const row = ws.getRow(startRow + i)
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

  try {
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
