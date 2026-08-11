import { useEffect, useRef, useState } from 'react'
import { FileEntry, ElectronAPI } from '../types'

interface BillRow {
  id: number
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

const EMPTY: BillRow = {
  id: 0, no: '', date: '', name: '', unit: '', qty: 0, price: 0, amount: '', person: '', remark: '',
}

interface Props {
  file: FileEntry
  api: ElectronAPI
  saving: boolean
  status: string
  onSave: (rows: BillRow[]) => void
  onSaveAndNext: (rows: BillRow[]) => void
  onClose: () => void
}

let _nextId = 1

/* 列定义：字段 → 默认宽度(px) */
interface ColDef {
  field: keyof BillRow
  label: string
  width: number
}
const COLUMNS: ColDef[] = [
  { field: 'no', label: '序号', width: 60 },
  { field: 'date', label: '日期', width: 90 },
  { field: 'name', label: '货品名称', width: 160 },
  { field: 'unit', label: '单位', width: 55 },
  { field: 'qty', label: '数量', width: 65 },
  { field: 'price', label: '单价', width: 80 },
  { field: 'amount', label: '金额', width: 90 },
  { field: 'person', label: '调货人', width: 90 },
  { field: 'remark', label: '备注', width: 130 },
]

/* 可编辑字段（金额可手工调整；数量×单价变化时仍会自动重算） */
const EDITABLE_FIELDS: (keyof BillRow)[] = ['no', 'date', 'name', 'unit', 'qty', 'price', 'amount', 'person', 'remark']

/* 数字字段：数量 / 单价 / 金额 */
const NUM_FIELDS: (keyof BillRow)[] = ['qty', 'price', 'amount']

const MIN_COL_WIDTH = 50

export function MemoPanel({
  file,
  api,
  saving,
  status,
  onSave,
  onSaveAndNext,
  onClose,
}: Props) {
  const [rows, setRows] = useState<BillRow[]>([
    { ...EMPTY, id: _nextId++ },
    { ...EMPTY, id: _nextId++ },
    { ...EMPTY, id: _nextId++ },
    { ...EMPTY, id: _nextId++ },
    { ...EMPTY, id: _nextId++ },
    { ...EMPTY, id: _nextId++ },
    { ...EMPTY, id: _nextId++ },
    { ...EMPTY, id: _nextId++ },
    { ...EMPTY, id: _nextId++ },
    { ...EMPTY, id: _nextId++ },
  ])

  /* 列宽（px）— 用户可拖动调整 */
  const [colWidths, setColWidths] = useState<number[]>(
    COLUMNS.map((c) => c.width),
  )

  // 每行每列的 input ref
  const fieldRefs = useRef<Map<number, HTMLInputElement[]>>(new Map())
  const pendingFocus = useRef<{ rowId: number; colIdx: number } | null>(null)

  useEffect(() => {
    const p = pendingFocus.current
    if (!p) return
    const target = fieldRefs.current.get(p.rowId)?.[p.colIdx]
    if (target) {
      pendingFocus.current = null
      target.focus()
    }
  }, [rows])

  useEffect(() => {
    if (rows[0]) {
      pendingFocus.current = { rowId: rows[0].id, colIdx: 0 }
    }
  }, [])

  useEffect(() => {
    const onKey = (e: KeyboardEvent) => {
      if (e.ctrlKey || e.metaKey) {
        if (e.key === 'Enter') {
          e.preventDefault()
          if (e.shiftKey) void onSaveAndNext(rows)
          else void onSave(rows)
          return
        }
      }
      if (e.key === 'Escape') onClose()
    }
    document.addEventListener('keydown', onKey)
    return () => document.removeEventListener('keydown', onKey)
  }, [rows, onSave, onSaveAndNext, onClose])

  /* ===== 列宽拖动 ===== */
  const resizeState = useRef<{
    colIdx: number
    startX: number
    startWidth: number
    moving: boolean
  } | null>(null)

  function onStartResize(colIdx: number, e: React.MouseEvent<HTMLDivElement>) {
    e.stopPropagation()
    e.preventDefault()
    const startX = e.clientX
    const startWidth = colWidths[colIdx]
    resizeState.current = { colIdx, startX, startWidth, moving: true }
    const onMouseMove = (me: MouseEvent) => {
      if (!resizeState.current) return
      const diff = me.clientX - startX
      const newWidth = Math.max(MIN_COL_WIDTH, startWidth + diff)
      setColWidths((prev) => {
        const next = [...prev]
        next[colIdx] = newWidth
        return next
      })
    }
    const onMouseUp = () => {
      resizeState.current = null
      window.removeEventListener('mousemove', onMouseMove)
      window.removeEventListener('mouseup', onMouseUp)
      document.body.style.cursor = ''
    }
    window.addEventListener('mousemove', onMouseMove)
    window.addEventListener('mouseup', onMouseUp)
    document.body.style.cursor = 'col-resize'
  }

  /* ===== 数据操作 ===== */
  function setRef(rowId: number, idx: number) {
    return (el: HTMLInputElement | null) => {
      if (!el) return
      let arr = fieldRefs.current.get(rowId)
      if (!arr) { arr = []; fieldRefs.current.set(rowId, arr) }
      arr[idx] = el
    }
  }

  // 键盘导航：方向键 / 回车在单元格间移动焦点
  //  - 回车 = 向右；到行尾自动换到下一行起始；到最后一行自动补行
  //  - 方向键 上下左右 在网格中移动，出界自动补行 / 环绕
  function handleCellKey(e: React.KeyboardEvent<HTMLInputElement>, rowId: number, colIdx: number) {
    const rowIdx = rows.findIndex((r) => r.id === rowId)
    if (rowIdx < 0) return

    let dRow = 0
    let dCol = 0
    switch (e.key) {
      case 'Enter':
      case 'ArrowRight':
        dCol = 1
        break
      case 'ArrowLeft':
        dCol = -1
        break
      case 'ArrowDown':
        dRow = 1
        break
      case 'ArrowUp':
        dRow = -1
        break
      default:
        return
    }
    e.preventDefault()

    const lastCol = EDITABLE_FIELDS.length - 1
    let r = rowIdx + dRow
    let c = colIdx + dCol

    // 列到边界水平环绕：超出最右列 → 回到第 0 列并进入下一行
    if (c > lastCol) {
      c = 0
      r += 1
    } else if (c < 0) {
      c = lastCol
      r -= 1
    }

    if (r < 0) r = 0

    if (r < rows.length) {
      const target = fieldRefs.current.get(rows[r].id)?.[c]
      if (target) {
        pendingFocus.current = null
        target.focus()
      }
    } else {
      // 超出最后一行：补足所需行数并聚焦目标格
      const addCount = r - rows.length + 1
      const newRows: BillRow[] = []
      for (let k = 0; k < addCount; k += 1) newRows.push({ ...EMPTY, id: _nextId++ })
      setRows((prev) => [...prev, ...newRows])
      pendingFocus.current = { rowId: newRows[addCount - 1].id, colIdx: c }
    }
  }

  function onChange(id: number, field: keyof BillRow, rawValue: string) {
    setRows((prev) =>
      prev.map((r) => {
        if (r.id !== id) return r
        const next = { ...r }
        if (field === 'amount') {
          // 金额留空 → 保持空（保存时留空），否则转数字
          next.amount = rawValue === '' ? '' : Number(rawValue)
        } else if (field === 'qty' || field === 'price') {
          next[field] = Number(rawValue) || 0
        } else {
          // 文本字段：no/date/name/unit/person/remark
          const text = next as unknown as Record<string, string>
          text[field] = rawValue
        }
        // 数量/单价变化时自动重算金额；金额手工留空则保留空
        if (field !== 'amount' && next.qty !== 0 && next.price !== 0) {
          next.amount = Number(next.qty) * Number(next.price)
        }
        return next
      }),
    )
  }

  function addRowWithId(id: number) {
    setRows((prev) => [...prev, { ...EMPTY, id }])
  }

  function removeRow(id: number) {
    setRows((prev) => (prev.length === 1 ? prev : prev.filter((r) => r.id !== id)))
  }

  function hasAnyData(): boolean {
    return rows.some(
      (r) => r.date || r.name || r.unit || r.person || r.remark || r.qty !== 0 || r.price !== 0 || (r.amount !== '' && r.amount !== 0),
    )
  }

  async function handleOpenInExcel() {
    await api.openFile(file.filePath)
  }

  const totalAmount = rows.reduce((s, r) => s + (r.amount || 0), 0)

  return (
    <div className="memo-panel">
      <div className="memo-header">
        <div className="memo-file">
          <span className="memo-file-name">{file.fileName}</span>
        </div>
        <div className="memo-header-actions">
          <button className="btn btn-outline btn-add" onClick={() => addRowWithId(_nextId++)}>
            + 行
          </button>
          <button className="btn btn-primary" disabled={saving || !hasAnyData()} onClick={() => onSave(rows)}>
            {saving ? '保存中...' : '保存 Ctrl+Enter'}
          </button>
          <button className="btn btn" disabled={saving || !hasAnyData()} onClick={() => onSaveAndNext(rows)}>
            保存并下一条
          </button>
          <button className="btn btn-outline" onClick={handleOpenInExcel} title="用 Excel 程序打开">
            Excel 打开
          </button>
          <button className="btn btn-small btn-close" onClick={onClose} title="Esc 关闭">
            关闭
          </button>
        </div>
      </div>

      <table className="memo-table">
        <thead>
          <tr>
            {COLUMNS.map((c, i) => (
              <th key={c.field} style={{ width: colWidths[i], minWidth: MIN_COL_WIDTH }}>
                <span>{c.label}</span>
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
          {rows.map((r) => {
            let colIdx = 0
            const cells = COLUMNS.map((c) => {
              const idx = colIdx
              colIdx += 1
              const isNum = NUM_FIELDS.includes(c.field)
              return (
                <td key={c.field} className={isNum ? 'td-num' : ''}>
                  <input
                    ref={setRef(r.id, idx)}
                    className="memo-input memo-input-inline"
                    type={isNum ? 'number' : 'text'}
                    min={isNum ? '0' : undefined}
                    step={isNum ? '1' : undefined}
                    value={(r[c.field] as number | string) || ''}
                    onChange={(e) =>
                      onChange(r.id, c.field, e.target.value)
                    }
                    onKeyDown={(e) => handleCellKey(e, r.id, idx)}
                  />
                </td>
              )
            })
            return (
              <tr key={r.id} className="memo-row">
                {cells}
              </tr>
            )
          })}
        </tbody>
        <tfoot>
          <tr>
            <td colSpan={6} className="memo-total-label">合计金额</td>
            <td><span className="memo-total">{totalAmount}</span></td>
            <td colSpan={2}></td>
          </tr>
        </tfoot>
      </table>

      <div className="memo-spacer" />

      <div className="memo-footer">
        <span className="memo-hint">Enter 下一格 · Ctrl+Enter 保存 · Ctrl+Shift+Enter 保存并下一条 · 拖表头调整列宽 · Esc 关闭</span>
        {status && (
          <span className={status.includes('失败') ? 'memo-status memo-error' : 'memo-status'}>
            {status}
          </span>
        )}
      </div>
    </div>
  )
}
