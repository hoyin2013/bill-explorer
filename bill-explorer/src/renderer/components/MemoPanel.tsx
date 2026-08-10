import { useEffect, useRef, useState } from 'react'
import { FileEntry, ElectronAPI } from '../types'

interface BillRow {
  id: number
  date: string
  name: string
  unit: string
  qty: number
  price: number
  amount: number
  person: string
  remark: string
}

const EMPTY: BillRow = {
  id: 0, date: '', name: '', unit: '', qty: 0, price: 0, amount: 0, person: '', remark: '',
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
  { field: 'date', label: '日期', width: 90 },
  { field: 'name', label: '货品名称', width: 160 },
  { field: 'unit', label: '单位', width: 55 },
  { field: 'qty', label: '数量', width: 65 },
  { field: 'price', label: '单价', width: 80 },
  { field: 'amount', label: '金额', width: 90 },
  { field: 'person', label: '调货人', width: 90 },
  { field: 'remark', label: '备注', width: 130 },
]

/* 可编辑字段（金额只读、由数量×单价自动计算） */
const EDITABLE_FIELDS: (keyof BillRow)[] = ['date', 'name', 'unit', 'qty', 'price', 'person', 'remark']

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
      pendingFocus.current = { rowId: rows[0].id, colIdx: 1 }
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

  function handleCellKey(e: React.KeyboardEvent<HTMLInputElement>, rowId: number, colIdx: number) {
    if (e.key === 'Enter') {
      e.preventDefault()
      const rowIdx = rows.findIndex((r) => r.id === rowId)
      const lastEditableIdx = EDITABLE_FIELDS.length - 1
      if (colIdx <= lastEditableIdx) {
        if (colIdx < lastEditableIdx) {
          const target = fieldRefs.current.get(rowId)?.[colIdx + 1]
          if (target) target.focus()
        } else {
          const nextRow = rows[rowIdx + 1]
          if (nextRow) {
            const target = fieldRefs.current.get(nextRow.id)?.[colIdx]
            if (target) target.focus()
          } else {
            const newId = _nextId++
            setRows((prev) => [...prev, { ...EMPTY, id: newId }])
            pendingFocus.current = { rowId: newId, colIdx }
          }
        }
      }
    }
  }

  function onChange(id: number, field: keyof BillRow, value: string | number) {
    setRows((prev) =>
      prev.map((r) => {
        if (r.id !== id) return r
        const next = { ...r, [field]: value }
        if (!isNaN(next.qty) && !isNaN(next.price)) {
          next.amount = +(next.qty * next.price)
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
      (r) => r.date || r.name || r.unit || r.person || r.remark || r.qty !== 0 || r.price !== 0,
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
            <th style={{ width: 32, minWidth: 32 }}>
              <span>#</span>
            </th>
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
          {rows.map((r, rowIdx) => {
            let colIdx = 0
            const cells = COLUMNS.map((c) => {
              const idx = colIdx
              colIdx += 1
              const isEditable = EDITABLE_FIELDS.includes(c.field)
              const isAmount = c.field === 'amount'
              if (isAmount) {
                return (
                  <td key={c.field}>
                    <span className="memo-amount">{r.amount ? r.amount.toFixed(2) : '0.00'}</span>
                  </td>
                )
              }
              return (
                <td key={c.field} className={c.field === 'qty' || c.field === 'price' ? 'td-num' : ''}>
                  <input
                    ref={setRef(r.id, idx)}
                    className="memo-input memo-input-inline"
                    type={c.field === 'qty' ? 'number' : c.field === 'price' ? 'number' : 'text'}
                    min={c.field === 'qty' || c.field === 'price' ? '0' : undefined}
                    step={c.field === 'qty' ? '1' : c.field === 'price' ? '0.01' : undefined}
                    value={(r[c.field] as number | string) || ''}
                    onChange={(e) =>
                      onChange(
                        r.id,
                        c.field,
                        c.field === 'qty' || c.field === 'price'
                          ? Number(e.target.value) || 0
                          : e.target.value,
                      )
                    }
                    onKeyDown={(e) => handleCellKey(e, r.id, idx)}
                  />
                </td>
              )
            })
            return (
              <tr key={r.id} className="memo-row">
                <td className="row-no">{rowIdx + 1}</td>
                {cells}
              </tr>
            )
          })}
        </tbody>
        <tfoot>
          <tr>
            <td></td>
            <td colSpan={4} className="memo-total-label">合计金额</td>
            <td><span className="memo-total">{totalAmount.toFixed(2)}</span></td>
            <td colSpan={3}></td>
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
