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

/* 参与「最近输入」记忆的文本列（数字列不参与） */
const TEXT_FIELDS: (keyof BillRow)[] = ['no', 'date', 'name', 'unit', 'person', 'remark']
const HISTORY_KEY = 'bill-explorer:memo-history:v1'
const HISTORY_CAP = 20

type ColHistory = Record<string, string[]>

function loadHistory(): ColHistory {
  try {
    const raw = localStorage.getItem(HISTORY_KEY)
    if (!raw) return {}
    const parsed = JSON.parse(raw)
    return parsed && typeof parsed === 'object' ? (parsed as ColHistory) : {}
  } catch {
    return {}
  }
}

function saveHistory(h: ColHistory) {
  try {
    localStorage.setItem(HISTORY_KEY, JSON.stringify(h))
  } catch {
    /* ignore */
  }
}

function recordOne(h: ColHistory, field: string, value: string): ColHistory {
  const v = value.trim()
  if (!v) return h
  const next = [v, ...(h[field] || []).filter((x) => x !== v)].slice(0, HISTORY_CAP)
  return { ...h, [field]: next }
}

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

  /* ---- 最近输入记忆（Excel 式快捷输入） ---- */
  const [colHistory, setColHistory] = useState<ColHistory>(() => loadHistory())
  const historyRef = useRef(colHistory)

  function updateHistory(field: string, value: string) {
    const v = value.trim()
    if (!v) return
    historyRef.current = recordOne(historyRef.current, field, v)
    setColHistory(historyRef.current)
    saveHistory(historyRef.current)
  }

  // 把当前面板所有文本列的值批量记入历史（保存 / 关闭前兜底）
  function recordAllRows() {
    let h = historyRef.current
    let changed = false
    for (const r of rows) {
      for (const f of TEXT_FIELDS) {
        const v = String((r as unknown as Record<string, unknown>)[f] || '').trim()
        if (!v) continue
        const nh = recordOne(h, f, v)
        if (nh !== h) {
          h = nh
          changed = true
        }
      }
    }
    if (changed) {
      historyRef.current = h
      setColHistory(h)
      saveHistory(h)
    }
  }

  /* ---- 单元格快捷输入下拉 ---- */
  interface SuggestState {
    rowId: number
    colIdx: number
    field: string
    list: string[]
    highlight: number
    pos: { left: number; top: number; width: number }
  }
  const [suggest, setSuggest] = useState<SuggestState | null>(null)
  const suggestRef = useRef<SuggestState | null>(null)
  const openTimer = useRef<ReturnType<typeof setTimeout> | null>(null)

  function closeSuggest() {
    if (openTimer.current) {
      clearTimeout(openTimer.current)
      openTimer.current = null
    }
    suggestRef.current = null
    setSuggest(null)
  }

  function buildSuggest(rowId: number, colIdx: number, field: string, filter: string): SuggestState | null {
    const all = historyRef.current[field] || []
    const f = filter.trim()
    const list = f ? all.filter((x) => x.toLowerCase().includes(f.toLowerCase())) : all
    if (list.length === 0) return null
    const input = fieldRefs.current.get(rowId)?.[colIdx]
    if (!input) return null
    const rect = input.getBoundingClientRect()
    // 预估下拉高度（每项约 22px + 标题栏）
    const dropH = Math.min(220, list.length * 22 + 30)
    const belowSpace = window.innerHeight - rect.bottom - 2
    const top = belowSpace >= dropH || rect.top - dropH - 2 < 0
      ? rect.bottom + 2
      : rect.top - dropH - 2
    return {
      rowId,
      colIdx,
      field,
      list,
      highlight: f ? Math.max(0, list.findIndex((x) => x.toLowerCase().includes(f.toLowerCase()))) : 0,
      pos: { left: rect.left, top, width: Math.max(rect.width, 140) },
    }
  }

  function selectSuggestion(s: SuggestState, value: string) {
    setRows((prev) =>
      prev.map((r) => {
        if (r.id !== s.rowId) return r
        const next = { ...r }
        ;(next as unknown as Record<string, string | number | ''>)[s.field] = value
        return next
      }),
    )
    updateHistory(s.field, value)
    const input = fieldRefs.current.get(s.rowId)?.[s.colIdx]
    closeSuggest()
    if (input) input.focus()
  }

  // 点击下拉外部（非输入框、非下拉框）时关闭
  useEffect(() => {
    const onDocMouseDown = (e: MouseEvent) => {
      const t = e.target as HTMLElement | null
      if (!t) return
      if (t.closest('.memo-suggest')) return
      if (t.classList.contains('memo-input-inline')) return
      closeSuggest()
    }
    document.addEventListener('mousedown', onDocMouseDown)
    return () => document.removeEventListener('mousedown', onDocMouseDown)
  }, [])

  // 高亮项滚动到可见区域
  useEffect(() => {
    if (!suggest) return
    const el = document.querySelector('.memo-suggest li.is-hl')
    el?.scrollIntoView({ block: 'nearest' })
  }, [suggest?.highlight])

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
          recordAllRows()
          if (e.shiftKey) void onSaveAndNext(rows)
          else void onSave(rows)
          return
        }
      }
      if (e.key === 'Escape') {
        // 下拉打开时 Esc 仅关闭下拉，由单元格层拦截（stopPropagation），这里作防御
        if (suggestRef.current) return
        recordAllRows()
        onClose()
      }
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
    const row = rows[rowIdx]

    /* ===== 最近输入下拉：优先处理 ===== */
    const openSg = suggestRef.current
    const onThisCell = !!openSg && openSg.rowId === rowId && openSg.colIdx === colIdx

    // Alt+↓ / Alt+↑：手动弹出最近输入（已填内容也可换）
    if (e.altKey && (e.key === 'ArrowDown' || e.key === 'ArrowUp')) {
      e.preventDefault()
      const field = COLUMNS[colIdx]?.field as string
      if (TEXT_FIELDS.includes(field as keyof BillRow)) {
        const cur = String((row as unknown as Record<string, unknown>)[field] || '')
        const sg = buildSuggest(rowId, colIdx, field, cur)
        if (sg) {
          suggestRef.current = sg
          setSuggest(sg)
        }
      }
      return
    }

    if (onThisCell) {
      if (e.key === 'ArrowDown') {
        e.preventDefault()
        const next = { ...openSg, highlight: Math.min(openSg.list.length - 1, openSg.highlight + 1) }
        suggestRef.current = next
        setSuggest(next)
        return
      }
      if (e.key === 'ArrowUp') {
        e.preventDefault()
        const next = { ...openSg, highlight: Math.max(0, openSg.highlight - 1) }
        suggestRef.current = next
        setSuggest(next)
        return
      }
      if (e.key === 'Enter' || e.key === 'Tab') {
        e.preventDefault()
        const item = openSg.list[openSg.highlight]
        if (item !== undefined) selectSuggestion(openSg, item)
        return
      }
      if (e.key === 'Escape') {
        e.preventDefault()
        e.stopPropagation()
        closeSuggest()
        return
      }
    }

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

  /* ===== 最近输入：聚焦 / 失焦 / 输入 ===== */
  function handleFocus(rowId: number, colIdx: number, field: string, value: string) {
    if (openTimer.current) {
      clearTimeout(openTimer.current)
      openTimer.current = null
    }
    if (!TEXT_FIELDS.includes(field as keyof BillRow)) return
    if (String(value || '').trim() !== '') return
    openTimer.current = setTimeout(() => {
      // 聚焦后延时弹出；期间若已输入内容则不弹
      const input = fieldRefs.current.get(rowId)?.[colIdx]
      if (input && input.value.trim() !== '') return
      const sg = buildSuggest(rowId, colIdx, field, '')
      if (sg) {
        suggestRef.current = sg
        setSuggest(sg)
      }
    }, 120)
  }

  function handleBlur(field: string, value: string) {
    if (TEXT_FIELDS.includes(field as keyof BillRow)) updateHistory(field, value)
    closeSuggest()
  }

  function onCellInput(rowId: number, colIdx: number, field: string, rawValue: string) {
    const s = suggestRef.current
    if (!s || s.rowId !== rowId || s.colIdx !== colIdx) return
    const sg = buildSuggest(rowId, colIdx, field, rawValue)
    if (sg) {
      suggestRef.current = sg
      setSuggest(sg)
    } else {
      closeSuggest()
    }
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
                    onChange={(e) => {
                      onChange(r.id, c.field, e.target.value)
                      onCellInput(r.id, idx, c.field as string, e.target.value)
                    }}
                    onFocus={() =>
                      handleFocus(r.id, idx, c.field as string, String((r[c.field] as number | string) ?? ''))
                    }
                    onBlur={() =>
                      handleBlur(c.field as string, String((r[c.field] as number | string) ?? ''))
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
        <span className="memo-hint">Enter 下一格 · Ctrl+Enter 保存 · Ctrl+Shift+Enter 保存并下一条 · Alt+↓ 最近输入 · 拖表头调整列宽 · Esc 关闭</span>
        {status && (
          <span className={status.includes('失败') ? 'memo-status memo-error' : 'memo-status'}>
            {status}
          </span>
        )}
      </div>

      {suggest && (
        <div
          className="memo-suggest"
          style={{ left: suggest.pos.left, top: suggest.pos.top, width: suggest.pos.width }}
          onMouseDown={(e) => e.preventDefault()}
        >
          <div className="memo-suggest-title">最近输入 · Enter 选择 · Esc 关闭</div>
          <ul>
            {suggest.list.map((item, i) => (
              <li
                key={`${suggest.field}-${i}`}
                className={i === suggest.highlight ? 'is-hl' : ''}
                onMouseEnter={() => {
                  const next = { ...suggest, highlight: i }
                  suggestRef.current = next
                  setSuggest(next)
                }}
                onClick={() => selectSuggestion(suggest, item)}
              >
                {item}
              </li>
            ))}
          </ul>
        </div>
      )}
    </div>
  )
}
