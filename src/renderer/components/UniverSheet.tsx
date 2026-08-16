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
  COL_COUNT,
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

  const markDirty = () => {
    if (!dirtyRef.current) {
      dirtyRef.current = true
      setDirty(true)
    }
  }

  // 保存：取回 Univer 数据 → 还原为 string[][] → 交给 ExcelJS 写回（保留全部既有约定）
  const handleSave = async () => {
    if (saving) return
    const univerAPI = univerRef.current
    if (!univerAPI) return
    setSaving(true)
    setStatus('')
    try {
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
        setStatus(res.message || '已保存')
        onSaved()
      }
    } catch (err) {
      setStatus('保存失败：' + (err instanceof Error ? err.message : '未知错误'))
    } finally {
      setSaving(false)
    }
  }
  const saveRef = useRef(handleSave)
  saveRef.current = handleSave

  // 关闭文件：若有未保存修改，先自动保存（不弹确认框），保存成功后再关闭；
  // 若保存失败则保持面板打开，避免静默丢数据。
  const handleClose = async () => {
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
    // 但会改变数据，必须标记为“有修改”以便关闭时自动保存。
    const dispCmd = univerAPI.addEvent(univerAPI.Event.CommandExecuted, (e: unknown) => {
      const ev = e as { id?: string }
      if (
        ev && typeof ev.id === 'string' &&
        SHEET_STRUCT_MUTATION_KEYWORDS.some((k) => ev.id!.includes(k))
      ) {
        markDirty()
      }
    })

    // 编辑即标记脏；数量/单价变化自动重算金额（自动金额行），
    // 金额列被编辑时按“是否==数量×单价”重分类该行（手工金额不被覆盖）。
    const disp = univerAPI.addEvent(univerAPI.Event.SheetValueChanged, (e: unknown) => {
      markDirty()
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
        if (!touchesQP && !touchesAmt) continue
        for (let r = startRow; r <= rg.endRow; r++) {
          try {
            if (touchesQP) recomputeAmount(ws, r, autoAmountRows.current)
            // 金额列被改（含一次粘贴覆盖到金额）：按“当前值==数量×单价”重新归类
            if (touchesAmt) onAmountChanged(ws, r, autoAmountRows.current)
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
      } catch (e) {
        if (!disposed) setStatus('打开失败：' + (e instanceof Error ? e.message : '未知错误'))
      }
    })()

    return () => {
      disposed = true
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
      {status && <div className="memo-status">{status}</div>}
    </div>
  )
}
