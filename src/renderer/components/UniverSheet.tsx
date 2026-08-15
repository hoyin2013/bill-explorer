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
  COL_COUNT,
} from '../univerAdapter'

type UniverAPI = ReturnType<typeof createUniver>['univerAPI']

interface Props {
  file: FileEntry
  api: ElectronAPI
  onClose: () => void
  onSaved: () => void
}

export function UniverSheet({ file, api, onClose, onSaved }: Props) {
  const containerRef = useRef<HTMLDivElement>(null)
  const univerRef = useRef<UniverAPI | null>(null)
  const dirtyRef = useRef(false)
  const fileRef = useRef(file)
  fileRef.current = file

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

  const handleClose = () => {
    if (dirtyRef.current && !window.confirm('有未保存的修改，确定关闭吗？')) return
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

    // 编辑即标记脏
    const disp = univerAPI.addEvent(univerAPI.Event.SheetValueChanged, () => markDirty())

    ;(async () => {
      try {
        const res = await api.loadSheet(fileRef.current.filePath)
        if (disposed) return
        if (res.error) {
          setStatus('打开失败：' + (res.message || ''))
          return
        }
        const data = rowsToWorkbookData(res.rows)
        univerAPI.createWorkbook(data)
        // 激活单元格落在数据末尾空行，便于继续录入 / 默认「填入」位置
        const target = univerAPI.getActiveSheet()
        if (target) {
          const ws = target.worksheet
          const r = Math.max(1, Math.min(res.rows.length + 1, 100000))
          try {
            ws.setActiveRange(ws.getRange(r, 0))
          } catch {
            /* 激活位置失败不影响使用 */
          }
        }
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
        univerAPI.dispose()
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
      const matrix = list.map(mapRecognizedToRow)
      try {
        ws.getRange(startRow, 0, matrix.length, COL_COUNT).setValues(matrix)
        ws.setActiveRange(ws.getRange(startRow + matrix.length, 0))
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
