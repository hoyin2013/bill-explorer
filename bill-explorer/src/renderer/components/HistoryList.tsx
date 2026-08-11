import { useRef, useState } from 'react'
import { HistoryRecord } from '../types'

interface Props {
  records: HistoryRecord[]
  onSelect: (filePath: string) => void
  onClear: () => void
}

// 历史面板最小高度
const MIN_H = 100

export function HistoryList({ records, onSelect, onClear }: Props) {
  // 高度（px），可通过顶部把手拖动调整
  const [height, setHeight] = useState(200)
  const dragRef = useRef<{ startY: number; startH: number } | null>(null)

  // 拖动顶部把手，上下调节面板高度
  function onResizeDown(e: React.MouseEvent) {
    e.preventDefault()
    e.stopPropagation()
    dragRef.current = { startY: e.clientY, startH: height }
    const onMove = (ev: MouseEvent) => {
      const d = dragRef.current
      if (!d) return
      // 鼠标向上拉（负 dy）使面板更高，向下拉使其更矮
      const h = Math.max(MIN_H, d.startH + (d.startY - ev.clientY))
      setHeight(h)
    }
    const onUp = () => {
      dragRef.current = null
      window.removeEventListener('mousemove', onMove)
      window.removeEventListener('mouseup', onUp)
    }
    window.addEventListener('mousemove', onMove)
    window.addEventListener('mouseup', onUp)
  }

  return (
    <div className="history-panel" style={{ height }}>
      <div className="history-resize-handle" onMouseDown={onResizeDown} title="拖动调整高度">
        <span className="history-grab" />
      </div>
      <div className="history-header">
        <span>最近修改</span>
        {records.length > 0 && (
          <button className="history-clear" onClick={onClear} title="清空历史记录">
            清空
          </button>
        )}
      </div>
      {records.length === 0 ? (
        <div className="history-empty">暂无记录</div>
      ) : (
        <ul className="history-list">
          {records.map((h) => (
            <li
              key={h.filePath}
              className="history-item"
              onClick={() => onSelect(h.filePath)}
              title={h.filePath}
            >
              <div className="history-name">{h.fileName}</div>
              <div className="history-time">{h.time}</div>
            </li>
          ))}
        </ul>
      )}
    </div>
  )
}