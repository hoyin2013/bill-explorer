import type { ReactNode } from 'react'

interface Props {
  workDir: string
  status: string
  loading: boolean
  onChoose: () => void
  onRefresh: () => void
  onOpenDir?: (dirPath: string) => void
}

export function DirectoryBar({ workDir, status, loading, onChoose, onRefresh, onOpenDir }: Props) {
  const dirPath = (
    <div
      className="dir-path"
      title={workDir ? '双击打开当前工作目录' : undefined}
      onDoubleClick={() => workDir && onOpenDir?.(workDir)}
    >
      {workDir ? (
        <>
          <span className="dir-label">当前工作目录：</span>
          <span className="dir-value" style={{ cursor: workDir ? 'pointer' : undefined }}>
            {workDir}
          </span>
        </>
      ) : (
        <span className="dir-hint">请点击上方按钮选择账单根目录</span>
      )}
      <span className="dir-status">{status}</span>
    </div>
  )
  return (
    <div className="dir-bar">
      <button className="btn" onClick={onChoose}>
        选择文件夹
      </button>
      <button
        className="btn"
        onClick={onRefresh}
        disabled={!workDir || loading}
      >
        刷新
      </button>
      {dirPath}
    </div>
  )
}
