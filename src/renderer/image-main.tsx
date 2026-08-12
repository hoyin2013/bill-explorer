import { StrictMode } from 'react'
import { createRoot } from 'react-dom/client'
import { ImageWindow } from './components/ImageWindow'
import { ElectronAPI } from './types'
import './index.css'

const api = window.electronAPI as ElectronAPI | undefined

function Root() {
  if (!api) {
    return <div style={{ padding: 20, color: '#f56c6c' }}>无法连接主进程（electronAPI 未加载）。</div>
  }
  return <ImageWindow api={api} />
}

createRoot(document.getElementById('root')!).render(
  <StrictMode>
    <Root />
  </StrictMode>,
)
