import { useEffect, useState } from 'react'
import { ElectronAPI, AppSettings } from '../types'

interface Props {
  api: ElectronAPI
  open: boolean
  onClose: () => void
  onSaved?: () => void
}

export function SettingsModal({ api, open, onClose, onSaved }: Props) {
  const [settings, setSettings] = useState<AppSettings>({
    aiConfig: { baseURL: '', apiKey: '', model: 'gpt-4o-mini', temperature: 0.2 },
    imageDir: '',
    prompt: '',
  })
  const [loading, setLoading] = useState(false)
  const [status, setStatus] = useState('')

  useEffect(() => {
    if (!open) return
    setStatus('')
    setLoading(true)
    api.getSettings()
      .then((s) => {
        setSettings({
          aiConfig: { ...s.aiConfig },
          imageDir: s.imageDir || '',
          prompt: s.prompt || '',
        })
      })
      .catch((err) => setStatus('加载设置失败：' + (err instanceof Error ? err.message : '')))
      .finally(() => setLoading(false))
  }, [api, open])

  if (!open) return null

  async function onChooseImageDir() {
    const dir = await api.selectImageDirectory()
    if (dir) {
      setSettings((prev) => ({ ...prev, imageDir: dir }))
    }
  }

  async function onSubmit(e: React.FormEvent) {
    e.preventDefault()
    setLoading(true)
    setStatus('')
    try {
      const res = await api.saveSettings({
        aiConfig: settings.aiConfig,
        imageDir: settings.imageDir,
        prompt: settings.prompt,
      })
      if (res.error) {
        setStatus('保存失败：' + (res.message || '未知错误'))
      } else {
        setStatus('已保存')
        onSaved?.()
        setTimeout(() => onClose(), 300)
      }
    } catch (err) {
      setStatus('保存失败：' + (err instanceof Error ? err.message : '未知错误'))
    } finally {
      setLoading(false)
    }
  }

  return (
    <div className="modal-overlay" onClick={(e) => e.target === e.currentTarget && onClose()}>
      <div className="modal-panel">
        <div className="modal-header">
          <h3>AI 与图片设置</h3>
          <button className="modal-close" type="button" onClick={onClose}>×</button>
        </div>
        <form onSubmit={onSubmit}>
          <div className="modal-body">
            <div className="form-section">
              <div className="form-title">AI 接口配置（OpenAI 兼容）</div>
              <label className="form-row">
                <span>Base URL</span>
                <input
                  type="text"
                  value={settings.aiConfig.baseURL}
                  placeholder="https://api.openai.com/v1"
                  onChange={(e) => setSettings((prev) => ({
                    ...prev,
                    aiConfig: { ...prev.aiConfig, baseURL: e.target.value },
                  }))}
                />
              </label>
              <label className="form-row">
                <span>API Key</span>
                <input
                  type="password"
                  value={settings.aiConfig.apiKey}
                  placeholder="sk-..."
                  onChange={(e) => setSettings((prev) => ({
                    ...prev,
                    aiConfig: { ...prev.aiConfig, apiKey: e.target.value },
                  }))}
                />
              </label>
              <label className="form-row">
                <span>模型</span>
                <input
                  type="text"
                  value={settings.aiConfig.model}
                  placeholder="gpt-4o-mini"
                  onChange={(e) => setSettings((prev) => ({
                    ...prev,
                    aiConfig: { ...prev.aiConfig, model: e.target.value },
                  }))}
                />
              </label>
              <label className="form-row">
                <span>Temperature ({settings.aiConfig.temperature})</span>
                <input
                  type="range"
                  min={0}
                  max={1}
                  step={0.1}
                  value={settings.aiConfig.temperature}
                  onChange={(e) => setSettings((prev) => ({
                    ...prev,
                    aiConfig: { ...prev.aiConfig, temperature: Number(e.target.value) },
                  }))}
                />
              </label>
            </div>

            <div className="form-section">
              <div className="form-title">图片目录</div>
              <div className="form-row dir-row">
                <input
                  type="text"
                  readOnly
                  value={settings.imageDir || '未设置'}
                  className="dir-input"
                />
                <button type="button" className="btn" onClick={onChooseImageDir}>
                  选择目录
                </button>
              </div>
            </div>

            <div className="form-section">
              <div className="form-title">识别提示词</div>
              <textarea
                className="prompt-textarea"
                rows={10}
                value={settings.prompt}
                onChange={(e) => setSettings((prev) => ({ ...prev, prompt: e.target.value }))}
              />
              <div className="form-hint">提示词会随图片一起发给 AI，要求输出 JSON 数组。</div>
            </div>

            {status && (
              <div className={status.includes('失败') ? 'modal-status error' : 'modal-status'}>
                {status}
              </div>
            )}
          </div>
          <div className="modal-footer">
            <button type="button" className="btn btn-outline" onClick={onClose} disabled={loading}>
              取消
            </button>
            <button type="submit" className="btn btn-primary" disabled={loading}>
              {loading ? '保存中…' : '保存'}
            </button>
          </div>
        </form>
      </div>
    </div>
  )
}
