import { useEffect, useState } from 'react'
import { ElectronAPI, AppSettings } from '../types'

interface Props {
  api: ElectronAPI
  open: boolean
  onClose: () => void
  onSaved?: () => void
}

// 常用视觉模型预设（一键切换）。本场景是「中文手写小票 OCR + 结构化 JSON 抽取」，
// 用专门的视觉模型比通用旗舰思考模型更便宜更快；qwen3.8-max 等思考模型已自动关闭思考。
const MODEL_PRESETS = [
  'qwen-vl-plus', // 推荐：便宜、快、中文 OCR 强，性价比最高
  'qwen3-vl-plus', // 视觉版，平衡效果与速度
  'qwen3-vl-flash', // 最快最便宜，量大大图首选
  'qwen3.8-max', // 当前常用旗舰（已自动关闭思考；仍偏慢偏贵）
  'qwen-max',
  'gpt-4o-mini',
]

export function SettingsModal({ api, open, onClose, onSaved }: Props) {
  const [settings, setSettings] = useState<AppSettings>({
    aiConfig: { baseURL: '', apiKey: '', model: 'gpt-4o-mini', temperature: 0.2, fastMode: true },
    imageDir: '',
    prompt: '',
    pythonPath: '',
    detectModel: '',
    enableDetect: true,
  })
  const [loading, setLoading] = useState(false)
  const [status, setStatus] = useState('')
  // 记录 mousedown 是否发生在遮罩层本身：只有「在遮罩上按下、且在遮罩上松开」才算真正点击遮罩关闭，
  // 避免在小票文本里拖拽选区（按下在文本框、松开在遮罩）被误判为点击遮罩而关闭整个设置界面。
  const [backdropMouseDown, setBackdropMouseDown] = useState(false)

  useEffect(() => {
    if (!open) return
    setStatus('')
    setLoading(true)
    Promise.all([
      api.getSettings().catch(() => null),
    ]).then(([s]) => {
      if (s) {
        setSettings({
          aiConfig: { ...s.aiConfig },
          imageDir: s.imageDir || '',
          prompt: s.prompt || '',
          pythonPath: s.pythonPath || '',
          detectModel: s.detectModel || '',
          enableDetect: s.enableDetect !== false,
        })
      } else {
        setStatus('加载设置失败')
      }
    }).finally(() => setLoading(false))
  }, [api, open])

  if (!open) return null

  function onBackdropMouseDown(e: React.MouseEvent) {
    // 仅在「按下点就是遮罩本身」时记为 true；在面板内按下则为 false
    setBackdropMouseDown(e.target === e.currentTarget)
  }
  function onBackdropClick(e: React.MouseEvent) {
    // 只有「按下在遮罩、且松开也在遮罩」才关闭，防止拖拽选区时误关
    if (backdropMouseDown && e.target === e.currentTarget) onClose()
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
        pythonPath: settings.pythonPath,
        detectModel: settings.detectModel,
        enableDetect: settings.enableDetect,
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
    <div className="modal-overlay" onMouseDown={onBackdropMouseDown} onClick={onBackdropClick}>
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
                <select
                  className="model-preset"
                  value={MODEL_PRESETS.includes(settings.aiConfig.model) ? settings.aiConfig.model : 'custom'}
                  onChange={(e) => {
                    const v = e.target.value
                    if (v !== 'custom') {
                      setSettings((prev) => ({ ...prev, aiConfig: { ...prev.aiConfig, model: v } }))
                    }
                  }}
                >
                  {MODEL_PRESETS.map((m) => (
                    <option key={m} value={m}>{m}</option>
                  ))}
                  <option value="custom">自定义…</option>
                </select>
              </label>
              <label className="form-row">
                <span>模型 ID</span>
                <input
                  type="text"
                  value={settings.aiConfig.model}
                  placeholder="例如 qwen-vl-plus"
                  onChange={(e) => setSettings((prev) => ({
                    ...prev,
                    aiConfig: { ...prev.aiConfig, model: e.target.value },
                  }))}
                />
              </label>
              <div className="form-hint">
                本场景（中文手写小票 OCR + 结构化抽取）推荐用专门视觉模型 <code>qwen-vl-plus</code> / <code>qwen3-vl-plus</code>：比 <code>qwen3.8-max</code> 等通用旗舰思考模型便宜且快。若仍用思考模型，已自动关闭思考以省 token。
              </div>
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
              <label className="form-row dir-row">
                <input
                  type="checkbox"
                  checked={settings.aiConfig.fastMode !== false}
                  onChange={(e) => setSettings((prev) => ({
                    ...prev,
                    aiConfig: { ...prev.aiConfig, fastMode: e.target.checked },
                  }))}
                />
                <span>快速模式（关闭模型思考/推理：OpenRouter 用 reasoning、阿里云百炼/Qwen 用 enable_thinking=false；显著加快响应并省 token；不支持的接口会自动忽略）</span>
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
                <button type="button" className="btn" onClick={onClose}>
                  前往设置
                </button>
              </div>
              <div className="form-hint">图片目录在右侧「小票识图」面板中点击「选择图片目录」设置，目录中非图片文件会自动过滤。</div>
            </div>

            <div className="form-section">
              <div className="form-title">识别提示词</div>
              <textarea
                className="prompt-textarea"
                rows={14}
                value={settings.prompt}
                onChange={(e) => setSettings((prev) => ({ ...prev, prompt: e.target.value }))}
              />
              <div className="form-hint">
                提示词会随图片一起发给 AI，要求其按新格式输出 JSON 数组（每张小票 = 一条记录，含 name/date/items）。
              </div>
            </div>

            <div className="form-section">
              <div className="form-title">小票检测增强（YOLOv8，可选）</div>
              <label className="form-row dir-row">
                <input
                  type="checkbox"
                  checked={settings.enableDetect !== false}
                  onChange={(e) => setSettings((prev) => ({ ...prev, enableDetect: e.target.checked }))}
                />
                <span>启用检测增强（先用模型框出每张小票再逐张识别，识别率更高；关闭则用整图识别）</span>
              </label>
              <label className="form-row">
                <span>Python 解释器路径</span>
                <input
                  type="text"
                  value={settings.pythonPath || ''}
                  placeholder="留空则使用默认 python（需在 PATH 中）"
                  onChange={(e) => setSettings((prev) => ({ ...prev, pythonPath: e.target.value }))}
                />
              </label>
              <label className="form-row">
                <span>检测模型路径（.pt）</span>
                <input
                  type="text"
                  value={settings.detectModel || ''}
                  placeholder="留空则使用 models/ticket_detect.pt"
                  onChange={(e) => setSettings((prev) => ({ ...prev, detectModel: e.target.value }))}
                />
              </label>
              <div className="form-hint">
                本机需安装 Python 及 <code>ultralytics</code>、<code>torch</code>、<code>pillow</code>。若未安装或模型不可用，识图会自动回退为整图识别，不影响使用。
              </div>
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
