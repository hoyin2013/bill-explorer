import { useEffect, useRef, useState } from 'react'
import { ElectronAPI, AIRecognizedRow } from '../types'

interface Props {
  api: ElectronAPI
}

// 从一批识别结果里提取去重后的人名
function uniquePersons(rows: AIRecognizedRow[]): string[] {
  const set = new Set<string>()
  for (const r of rows) {
    const p = (r.person || '').trim()
    if (p) set.add(p)
  }
  return Array.from(set)
}

export function ImageWindow({ api }: Props) {
  const [imageDir, setImageDir] = useState('')
  const [images, setImages] = useState<Array<{ name: string; path: string }>>([])
  const [selected, setSelected] = useState('')
  const [preview, setPreview] = useState('')
  const [rotate, setRotate] = useState(0)
  const [loading, setLoading] = useState(false)
  const [aiLoading, setAiLoading] = useState(false)
  const [status, setStatus] = useState('')
  const [rows, setRows] = useState<AIRecognizedRow[]>([])
  const [recognized, setRecognized] = useState<string[]>([])
  const scanningRef = useRef(false)

  // 把当前已识别的人名实时上报给主窗口（主窗口据此把对应 Excel 置顶）
  useEffect(() => {
    api.reportPersons(recognized)
  }, [recognized, api])

  async function loadImages() {
    if (!imageDir) return
    setLoading(true)
    setStatus('')
    try {
      const res = await api.listImages(imageDir)
      if (res.error) {
        setStatus(res.message || '读取图片列表失败')
        setImages([])
      } else {
        const imgs = res.images || []
        setImages(imgs)
        if (imgs.length) {
          setSelected(imgs[0].path)
          await loadPreview(imgs[0].path)
        } else {
          setStatus('该图片目录下没有图片')
        }
      }
    } catch (err) {
      setStatus('读取图片列表失败：' + (err instanceof Error ? err.message : ''))
    } finally {
      setLoading(false)
    }
  }

  async function loadPreview(path: string) {
    setSelected(path)
    setRotate(0)
    setRows([])
    setPreview('')
    const res = await api.readImageBase64(path)
    if (res.error) {
      setStatus(res.message || '读取图片失败')
    } else if (res.base64 && res.mime) {
      setPreview(`data:${res.mime};base64,${res.base64}`)
    }
  }

  // 识别单张图片：只展示结果，由用户点「填入当前录入」回填
  async function doRecognize(path: string) {
    if (!path) {
      setStatus('请先选择一张图片')
      return
    }
    setAiLoading(true)
    setStatus('AI 识别中…')
    setRows([])
    try {
      const res = await api.aiRecognize(path)
      if (res.error) {
        setStatus(res.message || '识别失败')
      } else {
        const rs = res.rows || []
        setRows(rs)
        const ps = uniquePersons(rs)
        if (ps.length) {
          setRecognized((prev) => Array.from(new Set([...prev, ...ps])))
          setStatus(`识别完成，共 ${rs.length} 条记录，提取人名：${ps.join('、')}`)
        } else {
          setStatus(`识别完成，共 ${rs.length} 条记录（未识别到人名）`)
        }
      }
    } catch (err) {
      setStatus('识别失败：' + (err instanceof Error ? err.message : '未知错误'))
    } finally {
      setAiLoading(false)
    }
  }

  // 扫描全部图片：逐张识别，聚合并上报所有人名（用于把左侧列表对应文件置顶）
  async function scanAll() {
    if (scanningRef.current) return
    if (!images.length) {
      setStatus('没有图片可扫描')
      return
    }
    scanningRef.current = true
    const all: AIRecognizedRow[] = []
    try {
      for (let i = 0; i < images.length; i++) {
        const img = images[i]
        setStatus(`正在扫描图片 ${i + 1}/${images.length}：${img.name}`)
        let res: { error?: boolean; message?: string; rows?: AIRecognizedRow[] } | undefined
        try {
          res = await api.aiRecognize(img.path)
        } catch {
          res = undefined
        }
        if (res && res.rows) all.push(...res.rows)
      }
      const ps = uniquePersons(all)
      setRows(all.slice(0, 50)) // 结果表仅展示前 50 条便于查看
      setRecognized((prev) => Array.from(new Set([...prev, ...ps])))
      setStatus(
        ps.length
          ? `扫描完成：${images.length} 张图片、${all.length} 条小票，提取到 ${ps.length} 个人名（${ps.join('、')}）。已把对应 Excel 置顶到左侧列表。`
          : `扫描完成：${images.length} 张图片、${all.length} 条小票，未识别到人名。`,
      )
    } finally {
      scanningRef.current = false
    }
  }

  async function chooseDir() {
    const dir = await api.selectImageDirectory()
    if (dir) {
      setImageDir(dir)
      await loadImages()
    }
  }

  function clearHits() {
    setRecognized([])
    setRows([])
    setStatus('已清除左侧列表的人名置顶')
  }

  function applyToMain() {
    if (!rows.length) {
      setStatus('没有可回填的识别结果')
      return
    }
    api.applyToMain(rows)
    setStatus(`已把 ${rows.length} 条识别结果发送到主窗口，点击主窗口网格即可核对后保存（不影响当前录入，除非你主动编辑）`)
  }

  // 初始化：读取设置中的图片目录
  useEffect(() => {
    api
      .getSettings()
      .then((s) => {
        if (s.imageDir) {
          setImageDir(s.imageDir)
          loadImages()
        } else {
          setStatus('请先选择小票图片目录')
        }
      })
      .catch(() => setStatus('读取设置失败'))
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [])

  function fmt(v: string | number | undefined) {
    return v == null ? '' : String(v)
  }

  return (
    <div className="image-window">
      <div className="image-panel">
        <div className="image-panel-header">
          <span className="image-panel-title">小票识图（独立窗口）</span>
          {imageDir && (
            <span className="image-panel-dir" title={imageDir}>
              {imageDir}
            </span>
          )}
        </div>

        <button className="btn btn-small btn-outline" onClick={chooseDir} style={{ alignSelf: 'flex-start' }}>
          选择图片目录
        </button>

        <div className="image-toolbar">
          <select
            className="image-select"
            value={selected}
            onChange={(e) => loadPreview(e.target.value)}
            disabled={loading || images.length === 0}
          >
            {images.length === 0 && <option value="">{loading ? '读取中…' : '无图片'}</option>}
            {images.map((img) => (
              <option key={img.path} value={img.path}>
                {img.name}
              </option>
            ))}
          </select>
          <button
            className="btn btn-small btn-outline"
            onClick={() => setRotate((r) => (r + 90) % 360)}
            disabled={!preview}
            title="旋转 90°"
          >
            旋转
          </button>
        </div>

        <div className="image-preview-wrap">
          {preview ? (
            <img
              src={preview}
              alt="小票预览"
              className="image-preview"
              style={{ transform: `rotate(${rotate}deg)` }}
            />
          ) : (
            <div className="image-empty">选择图片后在此预览，可对照录入</div>
          )}
        </div>

        <div className="ai-actions">
          <button
            className="btn btn-primary"
            onClick={() => doRecognize(selected)}
            disabled={aiLoading || !selected}
          >
            {aiLoading ? '识别中…' : 'AI 识别当前图片'}
          </button>
          <button
            className="btn btn-outline"
            onClick={scanAll}
            disabled={aiLoading || loading || images.length === 0}
            title="逐张识别全部图片，提取人名并置顶左侧对应 Excel"
          >
            扫描全部图片（提取人名）
          </button>
        </div>

        {recognized.length > 0 && (
          <div className="person-chips">
            <span className="person-chips-label">已识别人名：</span>
            {recognized.map((p) => (
              <span className="person-chip" key={p}>
                {p}
              </span>
            ))}
            <button className="btn btn-small btn-link" onClick={clearHits}>
              清除置顶
            </button>
          </div>
        )}

        {rows.length > 0 && (
          <div className="ai-results">
            <div className="ai-results-header">
              <span>识别结果（{rows.length} 条）</span>
              <div className="ai-result-actions">
                <button className="btn btn-small btn-primary" onClick={applyToMain}>
                  填入当前录入
                </button>
              </div>
            </div>
            <div className="ai-results-table-wrap">
              <table className="ai-results-table">
                <thead>
                  <tr>
                    <th>序号</th>
                    <th>日期</th>
                    <th>货品</th>
                    <th>单位</th>
                    <th>数量</th>
                    <th>单价</th>
                    <th>金额</th>
                    <th>调货人</th>
                    <th>备注</th>
                  </tr>
                </thead>
                <tbody>
                  {rows.map((r, i) => (
                    <tr key={i}>
                      <td>{fmt(r.no)}</td>
                      <td>{fmt(r.date)}</td>
                      <td>{fmt(r.name)}</td>
                      <td>{fmt(r.unit)}</td>
                      <td>{fmt(r.qty)}</td>
                      <td>{fmt(r.price)}</td>
                      <td>{fmt(r.amount)}</td>
                      <td>{fmt(r.person)}</td>
                      <td>{fmt(r.remark)}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>
        )}

        {status && (
          <div className={status.includes('失败') ? 'image-status error' : 'image-status'}>{status}</div>
        )}
      </div>
    </div>
  )
}
