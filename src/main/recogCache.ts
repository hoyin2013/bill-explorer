import * as crypto from 'crypto'
import * as fs from 'fs'
import * as path from 'path'
import { app } from 'electron'

// 识别结果缓存：以「图片内容 + 模型/提示词指纹」为键，命中后直接复用，避免重复请求 AI。
// 缓存同时落在内存（进程内，零 IO）与磁盘（userData/recog-cache，跨启动持久），保证「识别一次后，
// 只有新图片（内容或模型/提示词变化）才会重新请求」。注意：key 不含 apiKey，避免把密钥写进缓存文件。

interface ConfigFingerprint {
  baseURL?: string
  model?: string
  temperature?: number
}

const mem = new Map<string, unknown>()
let cacheDir: string | null = null

function dir(): string {
  if (!cacheDir) {
    cacheDir = path.join(app.getPath('userData'), 'recog-cache')
    fs.mkdirSync(cacheDir, { recursive: true })
  }
  return cacheDir
}

export function recogCacheKey(imageBase64: string, cfg: ConfigFingerprint, prompt: string): string {
  const img = crypto.createHash('sha256').update(imageBase64).digest('hex')
  const promptHash = crypto.createHash('sha256').update(prompt || '').digest('hex')
  const conf = `${cfg.baseURL || ''}|${cfg.model || ''}|${cfg.temperature ?? 0.2}|${promptHash}`
  return crypto.createHash('sha256').update(img + '|' + conf).digest('hex')
}

export function getRecogCached(key: string): unknown | null {
  if (mem.has(key)) return mem.get(key) ?? null
  try {
    const p = path.join(dir(), key + '.json')
    if (fs.existsSync(p)) {
      const raw = JSON.parse(fs.readFileSync(p, 'utf8'))
      mem.set(key, raw)
      return raw
    }
  } catch {
    /* 缓存损坏则忽略，当作未命中 */
  }
  return null
}

export function setRecogCached(key: string, value: unknown): void {
  mem.set(key, value)
  try {
    fs.writeFileSync(path.join(dir(), key + '.json'), JSON.stringify(value))
  } catch {
    /* 磁盘写入失败不影响主流程（内存仍有） */
  }
}

export function clearRecogCache(): number {
  const inMem = mem.size
  mem.clear()
  let files = 0
  try {
    const d = dir()
    for (const f of fs.readdirSync(d)) {
      if (f.endsWith('.json')) {
        try {
          fs.unlinkSync(path.join(d, f))
          files++
        } catch {
          /* 忽略单个删除失败 */
        }
      }
    }
  } catch {
    /* 目录读不到也无妨 */
  }
  return inMem + files
}
