// 小票检测封装：在主进程通过子进程调用 Python（ultralytics）运行的 YOLOv8 模型，
// 检测图中每张小票的矩形边界框，并返回逐张裁剪后的图片（base64）。
//
// 设计要点：
// 1. 输入图片使用 image-service 的 readImageBase64 —— 即前端预览所用的「缩放后 JPEG」，
//    因此检测框坐标天然对齐预览，裁剪图也正好可直接喂给视觉模型做单票识别。
// 2. 通过 stdin 把图片 base64 传给 Python 脚本，脚本把结果 JSON 打到 stdout。
// 3. 任何异常（Python 缺失 / 依赖未装 / 模型不存在 / 推理报错）都会让脚本输出
//    {"ok": false, "message": ...} 并以退出码 0 结束；Node 侧解析后优雅回退到整图识别，
//    绝不让识别功能因缺 Python 环境而崩溃。

import { spawn } from 'child_process'
import { readFileSync, existsSync } from 'fs'
import { join } from 'path'
import { readImageBase64 } from './image-service'

export interface DetectedBox {
  x: number
  y: number
  w: number
  h: number
  conf: number
  cls: number
  label?: string
  /** 该小票裁剪图被自动旋转的角度（度），由检测脚本给出 */
  angle?: number
}

export interface DetectResult {
  ok: boolean
  message?: string
  boxes: DetectedBox[]
  crops: string[] // base64 jpeg，与 boxes 一一对应
  imageWidth: number
  imageHeight: number
  /** 模型/Python 是否可用；false 时调用方应回退到整图识别 */
  modelAvailable: boolean
}

export interface DetectOptions {
  imagePath?: string
  /** 渲染端已按需旋转/缩放后的 base64（优先于 imagePath），用于让检测图与展示图一致 */
  imageBase64?: string
  modelPath?: string
  pythonPath?: string
  conf?: number
  /** 是否同时返回裁剪图（逐张识别需要，仅画框时不需要） */
  crops?: boolean
}

// 默认模型路径：dist-electron/main.js -> 项目根/models/ticket_detect.pt
export function getDefaultModelPath(): string {
  return join(__dirname, '..', 'models', 'ticket_detect.pt')
}

// 检测脚本路径：dist-electron/main.js -> 项目根/scripts/detect_tickets.py
function getScriptPath(): string {
  return join(__dirname, '..', 'scripts', 'detect_tickets.py')
}

// 单张图片检测的最长等待时间（模型冷加载可能较慢，给足 90s）
const DETECT_TIMEOUT_MS = 90_000

function runPython(
  pythonBin: string,
  args: string[],
  stdinText: string,
): Promise<{ stdout: string; stderr: string; code: number | null; error?: string }> {
  return new Promise((resolve) => {
    let proc
    try {
      proc = spawn(pythonBin, args, { windowsHide: true })
    } catch (err) {
      resolve({ stdout: '', stderr: '', code: null, error: err instanceof Error ? err.message : String(err) })
      return
    }

    let stdout = ''
    let stderr = ''
    let timedOut = false
    const timer = setTimeout(() => {
      timedOut = true
      try {
        proc.kill('SIGKILL')
      } catch {
        // ignore
      }
    }, DETECT_TIMEOUT_MS)

    proc.stdout.on('data', (d) => {
      stdout += d.toString()
    })
    proc.stderr.on('data', (d) => {
      stderr += d.toString()
    })
    proc.on('error', (err) => {
      clearTimeout(timer)
      resolve({ stdout, stderr, code: null, error: err.message })
    })
    proc.on('close', (code) => {
      clearTimeout(timer)
      if (timedOut) {
        resolve({ stdout, stderr, code, error: `检测超时（>${DETECT_TIMEOUT_MS / 1000}s），请检查 Python 环境或模型加载是否过慢` })
      } else {
        resolve({ stdout, stderr, code, error: undefined })
      }
    })

    try {
      proc.stdin.write(stdinText)
      proc.stdin.end()
    } catch {
      try {
        proc.kill()
      } catch {
        // ignore
      }
    }
  })
}

export async function detectTickets(opts: DetectOptions): Promise<DetectResult> {
  const empty: DetectResult = {
    ok: false,
    boxes: [],
    crops: [],
    imageWidth: 0,
    imageHeight: 0,
    modelAvailable: false,
  }

  // 1. 读取图片 base64：优先使用前端传来的（已旋转/缩放后的）base64，否则读原图缩放
  const img = opts.imageBase64 && opts.imageBase64.trim()
    ? { error: false, base64: opts.imageBase64.trim(), mime: 'image/jpeg' }
    : readImageBase64(opts.imagePath || '')
  if (img.error || !img.base64) {
    return { ...empty, message: img.message || '读取图片失败' }
  }

  const scriptPath = getScriptPath()
  if (!existsSync(scriptPath)) {
    return { ...empty, message: '检测脚本缺失：' + scriptPath }
  }

  const modelPath = opts.modelPath || getDefaultModelPath()
  const pythonBin = opts.pythonPath && opts.pythonPath.trim() ? opts.pythonPath.trim() : 'python'
  const args = [
    scriptPath,
    modelPath,
    '--conf',
    String(opts.conf ?? 0.25),
  ]
  if (!opts.crops) args.push('--no-crops')

  const run = await runPython(pythonBin, args, img.base64)

  if (run.error) {
    // Python 解释器都起不来（未安装 / 不在 PATH）→ 直接判定模型不可用，回退整图识别
    const msg = /spawn .* ENOENT/.test(run.error) || /ENOENT/.test(run.error)
      ? `未找到 Python 解释器「${pythonBin}」，请在设置中指定正确的 Python 路径，或安装 Python 与 ultralytics。`
      : `检测进程异常：${run.error}`
    return { ...empty, message: msg }
  }

  const out = run.stdout.trim()
  if (!out) {
    // 脚本没输出（例如崩溃打印到 stderr），按模型不可用回退
    return {
      ...empty,
      message: '检测脚本无输出：' + (run.stderr.trim() || `退出码 ${run.code}`),
    }
  }

  let parsed: unknown
  try {
    parsed = JSON.parse(out)
  } catch {
    return {
      ...empty,
      message: '检测脚本返回非 JSON：' + out.slice(0, 200),
    }
  }

  const data = parsed as Record<string, unknown>
  if (data.ok === false) {
    // 脚本主动报告不可用（缺依赖 / 模型不存在等）—— 标记为不可用以便回退
    return {
      ...empty,
      modelAvailable: false,
      message: String(data.message || '检测模型不可用'),
    }
  }

  const boxes = (Array.isArray(data.boxes) ? data.boxes : []) as DetectedBox[]
  const crops = (Array.isArray(data.crops) ? data.crops : []) as string[]
  return {
    ok: true,
    boxes,
    crops,
    imageWidth: Number(data.image_width) || 0,
    imageHeight: Number(data.image_height) || 0,
    modelAvailable: true,
  }
}

// 仅供测试：从文件读取 base64 并直接调用脚本（绕过 readImageBase64），便于离线验证解析逻辑。
export async function detectFromFile(
  imagePath: string,
  opts?: { modelPath?: string; pythonPath?: string; crops?: boolean },
): Promise<DetectResult> {
  const b64 = readFileSync(imagePath).toString('base64')
  const scriptPath = getScriptPath()
  const modelPath = opts?.modelPath || getDefaultModelPath()
  const pythonBin = opts?.pythonPath && opts.pythonPath.trim() ? opts.pythonPath.trim() : 'python'
  const args = [scriptPath, modelPath]
  if (!opts?.crops) args.push('--no-crops')
  const run = await runPython(pythonBin, args, b64)
  if (run.error || !run.stdout.trim()) {
    return { ok: false, boxes: [], crops: [], imageWidth: 0, imageHeight: 0, modelAvailable: false, message: run.error || run.stderr || 'no output' }
  }
  const data = JSON.parse(run.stdout) as Record<string, unknown>
  if (data.ok === false) {
    return { ok: false, boxes: [], crops: [], imageWidth: 0, imageHeight: 0, modelAvailable: false, message: String(data.message) }
  }
  return {
    ok: true,
    boxes: (data.boxes as DetectedBox[]) || [],
    crops: (data.crops as string[]) || [],
    imageWidth: Number(data.image_width) || 0,
    imageHeight: Number(data.image_height) || 0,
    modelAvailable: true,
  }
}
