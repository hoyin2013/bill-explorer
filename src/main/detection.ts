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

import { spawn, spawnSync } from 'child_process'
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

// 默认模型路径：dist-electron/main.js -> 项目根/models/ticket_detect.onnx
// （ONNX 为导出格式，运行时用 onnxruntime-node 推理，无需 Python / ultralytics / torch）
export function getDefaultModelPath(): string {
  return join(__dirname, '..', 'models', 'ticket_detect.onnx')
}

// 检测脚本路径：dist-electron/main.js -> 项目根/scripts/detect_tickets.py（仅旧版 Python 回退用）
function getScriptPath(): string {
  return join(__dirname, '..', 'scripts', 'detect_tickets.py')
}

// ONNX 推理脚本路径：dist-electron/main.js -> 项目根/scripts/detect_onnx.cjs
function getOnnxScriptPath(): string {
  return join(__dirname, '..', 'scripts', 'detect_onnx.cjs')
}

// 候选 Python 解释器（按优先级）。macOS / 多数 Linux 上命令是 python3，
// Windows 上是 python。这里不再写死「python」，避免 macOS 上报
// 「未找到 Python 解释器「python」」这样的吓人错误。
const PYTHON_CANDIDATES = [
  'python3',
  'python',
  'python3.13',
  'python3.12',
  'python3.11',
  'python3.10',
]

// 廉价检查：某解释器是否存在（--version 退出码 0 即认为可用）。
function pythonBinaryExists(bin: string): boolean {
  try {
    const r = spawnSync(bin, ['--version'], { timeout: 8000, windowsHide: true })
    return r.status === 0 && !r.error
  } catch {
    return false
  }
}

// 试探：某解释器能否 `import` 指定模块（用于挑选「已装 ultralytics」的最佳解释器）。
function pythonCanImport(bin: string, moduleName: string): boolean {
  try {
    const r = spawnSync(bin, ['-c', `import ${moduleName}`], { timeout: 15000, windowsHide: true })
    return r.status === 0 && !r.error
  } catch {
    return false
  }
}

// 探测结果缓存：undefined=尚未探测；null=探测过但不可用；string=可用的解释器路径。
let resolvedPython: string | null | undefined = undefined

/** 重置解释器探测缓存（设置变更后可调用以重新探测）。 */
export function resetPythonResolution(): void {
  resolvedPython = undefined
}

/**
 * 解析用于运行检测脚本的 Python 解释器：
 * 1. 若显式指定了 pythonPath，则只校验它是否存在（不存在返回 null）。
 * 2. 否则按优先级探测候选解释器——优先选「已安装 ultralytics」的，
 *    退而求其次选任意可用解释器（ultralytics 缺失时脚本会优雅提示安装）。
 * 3. 全部不可用则返回 null（调用方据此优雅回退，而不是抛出吓人的报错）。
 */
export async function resolvePython(forced?: string): Promise<string | null> {
  if (forced && forced.trim()) {
    const bin = forced.trim()
    return pythonBinaryExists(bin) ? bin : null
  }
  if (resolvedPython !== undefined) return resolvedPython
  let fallback: string | null = null
  for (const candidate of PYTHON_CANDIDATES) {
    if (pythonBinaryExists(candidate)) {
      fallback = candidate
      if (pythonCanImport(candidate, 'ultralytics')) {
        resolvedPython = candidate
        return candidate
      }
    }
  }
  resolvedPython = fallback
  return fallback
}

/** 探测检测增强的可用性，供 UI 提前给出提示并决定是否禁用按钮。
 *  现在基于 ONNX 运行时（纯 Node + onnxruntime-node），不再依赖 Python / ultralytics。 */
export async function detectEnvironment(): Promise<{
  modelExists: boolean
  runtimeReady: boolean
  detail: string
}> {
  const modelPath = getDefaultModelPath()
  const onnxModel = modelPath.toLowerCase().endsWith('.onnx')
    ? modelPath
    : modelPath.replace(/\.pt$/i, '.onnx')
  const modelExists = existsSync(onnxModel) || existsSync(modelPath)

  // 运行时就绪：当前进程即 Electron 自带的 Node，无需系统安装 node。
  // 直接在进程内解析 onnxruntime-node / jpeg-js（打包后它们位于 resources/app/node_modules）。
  // 这样在没有安装 Node 的目标电脑上也能正确判定，不会误报「检测运行时缺失」。
  let runtimeReady = false
  try {
    require.resolve('onnxruntime-node')
    require.resolve('jpeg-js')
    runtimeReady = true
  } catch {
    runtimeReady = false
  }

  let detail = ''
  if (!modelExists) {
    detail = '未找到检测模型（models/ticket_detect.onnx），请用 ultralytics 一次性导出 ONNX。'
  } else if (!runtimeReady) {
    detail = '检测运行时缺失：请运行 npm install onnxruntime-node jpeg-js（无需 Python）。'
  }
  return { modelExists, runtimeReady, detail }
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

// 运行 ONNX 推理：直接在 Electron 主进程（自带 Node）内调用 detect_onnx.cjs 的 runDetect，
// 不再 spawn 外部 `node` 进程 —— 这样在没有安装 Node 的目标电脑上也能正常推理。
function runOnnxInProcess(
  scriptPath: string,
  modelPath: string,
  opts: { conf?: number; iou?: number; imgsz?: number; noCrops?: boolean; noRotate?: boolean },
  stdinText: string,
): Promise<{ stdout: string; stderr: string; code: number | null; error?: string }> {
  return new Promise((resolve) => {
    const run = async () => {
      // 进程内 require 脚本（Electron 自带 Node 可直接加载 onnxruntime-node 原生绑定，N-API 跨版本兼容）
      const mod = require(scriptPath)
      const result = await mod.runDetect(stdinText, modelPath, opts)
      return { stdout: JSON.stringify(result), stderr: '', code: 0 }
    }
    run().then(
      (r) => resolve(r),
      (err) => resolve({ stdout: '', stderr: '', code: null, error: err instanceof Error ? err.message : String(err) }),
    )
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

  const modelPath = opts.modelPath || getDefaultModelPath()
  // 运行时优先用 ONNX（纯 Node + onnxruntime-node），无需 Python / ultralytics / torch。
  // 用户训练好的 ticket_detect.pt 需一次性导出为 ticket_detect.onnx（导出那步才需要 Python）。
  const onnxModel = modelPath.toLowerCase().endsWith('.onnx')
    ? modelPath
    : modelPath.replace(/\.pt$/i, '.onnx')

  // 统一解析子进程输出的 JSON 为 DetectResult（两个分支共用）
  const finish = (
    run: { stdout: string; stderr: string; code: number | null; error?: string },
    runnerLabel: string,
  ): DetectResult => {
    if (run.error) {
      return { ...empty, message: `${runnerLabel} 推理进程无法启动：${run.error}。` }
    }
    const out = run.stdout.trim()
    if (!out) {
      return { ...empty, message: `${runnerLabel} 推理无输出：${run.stderr.trim() || '退出码 ' + run.code}` }
    }
    let parsed: unknown
    try {
      parsed = JSON.parse(out)
    } catch {
      return { ...empty, message: `${runnerLabel} 推理返回非 JSON：` + out.slice(0, 200) }
    }
    const data = parsed as Record<string, unknown>
    if (data.ok === false) {
      return { ...empty, modelAvailable: false, message: String(data.message || '检测模型不可用') }
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

  // ---- 主路径：ONNX 运行时 ----
  if (existsSync(onnxModel)) {
    const scriptPath = getOnnxScriptPath()
    if (!existsSync(scriptPath)) {
      return { ...empty, message: 'ONNX 推理脚本缺失：' + scriptPath }
    }
    const run = await runOnnxInProcess(
      scriptPath,
      onnxModel,
      { conf: opts.conf ?? 0.25, iou: 0.45, imgsz: 640, noCrops: !opts.crops, noRotate: false },
      img.base64,
    )
    return finish(run, 'ONNX')
  }

  // ---- 回退：旧版 Python + ultralytics（仅当 .pt 在且装有 ultralytics 时可用） ----
  if (existsSync(modelPath) && modelPath.toLowerCase().endsWith('.pt')) {
    const scriptPath = getScriptPath()
    if (!existsSync(scriptPath)) {
      return { ...empty, message: '检测脚本缺失：' + scriptPath }
    }
    const pythonBin = await resolvePython(opts.pythonPath)
    if (!pythonBin) {
      return {
        ...empty,
        message:
          '检测模型需要 ONNX 运行时（models/ticket_detect.onnx）。请先导出 ONNX，或在设置中指定已安装 ultralytics 的 Python 路径。',
      }
    }
    const args = [scriptPath, modelPath, '--conf', String(opts.conf ?? 0.25)]
    if (!opts.crops) args.push('--no-crops')
    const run = await runPython(pythonBin, args, img.base64)
    return finish(run, 'Python')
  }

  // ---- 都没有 ----
  return {
    ...empty,
    message:
      '未找到检测模型（models/ticket_detect.onnx）。请用 ultralytics 一次性导出：YOLO("' +
      modelPath +
      '").export(format="onnx", imgsz=640)。',
  }
}

// 仅供测试：从文件读取 base64 并直接调用脚本（绕过 readImageBase64），便于离线验证解析逻辑。
export async function detectFromFile(
  imagePath: string,
  opts?: { modelPath?: string; pythonPath?: string; crops?: boolean },
): Promise<DetectResult> {
  const b64 = readFileSync(imagePath).toString('base64')
  const modelPath = opts?.modelPath || getDefaultModelPath()
  const onnxModel = modelPath.toLowerCase().endsWith('.onnx')
    ? modelPath
    : modelPath.replace(/\.pt$/i, '.onnx')

  if (existsSync(onnxModel)) {
    const scriptPath = getOnnxScriptPath()
    const run = await runOnnxInProcess(
      scriptPath,
      onnxModel,
      { conf: 0.25, iou: 0.45, imgsz: 640, noCrops: !opts?.crops, noRotate: false },
      b64,
    )
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

  // 回退：Python + ultralytics
  const scriptPath = getScriptPath()
  const pythonBin = await resolvePython(opts?.pythonPath)
  if (!pythonBin) {
    return { ok: false, boxes: [], crops: [], imageWidth: 0, imageHeight: 0, modelAvailable: false, message: '未找到可用的 Python 解释器' }
  }
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
