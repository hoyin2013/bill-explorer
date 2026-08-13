#!/usr/bin/env node
// -*- coding: utf-8 -*-
/**
 * 小票检测器（YOLOv8 · ONNX 运行时，纯 Node 实现，无需 Python / ultralytics / torch）
 * ==========================================================================
 * 设计要点：
 * 1. 读取 stdin 的 base64 图片（JPEG/PNG），用 jpeg-js 解码为 RGB。
 * 2. 与 ultralytics 一致的 letterbox 预处理（等比缩放到 640，灰边补满），
 *    保证检测框坐标可精确映射回原图。
 * 3. 用 onnxruntime-node 跑 YOLOv8 的 ONNX 模型（ticket_detect.onnx），
 *    后处理：解码框 + 类分数(sigmoid) + 按类 NMS，并把框从 640 空间逆 letterbox 回原图坐标。
 * 4. 每个框裁剪（四周外扩 6px），用「逐行墨量方差」在 0/90/180/270 中选最佳旋转角使文字正向，
 *    再编码为 JPEG base64，与 detect_tickets.py 的输出结构完全一致。
 *
 * 输出（stdout，单行 JSON）：
 *   { ok:true, image_width:W, image_height:H, boxes:[{x,y,w,h,conf,cls,label,angle}], crops:[base64...] }
 * 失败则输出 {"ok":false,"message":...} 并以退出码 0 结束，便于 Node 侧优雅回退。
 *
 * 用法（由 Electron 主进程调用）：
 *   node detect_onnx.mjs <model_path> [--conf 0.25] [--iou 0.45] [--imgsz 640] [--no-crops] [--no-rotate]
 * 图片经 stdin 传入；坐标与裁剪均基于传入的图片（即预览缩放后的图）。
 */

import ort from 'onnxruntime-node'
import jpeg from 'jpeg-js'

const IMGSZ = 640
const PAD = 6

function fail(message) {
  process.stdout.write(JSON.stringify({ ok: false, message: String(message) }))
  process.stdout.write('\n')
  process.exit(0)
}

function sigmoid(x) {
  return 1 / (1 + Math.exp(-x))
}

function clamp(v, lo, hi) {
  return v < lo ? lo : v > hi ? hi : v
}

// ---------- 图像基础操作（纯 RGB Uint8Array） ----------

// 双线性缩放 RGB
function bilinearResize(src, sw, sh, dw, dh) {
  const out = new Uint8Array(dw * dh * 3)
  if (dw === 0 || dh === 0) return out
  for (let y = 0; y < dh; y++) {
    const sy = (y + 0.5) * sh / dh - 0.5
    const y0 = Math.max(0, Math.min(sh - 1, Math.floor(sy)))
    const y1 = Math.max(0, Math.min(sh - 1, y0 + 1))
    const fy = sy - y0
    for (let x = 0; x < dw; x++) {
      const sx = (x + 0.5) * sw / dw - 0.5
      const x0 = Math.max(0, Math.min(sw - 1, Math.floor(sx)))
      const x1 = Math.max(0, Math.min(sw - 1, x0 + 1))
      const fx = sx - x0
      for (let c = 0; c < 3; c++) {
        const p00 = src[(y0 * sw + x0) * 3 + c]
        const p10 = src[(y0 * sw + x1) * 3 + c]
        const p01 = src[(y1 * sw + x0) * 3 + c]
        const p11 = src[(y1 * sw + x1) * 3 + c]
        const top = p00 + (p10 - p00) * fx
        const bot = p01 + (p11 - p01) * fx
        out[(y * dw + x) * 3 + c] = Math.round(top + (bot - top) * fy)
      }
    }
  }
  return out
}

// 逆时针旋转 RGB k*90 度（k=0/1/2/3）。与 PIL rotate(expand=True) 方向一致，
// 用于「计算方差」和「实际旋转」同一约定，保证选出使文字水平的角。
function rotateCCW(src, w, h, k) {
  if (k === 0) return src
  if (k === 1) {
    const nw = h
    const nh = w
    const out = new Uint8Array(nw * nh * 3)
    for (let y = 0; y < h; y++) {
      for (let x = 0; x < w; x++) {
        const s = (y * w + x) * 3
        const d = ((w - 1 - x) * nw + y) * 3
        out[d] = src[s]
        out[d + 1] = src[s + 1]
        out[d + 2] = src[s + 2]
      }
    }
    return out
  }
  if (k === 2) {
    const out = new Uint8Array(w * h * 3)
    for (let y = 0; y < h; y++) {
      for (let x = 0; x < w; x++) {
        const s = (y * w + x) * 3
        const d = ((h - 1 - y) * w + (w - 1 - x)) * 3
        out[d] = src[s]
        out[d + 1] = src[s + 1]
        out[d + 2] = src[s + 2]
      }
    }
    return out
  }
  // k === 3 : 逆时针 270 = 顺时针 90
  const nw = h
  const nh = w
  const out = new Uint8Array(nw * nh * 3)
  for (let y = 0; y < h; y++) {
    for (let x = 0; x < w; x++) {
      const s = (y * w + x) * 3
      const d = (x * nw + (h - 1 - y)) * 3
      out[d] = src[s]
      out[d + 1] = src[s + 1]
      out[d + 2] = src[s + 2]
    }
  }
  return out
}

// 逐行墨量方差：文字水平时每行墨量在文本行处呈明显峰谷，方差最大
function rowProjectionVariance(rgb, w, h) {
  if (w === 0 || h === 0) return 0
  const rowSums = new Float64Array(h)
  let sum = 0
  for (let y = 0; y < h; y++) {
    let s = 0
    for (let x = 0; x < w; x++) {
      const i = (y * w + x) * 3
      const gray = 0.299 * rgb[i] + 0.587 * rgb[i + 1] + 0.114 * rgb[i + 2]
      s += 255 - gray
    }
    rowSums[y] = s
    sum += s
  }
  const mean = sum / h
  let v = 0
  for (let y = 0; y < h; y++) {
    const d = rowSums[y] - mean
    v += d * d
  }
  return v / h
}

// 在 0/90/180/270 中选使行投影方差最大的角度（文字行最水平）
function bestRotation(rgb, w, h) {
  const maxside = Math.max(w, h)
  let rw = w
  let rh = h
  let r = rgb
  if (maxside > 400) {
    const ratio = 400 / maxside
    rw = Math.max(1, Math.round(w * ratio))
    rh = Math.max(1, Math.round(h * ratio))
    r = bilinearResize(rgb, w, h, rw, rh)
  }
  let best = -1
  let bestAngle = 0
  for (const angle of [0, 90, 180, 270]) {
    const k = angle / 90
    const rot = k === 0 ? r : rotateCCW(r, rw, rh, k)
    let rw2 = rw
    let rh2 = rh
    if (angle === 90 || angle === 270) {
      rw2 = rh
      rh2 = rw
    }
    const score = rowProjectionVariance(rot, rw2, rh2)
    if (score > best) {
      best = score
      bestAngle = angle
    }
  }
  return bestAngle
}

function readStdin() {
  return new Promise((resolve, reject) => {
    let buf = ''
    process.stdin.setEncoding('utf8')
    process.stdin.on('data', (d) => (buf += d))
    process.stdin.on('end', () => resolve(buf))
    process.stdin.on('error', reject)
  })
}

async function main() {
  const argv = process.argv.slice(2)
  // 正确的参数解析：遇到 --flag value 时把 value 一并消费，避免像 0.25 这类值被误当作位置参数（模型路径）。
  const positional = []
  const flags = {}
  for (let i = 0; i < argv.length; i++) {
    const a = argv[i]
    if (a.startsWith('--')) {
      if (i + 1 < argv.length && !argv[i + 1].startsWith('--')) {
        flags[a] = argv[i + 1]
        i++ // 消费掉它的值
      } else {
        flags[a] = true
      }
    } else {
      positional.push(a)
    }
  }
  const modelPath = positional[0] || 'models/ticket_detect.onnx'
  const hasFlag = (f) => !!flags[f]
  const getNum = (flag, def) =>
    flags[flag] !== undefined ? parseFloat(flags[flag]) : def
  const conf = getNum('--conf', 0.25)
  const iou = getNum('--iou', 0.45)
  const imgsz = Math.round(getNum('--imgsz', IMGSZ))
  const noCrops = hasFlag('--no-crops')
  const noRotate = hasFlag('--no-rotate')

  // ---- 1. 读取 stdin base64 ----
  let raw
  try {
    raw = (await readStdin()).trim()
  } catch (e) {
    fail('读取 stdin 失败：' + e.message)
    return
  }
  if (!raw) {
    fail('未接收到图片数据（stdin 为空）')
    return
  }
  let imgBytes
  try {
    imgBytes = Buffer.from(raw, 'base64')
  } catch (e) {
    fail('图片 base64 解码失败：' + e.message)
    return
  }

  // ---- 2. 解码为 RGB（原图坐标） ----
  let decoded
  try {
    decoded = jpeg.decode(imgBytes, { useTightestPalette: false, maxMemoryUsageInMB: 1024 })
  } catch (e) {
    fail('图片解码失败（仅支持 JPEG）：' + e.message)
    return
  }
  const W = decoded.width
  const H = decoded.height
  const rgba = decoded.data // Uint8Array, RGBA
  const rgb = new Uint8Array(W * H * 3)
  for (let i = 0; i < W * H; i++) {
    rgb[i * 3] = rgba[i * 4]
    rgb[i * 3 + 1] = rgba[i * 4 + 1]
    rgb[i * 3 + 2] = rgba[i * 4 + 2]
  }

  // ---- 3. 加载 ONNX 模型 ----
  if (!modelPath || modelPath.endsWith('.pt')) {
    fail('ONNX 模型未找到，请先用 ultralytics 把 .pt 导出为 .onnx（导出步骤才需要一次 Python）。')
    return
  }
  let session
  try {
    session = await ort.InferenceSession.create(modelPath, {
      executionProviders: ['cpu'],
      graphOptimizationLevel: 'all',
    })
  } catch (e) {
    fail('加载 ONNX 模型失败：' + e.message)
    return
  }

  // ---- 4. letterbox 预处理到 imgsz ----
  const r = Math.min(imgsz / W, imgsz / H)
  const nw = Math.max(1, Math.round(W * r))
  const nh = Math.max(1, Math.round(H * r))
  const padXf = (imgsz - nw) / 2
  const padYf = (imgsz - nh) / 2
  const resized = bilinearResize(rgb, W, H, nw, nh)
  const input = new Float32Array(3 * imgsz * imgsz)
  input.fill(114 / 255) // 灰边
  for (let y = 0; y < nh; y++) {
    for (let x = 0; x < nw; x++) {
      const dx = Math.floor(padXf) + x
      const dy = Math.floor(padYf) + y
      if (dx < 0 || dy < 0 || dx >= imgsz || dy >= imgsz) continue
      const si = (y * nw + x) * 3
      const di = dy * imgsz + dx
      input[di] = resized[si] / 255
      input[imgsz * imgsz + di] = resized[si + 1] / 255
      input[2 * imgsz * imgsz + di] = resized[si + 2] / 255
    }
  }

  // ---- 5. 推理 ----
  let output
  try {
    const tensor = new ort.Tensor('float32', input, [1, 3, imgsz, imgsz])
    const feeds = {}
    feeds[session.inputNames[0]] = tensor
    const res = await session.run(feeds)
    output = res[session.outputNames[0]]
  } catch (e) {
    fail('ONNX 推理失败：' + e.message)
    return
  }

  // ---- 6. 解析输出 [1, C, A]（YOLOv8 ONNX：C=4+类别，A=anchors 为大维度） ----
  const dims = output.dims
  let C // 4 + nc
  let A // anchors（大维度）
  // anchors 数（A）远大于通道数（C），以维度大小判断布局，避免把大维度误当成通道。
  A = Math.max(dims[1], dims[2])
  C = Math.min(dims[1], dims[2])
  // 通道维度是 size=C 的那个：在中维(dim1) 或 末维(dim2)。
  // 行优先展平规则：通道在中维 -> offset = c*A + a；通道在末维 -> offset = a*C + c。
  // 之前把两个分支写反，导致「类分数」读到框坐标，sigmoid 后恒为 1.0，阈值失效、碎片框全漏进。
  const chanDimIs1 = dims[1] === C
  const data = output.data
  const valAt = (a, c) => (chanDimIs1 ? data[c * A + a] : data[a * C + c])
  const nc = C - 4
  if (nc < 1) {
    fail('模型输出通道数异常（C=' + C + '），无法解析类别。')
    return
  }

  // 类分数是否需 sigmoid：只检查「类」通道自身的值域，不要被框坐标（0~640）带偏。
  // 本模型导出的 ONNX 已把 sigmoid 烘焙进图里，类分数直接落在 [0,1]（已是概率），
  // 此时不应再 sigmoid（否则 sigmoid(0.97)=0.725，等于把阈值废掉、所有锚点全过）。
  // 若类分数超出 [0,1]（原始 logit），才需要 sigmoid。
  let needSigmoid = false
  for (let a = 0; a < A && !needSigmoid; a++) {
    for (let c = 4; c < C; c++) {
      const v = valAt(a, c)
      if (v > 1.5 || v < -0.5) {
        needSigmoid = true
        break
      }
    }
  }

  const rawBoxes = []
  for (let a = 0; a < A; a++) {
    const cx = valAt(a, 0)
    const cy = valAt(a, 1)
    const bw = valAt(a, 2)
    const bh = valAt(a, 3)
    let best = 0
    let bestCls = 0
    for (let c = 0; c < nc; c++) {
      let s = valAt(a, 4 + c)
      if (needSigmoid) s = sigmoid(s)
      if (s > best) {
        best = s
        bestCls = c
      }
    }
    if (best < conf) continue
    // 640 空间 → 原图坐标（逆 letterbox）
    let x1 = (cx - bw / 2 - padXf) / r
    let y1 = (cy - bh / 2 - padYf) / r
    let x2 = (cx + bw / 2 - padXf) / r
    let y2 = (cy + bh / 2 - padYf) / r
    x1 = clamp(x1, 0, W)
    y1 = clamp(y1, 0, H)
    x2 = clamp(x2, 0, W)
    y2 = clamp(y2, 0, H)
    const w = x2 - x1
    const h = y2 - y1
    if (w < 4 || h < 4) continue
    rawBoxes.push({ x: x1, y: y1, w, h, conf: best, cls: bestCls })
  }

  // ---- 7. 按类 NMS ----
  function calcIou(b1, b2) {
    const ix1 = Math.max(b1.x, b2.x)
    const iy1 = Math.max(b1.y, b2.y)
    const ix2 = Math.min(b1.x + b1.w, b2.x + b2.w)
    const iy2 = Math.min(b1.y + b1.h, b2.y + b2.h)
    const iw = Math.max(0, ix2 - ix1)
    const ih = Math.max(0, iy2 - iy1)
    const inter = iw * ih
    const area1 = b1.w * b1.h
    const area2 = b2.w * b2.h
    const u = area1 + area2 - inter
    return u <= 0 ? 0 : inter / u
  }
  const sortedBoxes = rawBoxes.slice().sort((a, b) => b.conf - a.conf)
  const used = new Array(sortedBoxes.length).fill(false)
  const kept = []
  for (let i = 0; i < sortedBoxes.length; i++) {
    if (used[i]) continue
    kept.push(sortedBoxes[i])
    for (let j = i + 1; j < sortedBoxes.length; j++) {
      if (used[j] || sortedBoxes[j].cls !== sortedBoxes[i].cls) continue
      if (calcIou(sortedBoxes[i], sortedBoxes[j]) > iou) used[j] = true
    }
  }

  // ---- 8. 裁剪 + 自动旋转（与 detect_tickets.py 一致） ----
  // 收集 {box, crop} 成对结构，最后统一按 (y,x) 排序，保证 boxes 与 crops 顺序一致。
  const items = []
  for (const b of kept) {
    const x1 = b.x
    const y1 = b.y
    const x2 = b.x + b.w
    const y2 = b.y + b.h
    const cx1 = Math.max(0, Math.floor(x1) - PAD)
    const cy1 = Math.max(0, Math.floor(y1) - PAD)
    const cx2 = Math.min(W, Math.ceil(x2) + PAD)
    const cy2 = Math.min(H, Math.ceil(y2) + PAD)
    const cw = cx2 - cx1
    const ch = cy2 - cy1
    // 提取区域 RGB
    const cropRGB = new Uint8Array(cw * ch * 3)
    for (let y = 0; y < ch; y++) {
      const sRow = (cy1 + y) * W + cx1
      const dRow = y * cw
      for (let x = 0; x < cw; x++) {
        const si = (sRow + x) * 3
        const di = (dRow + x) * 3
        cropRGB[di] = rgb[si]
        cropRGB[di + 1] = rgb[si + 1]
        cropRGB[di + 2] = rgb[si + 2]
      }
    }
    let angle = 0
    let outRGB = cropRGB
    let ow2 = cw
    let oh2 = ch
    if (!noRotate) {
      angle = bestRotation(cropRGB, cw, ch)
      if (angle !== 0) {
        outRGB = rotateCCW(cropRGB, cw, ch, angle / 90)
        if (angle === 90 || angle === 270) {
          ow2 = ch
          oh2 = cw
        }
      }
    }
    let crop = ''
    if (!noCrops) {
      const rgba2 = new Uint8Array(ow2 * oh2 * 4)
      for (let p = 0; p < ow2 * oh2; p++) {
        rgba2[p * 4] = outRGB[p * 3]
        rgba2[p * 4 + 1] = outRGB[p * 3 + 1]
        rgba2[p * 4 + 2] = outRGB[p * 3 + 2]
        rgba2[p * 4 + 3] = 255
      }
      let jpegData
      try {
        jpegData = jpeg.encode({ data: rgba2, width: ow2, height: oh2 }, 90).data
      } catch (e) {
        fail('裁剪图编码失败：' + e.message)
        return
      }
      crop = Buffer.from(jpegData).toString('base64')
    }
    items.push({
      box: {
        x: Math.round(x1),
        y: Math.round(y1),
        w: Math.round(b.w),
        h: Math.round(b.h),
        conf: Math.round(b.conf * 10000) / 10000,
        cls: b.cls,
        label: String(b.cls),
        angle,
      },
      crop,
    })
  }

  // 按从上到下、从左到右排序，保证逐张识别顺序稳定（与 Python 脚本一致）
  items.sort((a, b) => a.box.y - b.box.y || a.box.x - b.box.x)
  const boxes = items.map((it) => it.box)
  const crops = noCrops ? [] : items.map((it) => it.crop)

  process.stdout.write(
    JSON.stringify(
      {
        ok: true,
        image_width: W,
        image_height: H,
        boxes,
        crops,
      },
    ),
  )
  process.stdout.write('\n')
}

main().catch((e) => {
  try {
    process.stdout.write(JSON.stringify({ ok: false, message: '运行异常：' + (e && e.message ? e.message : String(e)) }))
    process.stdout.write('\n')
  } catch {
    // ignore
  }
  process.exit(0)
})
