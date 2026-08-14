#!/usr/bin/env node
// -*- coding: utf-8 -*-
/**
 * 小票检测器（YOLOv8 · ONNX 运行时，纯 Node 实现，无需 Python / ultralytics / torch）
 * ==========================================================================
 * 本文件为 CommonJS（.cjs），既可由 Electron 主进程「进程内 require」调用 runDetect，
 * 也可作为 CLI 单独运行（node detect_onnx.cjs <model> ... 图片经 stdin 传入）。
 *
 * runDetect(base64, modelPath, opts) -> Promise<object>
 *   成功：{ ok:true, image_width, image_height, boxes:[{x,y,w,h,conf,cls,label,angle}], crops:[base64...] }
 *   失败：{ ok:false, message }
 * 注意：本模块绝不在进程内调用 process.exit（那是 CLI 单跑时才需要），
 * 否则会被 Electron 主进程 require 时把整个 app 干掉。
 */
'use strict'

const ortRaw = require('onnxruntime-node')
const ort = ortRaw.default || ortRaw
const jpeg = require('jpeg-js')

const IMGSZ = 640
const PAD = 6

function errObj(message) {
  return { ok: false, message: String(message) }
}

function sigmoid(x) {
  return 1 / (1 + Math.exp(-x))
}

function clamp(v, lo, hi) {
  return v < lo ? lo : v > hi ? hi : v
}

// ---------- 图像基础操作（纯 RGB Uint8Array） ----------

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

// 逆时针旋转 RGB k*90 度（k=0/1/2/3）
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

// 核心：检测一张 base64 图片，返回结果对象（不写 stdout、不 exit）。
async function runDetect(base64, modelPath, opts) {
  opts = opts || {}
  const conf = typeof opts.conf === 'number' ? opts.conf : 0.25
  const iou = typeof opts.iou === 'number' ? opts.iou : 0.45
  const imgsz = Math.round(typeof opts.imgsz === 'number' ? opts.imgsz : IMGSZ)
  const noCrops = !!opts.noCrops
  const noRotate = !!opts.noRotate

  if (!base64 || !base64.trim()) {
    return errObj('未接收到图片数据（base64 为空）')
  }
  let imgBytes
  try {
    imgBytes = Buffer.from(base64.trim(), 'base64')
  } catch (e) {
    return errObj('图片 base64 解码失败：' + e.message)
  }

  let decoded
  try {
    decoded = jpeg.decode(imgBytes, { useTightestPalette: false, maxMemoryUsageInMB: 1024 })
  } catch (e) {
    return errObj('图片解码失败（仅支持 JPEG）：' + e.message)
  }
  const W = decoded.width
  const H = decoded.height
  const rgba = decoded.data
  const rgb = new Uint8Array(W * H * 3)
  for (let i = 0; i < W * H; i++) {
    rgb[i * 3] = rgba[i * 4]
    rgb[i * 3 + 1] = rgba[i * 4 + 1]
    rgb[i * 3 + 2] = rgba[i * 4 + 2]
  }

  if (!modelPath || modelPath.endsWith('.pt')) {
    return errObj('ONNX 模型未找到，请先用 ultralytics 把 .pt 导出为 .onnx（导出步骤才需要一次 Python）。')
  }
  let session
  try {
    session = await ort.InferenceSession.create(modelPath, {
      executionProviders: ['cpu'],
      graphOptimizationLevel: 'all',
    })
  } catch (e) {
    return errObj('加载 ONNX 模型失败：' + e.message)
  }

  // letterbox 预处理到 imgsz
  const r = Math.min(imgsz / W, imgsz / H)
  const nw = Math.max(1, Math.round(W * r))
  const nh = Math.max(1, Math.round(H * r))
  const padXf = (imgsz - nw) / 2
  const padYf = (imgsz - nh) / 2
  const resized = bilinearResize(rgb, W, H, nw, nh)
  const input = new Float32Array(3 * imgsz * imgsz)
  input.fill(114 / 255)
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

  let output
  try {
    const tensor = new ort.Tensor('float32', input, [1, 3, imgsz, imgsz])
    const feeds = {}
    feeds[session.inputNames[0]] = tensor
    const res = await session.run(feeds)
    output = res[session.outputNames[0]]
  } catch (e) {
    return errObj('ONNX 推理失败：' + e.message)
  }

  // 解析输出 [1, C, A]
  const dims = output.dims
  const A = Math.max(dims[1], dims[2])
  const C = Math.min(dims[1], dims[2])
  const chanDimIs1 = dims[1] === C
  const data = output.data
  const valAt = (a, c) => (chanDimIs1 ? data[c * A + a] : data[a * C + c])
  const nc = C - 4
  if (nc < 1) {
    return errObj('模型输出通道数异常（C=' + C + '），无法解析类别。')
  }

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
        return errObj('裁剪图编码失败：' + e.message)
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

  items.sort((a, b) => a.box.y - b.box.y || a.box.x - b.box.x)
  const boxes = items.map((it) => it.box)
  const crops = noCrops ? [] : items.map((it) => it.crop)

  return { ok: true, image_width: W, image_height: H, boxes, crops }
}

// CLI 模式：node detect_onnx.cjs <model> [--conf 0.25] [--no-crops] ...
if (require.main === module) {
  ;(async () => {
    const argv = process.argv.slice(2)
    const positional = []
    const flags = {}
    for (let i = 0; i < argv.length; i++) {
      const a = argv[i]
      if (a.startsWith('--')) {
        if (i + 1 < argv.length && !argv[i + 1].startsWith('--')) {
          flags[a] = argv[i + 1]
          i++
        } else {
          flags[a] = true
        }
      } else {
        positional.push(a)
      }
    }
    const modelPath = positional[0] || 'models/ticket_detect.onnx'
    const getNum = (f, def) => (flags[f] !== undefined ? parseFloat(flags[f]) : def)
    const raw = (await readStdin()).trim()
    const result = await runDetect(raw, modelPath, {
      conf: getNum('--conf', 0.25),
      iou: getNum('--iou', 0.45),
      imgsz: getNum('--imgsz', IMGSZ),
      noCrops: !!flags['--no-crops'],
      noRotate: !!flags['--no-rotate'],
    })
    process.stdout.write(JSON.stringify(result) + '\n')
  })().catch((e) => {
    process.stdout.write(JSON.stringify({ ok: false, message: '运行异常：' + (e && e.message ? e.message : String(e)) }) + '\n')
  })
}

module.exports = { runDetect }
