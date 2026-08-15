// 定量对比：整图单次调用 vs 逐张裁剪 20 次调用 的 token 消耗
// 忠实复刻 app 实际发送内容（提示词/坐标/图片尺寸），抓 Ark Responses API 的 usage。
import fs from 'node:fs'
import path from 'node:path'
import { createRequire } from 'node:module'
const require = createRequire(import.meta.url)

const ARK_API_KEY = process.env.ARK_API_KEY
const MODEL = process.env.ARK_MODEL || 'doubao-seed-evolving'
const PIC = path.resolve('pic/1.jpg')

// ---- 1. 从 ai-service.ts 抽取真实提示词（与 app 完全一致）----
const aiSrc = fs.readFileSync('src/main/ai-service.ts', 'utf8')
function extractConst(name) {
  const re = new RegExp('export const ' + name + ' = `([\\s\\S]*?)`(\\n|$)', 'm')
  const m = aiSrc.match(re)
  if (!m) throw new Error('未找到 ' + name)
  return m[1]
}
const DEFAULT_PROMPT = extractConst('DEFAULT_PROMPT')
const SINGLE_TICKET_PROMPT = extractConst('SINGLE_TICKET_PROMPT')

function buildDetectPrompt(prompt, boxes, w, h) {
  if (!boxes.length) return prompt
  const lines = boxes
    .map((b, i) => `${i + 1}: 左上角(x=${Math.round(b.x)}, y=${Math.round(b.y)}) 宽高(w=${Math.round(b.w)}, h=${Math.round(b.h)})`)
    .join('\n')
  return (
    `${prompt}\n\n` +
    `（辅助定位）本地检测已在原图中定位到 ${boxes.length} 张小票，原图尺寸 ${w}×${h}（像素）。` +
    `每张小票的边界框（左上角坐标 + 宽高，单位像素）如下：\n${lines}\n` +
    `请依据这些框的位置，在同一张图里逐张识别每张小票；确保结果数组恰好包含 ${boxes.length} 项，` +
    `第 i 项对应编号 i 的框。`
  )
}

// ---- 2. 检测：取 boxes + crops（与 app 一致）----
const det = require('./detect_onnx.cjs')
const picB64 = fs.readFileSync(PIC).toString('base64')
const dres = await det.runDetect(picB64, 'models/ticket_detect.onnx', { conf: 0.25, iou: 0.45 })
if (!dres.ok) { console.error('检测失败:', dres.message); process.exit(1) }
const boxes = dres.boxes
const W = dres.image_width, H = dres.image_height
console.error(`检测到 ${boxes.length} 张小票，原图 ${W}x${H}`)

// ---- 3. Ark Responses API 单次调用，返回 usage ----
function callUsage(dataUri, promptText, maxOut) {
  return new Promise((resolve, reject) => {
    const body = JSON.stringify({
      model: MODEL,
      input: [
        { role: 'system', content: '你是一位严谨的中文票据识别助手，只输出合法 JSON，且顶层必须是一个数组。' },
        { role: 'user', content: [ { type: 'input_image', image_url: dataUri }, { type: 'input_text', text: promptText } ] },
      ],
      max_output_tokens: maxOut,
      temperature: 0.2,
      thinking: { type: 'disabled' },
    })
    const req = require('node:https').request(
      'https://ark.cn-beijing.volces.com/api/v3/responses',
      { method: 'POST', headers: { Authorization: 'Bearer ' + ARK_API_KEY, 'Content-Type': 'application/json' } },
      (res) => {
        let d = ''; res.on('data', (c) => (d += c))
        res.on('end', () => {
          try {
            const j = JSON.parse(d)
            if (j.error) return reject(new Error(j.error.message || JSON.stringify(j.error)))
            resolve(j.usage || { input_tokens: 0, output_tokens: 0, total_tokens: 0, note: 'no-usage' })
          } catch (e) { reject(new Error('解析失败: ' + d.slice(0, 200))) }
        })
      })
    req.on('error', reject)
    req.write(body); req.end()
  })
}

// ---- 4a. 整图单次 ----
const t0 = Date.now()
const wholeUsage = await callUsage('data:image/jpeg;base64,' + picB64, buildDetectPrompt(DEFAULT_PROMPT, boxes, W, H), 4096)
const wholeMs = Date.now() - t0
console.error(`整图调用完成 ${wholeMs}ms`)

// ---- 4b. 逐张裁剪 20 次 ----
const sum = { input_tokens: 0, output_tokens: 0, total_tokens: 0 }
const perCrop = []
let failed = 0
for (let i = 0; i < dres.crops.length; i++) {
  const uri = 'data:image/jpeg;base64,' + dres.crops[i]
  try {
    const u = await callUsage(uri, SINGLE_TICKET_PROMPT, 2048)
    sum.input_tokens += u.input_tokens
    sum.output_tokens += u.output_tokens
    sum.total_tokens += u.total_tokens
    perCrop.push(u)
  } catch (e) {
    failed++; console.error(`crop#${i + 1} 失败: ${e.message}`)
  }
}
console.error(`逐张调用完成，失败 ${failed} 张`)

// ---- 5. 汇总 ----
const wholePx = W * H
const cropPxSum = dres.crops.reduce((a, c) => {
  const buf = Buffer.from(c, 'base64')
  let off = 2
  while (off < buf.length) { if (buf[off] !== 0xFF) break; const m = buf[off + 1]; if (m >= 0xC0 && m <= 0xCF && m !== 0xC4 && m !== 0xC8 && m !== 0xCC) return a + buf.readUInt16BE(off + 5) * buf.readUInt16BE(off + 7); off += 2 + buf.readUInt16BE(off + 2) }
  return a
}, 0)

const out = {
  image: { whole_px: wholePx, crop_px_sum: cropPxSum, crop_px_ratio: +(cropPxSum / wholePx).toFixed(3), crop_count: boxes.length },
  whole_image_single_call: { ...wholeUsage, ms: wholeMs, calls: 1 },
  per_crop_20_calls: { ...sum, calls: boxes.length, failed },
  comparison: {
    whole_total_tokens: wholeUsage.total_tokens,
    percrop_total_tokens: sum.total_tokens,
    percrop_total_minus_whole: sum.total_tokens - wholeUsage.total_tokens,
    percrop_is_x_times_whole: +(sum.total_tokens / Math.max(1, wholeUsage.total_tokens)).toFixed(2),
    whole_saves_tokens_vs_percrop: sum.total_tokens - wholeUsage.total_tokens,
  },
}
process.stdout.write(JSON.stringify(out, null, 2) + '\n')
