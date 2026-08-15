#!/usr/bin/env node
// 用火山方舟（Volcano Engine Ark）Responses API 抽取手写小票结构化信息。
// 默认关闭思考（不传 reasoning）；仅当 fastMode=false 时才开启思考。
//
// 用法：
//   ARK_API_KEY=xxx node scripts/ark_extract.mjs [图片路径] [model]
//   ARK_API_KEY=xxx ARK_BASE_URL=... ARK_MODEL=... node scripts/ark_extract.mjs
//
// 说明：图片文件名清单不再塞进提示词（省 token）；如需把结果人名对齐到已知清单，
//       请在调用方对输出做相似度匹配（本项目由 correctPersonNames 负责）。

import { readFileSync } from 'node:fs'
import { resolve, dirname } from 'node:path'
import { fileURLToPath } from 'node:url'

const __dirname = dirname(fileURLToPath(import.meta.url))
const ARK_BASE_URL = process.env.ARK_BASE_URL || 'https://ark.cn-beijing.volces.com/api/v3/responses'
const ARK_API_KEY = process.env.ARK_API_KEY || ''
const MODEL = process.argv[3] || process.env.ARK_MODEL || 'doubao-seed-evolving'
const IMAGE = resolve(process.argv[2] || resolve(__dirname, '../pic/1.jpg'))

// 新的小票抽取提示词（与 src/main/ai-service.ts 中的 SINGLE_TICKET_PROMPT 保持一致）
const PROMPT = `任务：图片内是一张手写销售小票，请完成结构化信息抽取。

字段定义：
1. name：销售小票**右上角**的手写客户人名（注意是小票右上方的手写体，不是货品名、地址或店名），**必须有值、禁止置 null**；
2. date：小票手写日期，xx年xx月xx日，输出格式严格 YY-MM-DD，识别不到置 null
3. items：商品数组，识别不到商品则为空数组 []
   单品子字段：
   - name：品名（可能写在货品编号和品名区域），识别失败置 null
   - unit：单位，可能跨行填写，表示其余商品或表头共用了同一个单位，应沿用该共用单位（见硬性规则5）；
   - count：数量，可能跨行填写，表示其余商品相同数量，数字类型（见硬性规则5）；
   - price：单价，数字类型，识别失败置 null

硬性规则：
1. 最终只输出标准 JSON，禁止输出任何解释、说明、注释、markdown、多余文字，不能加 \`\`\`json 标记
2. JSON 顶层必须是数组，数组每一项对应一张小票
3. 商品品名、单位、数量、单价、金额等字段：识别不清、模糊、无法辨认的内容直接赋值 null，不要脑补猜测文字和数字；
4. 数字类字段只输出纯数字，不要带 元、个 等文字符号
5. 一张小票内的多个商品常常「共用」同一个单位或数量：手写时往往只在小票某处（表头、左侧、商品清单上方或一侧）写一次，而非每行都写。此时必须把这个共用的单位 / 数量**代入（继承）**到每一个没有单独写明单位 / 数量的商品行中——即该行 unit / count 填共用值，绝对不要因为「只写了一次」就把它置为 null。只有当某个商品确实连共用的单位 / 数量都不适用时，才允许置 null。
6. 输出 JSON 只允许包含 name、date、items 三个字段。
7. 严禁将**单位**（如 块、个、件、套、米、kg 等）填入 items[].name。name 只放品名/货品名称；若品名区域识别不清，name 填 null，而 unit 应填识别到的单位。`

function extractJson(s) {
  const t = (s || '').trim()
  if (t.startsWith('{') || t.startsWith('[')) {
    try { return JSON.parse(t) } catch {}
  }
  const m = t.match(/```(?:json)?\s*([\s\S]*?)```/i)
  if (m) { try { return JSON.parse(m[1].trim()) } catch {} }
  const b = t.match(/\{[\s\S]*\}/)
  if (b) { try { return JSON.parse(b[0]) } catch {} }
  throw new Error('无法从模型输出中解析出 JSON')
}

async function main() {
  if (!ARK_API_KEY) {
    console.error('缺少 ARK_API_KEY 环境变量。')
    console.error('用法: ARK_API_KEY=xxx node scripts/ark_extract.mjs [图片路径] [model]')
    process.exit(2)
  }

  let b64
  try {
    b64 = readFileSync(IMAGE).toString('base64')
  } catch (e) {
    console.error('读取图片失败:', e.message)
    process.exit(1)
  }
  if (b64.length > 6_000_000) {
    console.warn(`警告：图片 base64 约 ${(b64.length / 1e6).toFixed(1)}MB，可能超出接口限制，建议先压缩。`)
  }

  const dataUri = `data:image/jpeg;base64,${b64}`
  const body = {
    model: MODEL,
    input: [
      { role: 'system', content: '你是一位严谨的中文票据识别助手，只输出合法 JSON，且顶层必须是一个数组。' },
      {
        role: 'user',
        content: [
          { type: 'input_image', image_url: dataUri },
          { type: 'input_text', text: PROMPT },
        ],
      },
    ],
    // 关闭深度思考（doubao-seed-evolving 默认开启思考，会吃掉大量输出 token 且不吐最终文本）
    thinking: { type: 'disabled' },
    max_output_tokens: 4096,
    temperature: 0.2,
  }

  console.error(`→ Ark Responses API: ${ARK_BASE_URL}\n→ model: ${MODEL}`)
  const res = await fetch(ARK_BASE_URL, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json', Authorization: `Bearer ${ARK_API_KEY}` },
    body: JSON.stringify(body),
  })
  const text = await res.text()
  if (!res.ok) {
    console.error(`HTTP ${res.status}: ${text.slice(0, 600)}`)
    process.exit(1)
  }

  let json
  try { json = JSON.parse(text) } catch {
    console.error('返回非 JSON:\n' + text.slice(0, 600))
    process.exit(1)
  }
  const output = Array.isArray(json.output) ? json.output : []
  let content = ''
  for (const item of output) {
    if (item.type === 'message' && Array.isArray(item.content)) {
      for (const c of item.content) {
        if ((c.type === 'output_text' || c.type === 'text') && typeof c.text === 'string') content += c.text
      }
    }
  }
  if (!content) {
    console.error('未从响应中提取到文本:\n' + text.slice(0, 600))
    process.exit(1)
  }

  const parsed = extractJson(content)
  process.stdout.write(JSON.stringify(parsed, null, 2) + '\n')
}

main().catch((e) => { console.error('运行失败:', e); process.exit(1) })
