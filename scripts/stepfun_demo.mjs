#!/usr/bin/env node
// 验证脚本：用「现有 stepfun 配置」（config.json 里的 aiConfig）跑一遍新的小票抽取提示词，
// 证明新提示词在真实图片上可正常抽取。仅用于验证，默认接入仍是 Ark（见 ark_extract.mjs）。
//
// 用法：node scripts/stepfun_demo.mjs [图片路径] [config.json 路径]

import { readFileSync } from 'node:fs'
import { resolve, dirname } from 'node:path'
import { fileURLToPath } from 'node:url'
import { homedir } from 'node:os'

const __dirname = dirname(fileURLToPath(import.meta.url))
const IMAGE = resolve(process.argv[2] || resolve(__dirname, '../pic/1.jpg'))
const CONFIG = resolve(process.argv[3] || resolve(homedir(), 'Library/Application Support/bill-explorer/config.json'))

const PROMPT = `任务：图片内有多张手写销售小票，请先分割出每一张独立小票，逐张完成结构化信息抽取。

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
7. 不要合并多张小票的数据，一张小票对应数组内一个对象。`

function extractJson(s) {
  const t = (s || '').trim()
  if (t.startsWith('{') || t.startsWith('[')) { try { return JSON.parse(t) } catch {} }
  const m = t.match(/```(?:json)?\s*([\s\S]*?)```/i)
  if (m) { try { return JSON.parse(m[1].trim()) } catch {} }
  const b = t.match(/\{[\s\S]*\}/)
  if (b) { try { return JSON.parse(b[0]) } catch {} }
  throw new Error('无法解析 JSON')
}

async function main() {
  let cfg
  try { cfg = JSON.parse(readFileSync(CONFIG, 'utf8')) } catch (e) {
    console.error('读取 config.json 失败:', e.message); process.exit(1)
  }
  const ai = cfg.settings?.aiConfig
  if (!ai?.apiKey || !ai?.baseURL) { console.error('config.json 中缺少 aiConfig.apiKey / baseURL'); process.exit(1) }

  const b64 = readFileSync(IMAGE).toString('base64')
  const dataUri = `data:image/jpeg;base64,${b64}`
  const url = `${ai.baseURL.replace(/\/$/, '')}/chat/completions`
  const body = {
    model: ai.model,
    temperature: 0.2,
    max_tokens: 8000,
    // 关闭思考：否则推理模型会把答案塞进 reasoning 字段、content 为空
    enable_thinking: false,
    reasoning: { enabled: false },
    messages: [
      { role: 'system', content: '你是一位严谨的中文票据识别助手，只输出合法 JSON，且顶层必须是一个数组。' },
      { role: 'user', content: [{ type: 'text', text: PROMPT }, { type: 'image_url', image_url: { url: dataUri } }] },
    ],
  }

  console.error(`→ stepfun chat/completions: ${url}\n→ model: ${ai.model}`)
  const res = await fetch(url, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json', Authorization: `Bearer ${ai.apiKey}` },
    body: JSON.stringify(body),
  })
  const text = await res.text()
  if (!res.ok) { console.error(`HTTP ${res.status}: ${text.slice(0, 600)}`); process.exit(1) }
  const json = JSON.parse(text)
  const msg = json?.choices?.[0]?.message || {}
  // stepfun 的推理模型把最终答案放在 reasoning 字段、content 可能为空，这里做兜底
  const content = msg.content || msg.reasoning || msg.reasoning_content || ''
  if (!content) { console.error('返回无 content/reasoning:\n' + text.slice(0, 600)); process.exit(1) }
  let parsed
  try {
    parsed = extractJson(content)
  } catch {
    // 推理模型有时把 JSON 藏在 reasoning 末尾，做一次兜底截取
    const m = content.match(/\[\s*\{[\s\S]*\}\s*\]\s*$/) || content.match(/\[[\s\S]*\]/)
    if (m) { try { parsed = JSON.parse(m[0]) } catch { parsed = null } }
  }
  if (!parsed) { console.error('无法解析出 JSON，原始内容末尾：\n' + content.slice(-800)); process.exit(1) }
  process.stdout.write(JSON.stringify(parsed, null, 2) + '\n')
}

main().catch((e) => { console.error('运行失败:', e); process.exit(1) })
