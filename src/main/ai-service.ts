// AI 小票识别服务：调用 OpenAI 兼容的 chat/completions 接口，
// 把图片里的手写/印刷小票内容识别成结构化账单字段。

import { scanDirectory } from './file-service'
import { readImageBase64, downscaleImageBase64 } from './image-service'
import { detectTickets, type DetectedBox } from './detection'
import { fuzzyMatchName } from '../shared/person'

function getFetch() {
  return (globalThis as Record<string, unknown>).fetch as (
    input: string,
    init?: Record<string, unknown>,
  ) => Promise<unknown>
}

export interface AIConfig {
  baseURL: string
  apiKey: string
  model: string
  temperature: number
  /** 快速模式：关闭模型「思考/推理」(reasoning)，显著加快响应；默认开启 */
  fastMode?: boolean
}

export interface AIRecognizedRow {
  no?: string | number
  date?: string
  name?: string
  unit?: string
  qty?: string | number
  price?: string | number
  amount?: string | number
  person?: string
  remark?: string
  /** 仅前端使用：记录该行识别结果来自哪张图片，便于核对（AI 不会返回此字段） */
  source?: string
  /** 人名是否已被「人名清单」自动修正过（前端展示用） */
  personCorrected?: boolean
}

export const DEFAULT_PROMPT = `任务：图片内有多张手写销售小票，请先分割出每一张独立小票，逐张完成结构化信息抽取。

字段定义：
1. name：销售小票**上方或右上角**的手写客户人名（注意是小票上方 / 右上方的手写体，不是货品名、地址或店名），**必须有值、禁止置 null**；手写模糊 / 潦草 / 认不出时，必须从文末【已知客户人名清单】中挑选最贴近的人名填入（在清单内猜测，不要凭空臆造清单外的名字；仅当确实不像任何已知人名时才允许输出新名）
2. date：小票手写日期，输出格式严格 YY-MM-DD，识别不到置 null
3. items：商品数组，识别不到商品则为空数组 []
   单品子字段：
   - name：商品品名，识别失败置 null
   - unit：单位，若本行未单独写明、但本票其余商品或表头共用了同一个单位，应沿用该共用单位（见硬性规则6）；确实没有则置 null
   - count：数量，数字类型，若本行未单独写明、但本票共用了同一个数量，应沿用（见硬性规则6）；确实没有则置 null
   - price：单价，数字类型，识别失败置 null
   - amount：金额，数字类型，识别失败置 null

硬性规则：
1. 最终只输出标准 JSON，禁止输出任何解释、说明、注释、markdown、多余文字，不能加 \`\`\`json 标记
2. JSON 顶层必须是数组，数组每一项对应一张小票
3. 商品品名、单位、数量、单价、金额、日期等字段：识别不清、模糊、无法辨认的内容直接赋值 null，不要脑补猜测文字和数字；**唯独 name（客户人名）例外——必须填充，模糊时从【已知客户人名清单】中选取最相似的人名，禁止置 null**
4. 不要合并多张小票的数据，一张小票对应数组内一个对象
5. 数字类字段只输出纯数字，不要带 元、个 等文字符号
6. 一张小票内的多个商品常常「共用」同一个单位或数量：手写时往往只在小票某处（表头、左侧、商品清单上方或一侧）写一次，而非每行都写。此时必须把这个共用的单位 / 数量**代入（继承）**到每一个没有单独写明单位 / 数量的商品行中——即该行 unit / count 填共用值，绝对不要因为「只写了一次」就把它置为 null。只有当某个商品确实连共用的单位 / 数量都不适用时，才允许置 null。
7. 严禁输出 person / 调货人 / 供货人 / 经手人 / 备注（remark）等任何额外字段。小票上若出现调货人、供货方、经手人、备注等信息，一律忽略，不要写进结果；客户人名只写在 name 字段。
8. 输出 JSON 只允许包含 name、date、items 三个字段；不要再新增 person、调货人、备注 等任何额外字段，也不要把调货人 / 供货人当成客户人名。`

// 单张小票（已裁剪出来）的识别提示词：图里只有一张票，让 AI 聚焦单票，减少背景干扰。
// 字段定义与硬性规则与 DEFAULT_PROMPT 保持一致，仅去掉"先分割多张"的步骤。
export const SINGLE_TICKET_PROMPT = `任务：这是一张已被裁剪出来的单独手写销售小票。请识别其中的客户人名与商品清单，完成结构化信息抽取。

字段定义：
1. name：销售小票**上方或右上角**的手写客户人名（注意是小票上方 / 右上方的手写体，不是货品名、地址或店名），**必须有值、禁止置 null**；手写模糊 / 潦草 / 认不出时，必须从文末【已知客户人名清单】中挑选最贴近的人名填入（在清单内猜测，不要凭空臆造清单外的名字；仅当确实不像任何已知人名时才允许输出新名）
2. date：小票手写日期，输出格式严格 YY-MM-DD，识别不到置 null
3. items：商品数组，识别不到商品则为空数组 []
   单品子字段：
   - name：商品品名，识别失败置 null
   - unit：单位，若本行未单独写明、但本票其余商品或表头共用了同一个单位，应沿用该共用单位（见硬性规则6）；确实没有则置 null
   - count：数量，数字类型，若本行未单独写明、但本票共用了同一个数量，应沿用（见硬性规则6）；确实没有则置 null
   - price：单价，数字类型，识别失败置 null
   - amount：金额，数字类型，识别失败置 null

硬性规则：
1. 最终只输出标准 JSON，禁止输出任何解释、说明、注释、markdown、多余文字，不能加 \`\`\`json 标记
2. JSON 顶层必须是数组，即使只有一张小票也要用数组（数组长度为 1）
3. 商品品名、单位、数量、单价、金额、日期等字段：识别不清、模糊、无法辨认的内容直接赋值 null，不要脑补；**唯独 name（客户人名）例外——必须填充，模糊时从【已知客户人名清单】中选取最相似的人名，禁止置 null**
4. 一张小票对应数组内一个对象
5. 数字类字段只输出纯数字，不要带 元、个 等文字符号
6. 同一张小票内的多个商品常常「共用」同一个单位或数量：手写时往往只在小票某处（表头、左侧、商品清单上方或一侧）写一次，而非每行都写。此时必须把这个共用的单位 / 数量**代入（继承）**到每一个没有单独写明单位 / 数量的商品行中——即该行 unit / count 填共用值，绝对不要因为「只写了一次」就把它置为 null。只有当某个商品确实连共用的单位 / 数量都不适用时，才允许置 null。
7. 严禁输出 person / 调货人 / 供货人 / 经手人 / 备注（remark）等任何额外字段。小票上若出现调货人、供货方、经手人、备注等信息，一律忽略，不要写进结果；客户人名只写在 name 字段。
8. 输出 JSON 只允许包含 name、date、items 三个字段；不要再新增 person、调货人、备注 等任何额外字段，也不要把调货人 / 供货人当成客户人名。`

// 单次请求超时（毫秒）。小票图识别通常较慢，给足 120s。
const AI_TIMEOUT_MS = 120_000
// 输出 token 上限：整图识别（小票较多）用较高值，单张小票识别用较低值以省 token。
const MAX_TOKENS = 4096
const MAX_TOKENS_CROP = 2048
// 网络抖动 / 5xx / 429 的退避重试次数（不含首次）。
const RETRY_COUNT = 2
// 单张裁剪图送视觉模型前最长边上限（px）：降低视觉 token 消耗与耗时。
const CROP_MAX_DIM = 1280
const CROP_QUALITY = 80

function isValidConfig(c?: AIConfig): boolean {
  return !!c && typeof c.baseURL === 'string' && c.baseURL.trim() !== '' &&
    typeof c.apiKey === 'string' && c.apiKey.trim() !== '' &&
    typeof c.model === 'string' && c.model.trim() !== ''
}

function sleep(ms: number) {
  return new Promise((resolve) => setTimeout(resolve, ms))
}

// 可重试的错误：网络异常、服务端 5xx、限流 429 等。
class RetryableError extends Error {}

function extractJson(text: string): unknown {
  const trimmed = text.trim()
  // 1. 整个内容就是 JSON
  if (trimmed.startsWith('{') || trimmed.startsWith('[')) {
    try {
      return JSON.parse(trimmed)
    } catch {
      // fallthrough
    }
  }
  // 2. 从 ```json ... ``` 提取
  const codeMatch = trimmed.match(/```(?:json)?\s*([\s\S]*?)```/i)
  if (codeMatch) {
    try {
      return JSON.parse(codeMatch[1].trim())
    } catch {
      // fallthrough
    }
  }
  // 3. 从文本中找第一个 { ... }
  const braceMatch = trimmed.match(/\{[\s\S]*\}/)
  if (braceMatch) {
    try {
      return JSON.parse(braceMatch[0])
    } catch {
      // fallthrough
    }
  }
  throw new Error('AI 返回内容不是可解析的 JSON')
}

// 把 AI 返回的（可能是嵌套 items 的新版结构，或扁平旧版结构）标准化为扁平的账单行。
// 新版：顶层数组，每项 = { name(客户人名), date, items: [{name,unit,count,price,amount}] }
// 旧版：每项含 name(货品)/person/unit/qty/price/amount 等扁平字段（无 items）
function normalizeRows(parsed: unknown): AIRecognizedRow[] {
  if (!parsed || typeof parsed !== 'object') return []
  const data = parsed as Record<string, unknown>

  let receipts: unknown[] = []
  if (Array.isArray(parsed)) receipts = parsed
  else if (Array.isArray(data.receipts)) receipts = data.receipts
  else if (Array.isArray(data.rows)) receipts = data.rows
  else if (Array.isArray(data.items)) receipts = data.items
  else receipts = [data] // 整段就是一张小票

  const out: AIRecognizedRow[] = []
  for (const r of receipts) {
    if (!r || typeof r !== 'object') continue
    const rec = r as Record<string, unknown>
    // 客户人名取顶层 name 字段（提示词规定客户名只写在 name）；
    // 仅当无 name 时（旧版扁平格式把客户名放在 person）才回退到 person，
    // 这样可避免模型把票面上的「调货人/供货人」写入 person 后覆盖真正的客户名。
    const person = strField(rec.name) || strField(rec.person)
    const date = strField(rec.date)
    const items = Array.isArray(rec.items) ? (rec.items as Record<string, unknown>[]) : null
    if (items && items.length) {
      for (const it of items) out.push(flattenItem(person, date, it))
    } else {
      // 没有商品数组：把该对象整体当作一条商品记录
      out.push(flattenItem(person, date, rec))
    }
  }
  return out
}

function strField(v: unknown): string {
  if (v == null) return ''
  return String(v)
}

// 解析 YY-MM-DD / YYYY-MM-DD 为 Date；非法返回 null
function parseYmd(s: string): Date | null {
  const m = (s || '').trim().match(/^(\d{2,4})-(\d{1,2})-(\d{1,2})$/)
  if (!m) return null
  const y = Number(m[1])
  const mo = Number(m[2])
  const d = Number(m[3])
  if (mo < 1 || mo > 12 || d < 1 || d > 31) return null
  const year = y < 100 ? 2000 + y : y
  const dt = new Date(year, mo - 1, d)
  return isNaN(dt.getTime()) ? null : dt
}

// 两个日期相差天数（绝对值）
function dayDiff(a: Date, b: Date): number {
  const MS = 24 * 3600 * 1000
  return Math.round((a.getTime() - b.getTime()) / MS)
}

// 日期修正：同一张图里的多张小票大概率是同一天（最多相差几天）。
// 以「第一张填了日期的小票」为参照：
//   - 后续小票缺失日期 → 直接填入参照日；
//   - 与参照日相差 ≤2 天（视为同一天的 OCR 误差）→ 修正为参照日；
//   - 与参照日相差 >2 天 → 保留原值（确为不同日期）。
function normalizeTicketDates(tickets: RecognizedTicket[]): RecognizedTicket[] {
  let ref = ''
  for (const t of tickets) {
    const d = t.rows.find((r) => (r.date || '').trim())?.date
    if (d) {
      ref = d.trim()
      break
    }
  }
  if (!ref) return tickets
  const refTime = parseYmd(ref)
  return tickets.map((t) => ({
    ...t,
    rows: t.rows.map((r) => {
      const d = (r.date || '').trim()
      if (!d) return { ...r, date: ref }
      if (!refTime) return { ...r, date: ref }
      const dt = parseYmd(d)
      if (dt && Math.abs(dayDiff(refTime, dt)) <= 2) return { ...r, date: ref }
      return r
    }),
  }))
}

function flattenItem(person: string, date: string, it: Record<string, unknown>): AIRecognizedRow {
  return {
    no: '',
    date,
    // 商品品名：新版取 items[].name；旧版兜底 goods/item
    name: strField(it.name ?? it.goods ?? it.item),
    unit: strField(it.unit),
    // 数量：新版 count；旧版 qty/quantity
    qty: strField(it.count ?? it.qty ?? it.quantity),
    price: strField(it.price),
    amount: strField(it.amount ?? it.money ?? it.total),
    person,
    remark: strField(it.remark),
    source: strField(it.source),
  }
}

// ============ 人名清单（遍历账单目录，取文件名中的人名） ============

// 遍历账单根目录，从每个 .xlsx/.xls 文件名中解析出客户人名，去重排序后返回清单。
// 清单用于：① 附加到提示词，让 AI 在已知范围内识别；② 识别后把人名修正为清单中最接近的规范写法。
export async function buildNameList(workDir: string): Promise<string[]> {
  if (!workDir) return []
  try {
    const res = await scanDirectory(workDir)
    const files = res.files || []
    const set = new Set<string>()
    for (const f of files) {
      // 取完整文件名（不含扩展名）作为清单项；
      // ~ 开头的 Excel 临时锁文件、. 开头的隐藏文件已由 scanDirectory 过滤
      const base = f.fileName.replace(/\.(xlsx|xls)$/i, '')
      if (!base) continue
      set.add(base)
    }
    return Array.from(set).sort((a, b) => a.localeCompare(b, 'zh'))
  } catch {
    return []
  }
}

// 把人名清单附加到用户提示词之后（若为空则不附加）。
export function buildAugmentedPrompt(basePrompt: string, nameList: string[]): string {
  const prompt = basePrompt && basePrompt.trim() ? basePrompt : DEFAULT_PROMPT
  if (!nameList.length) return prompt
  const listText = nameList.join('、')
  const section = `

【已知账单文件名清单（客户 / 单据标识，务必在此范围内匹配）】
以下是从账单目录的完整 Excel 文件名整理出的已知清单，共 ${nameList.length} 个：
${listText}
注意：清单项为完整文件名，可能包含前缀编号、地点等修饰（如 "DB210 袁文志"、"1日 华信泰博 1716"）。请识别其中的客户人名，并将结果对应到清单中字形 / 读音最接近的一项——**人名可以是清单项的一部分，不必与清单项字面完全一致**（例如清单 "DB210 袁文志" 对应人名 "袁文志"）。请把结果修正为清单中的完整写法（含前缀编号）。只有当识别的人名明显是清单中完全没有的新客户时，才允许保留原识别结果。`
  return prompt + section
}

// 识别结果回填前，把每行的人名按清单修正为最接近的一项；返回修正后的行与修正次数。
export function correctPersonNames(
  rows: AIRecognizedRow[],
  nameList: string[],
): { rows: AIRecognizedRow[]; corrected: number } {
  if (!nameList.length || !rows.length) return { rows, corrected: 0 }
  let corrected = 0
  const out = rows.map((r) => {
    const p = (r.person || '').trim()
    if (!p) return r
    // 模糊匹配：识别名不必与清单项字面完全一致（人名可能只是清单项的一部分）
    const fixed = fuzzyMatchName(p, nameList).matched
    if (fixed && fixed !== p) {
      corrected++
      return { ...r, person: fixed, personCorrected: true }
    }
    return r
  })
  return { rows: out, corrected }
}

// 单次请求的结果
type ReqResult =
  | { status: 'ok'; rows: AIRecognizedRow[]; note?: string }
  | { status: 'imageVariant'; error: string } // 图片内容格式被服务端拒绝，需换格式重试
  | { status: 'jsonMode'; error: string } // 图片格式 OK，但 response_format=json_object 不被支持
  | { status: 'fail'; message: string }

async function doRequest(
  url: string,
  config: AIConfig,
  body: Record<string, unknown>,
  signal: AbortSignal,
): Promise<ReqResult> {
  let responseText = ''
  try {
    const res = (await getFetch()(url, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        Authorization: `Bearer ${config.apiKey}`,
      },
      body: JSON.stringify(body),
      signal: signal as unknown as AbortSignal,
    })) as { ok: boolean; status: number; statusText: string; text: () => Promise<string> }
    responseText = await res.text()

    if (!res.ok) {
      const text = responseText
      // 服务端拒绝当前图片格式（常见 400：unknown variant image_url），换下一种格式。
      const imageVariant =
        res.status === 400 && /unknown variant/i.test(text) && !/unknown variant `text`/.test(text)
      if (imageVariant) {
        return { status: 'imageVariant', error: `AI 接口错误 ${res.status}：${text.slice(0, 200)}` }
      }
      // 图片格式 OK，但 response_format=json_object 不被支持 → 退回纯 prompt 重试同格式。
      const jsonMode = res.status === 400 && /response_format|json_object|json mode|'json'|"json"/i.test(text)
      if (jsonMode) {
        return { status: 'jsonMode', error: `AI 接口错误 ${res.status}：${text.slice(0, 200)}` }
      }
      // 限流 / 服务端错误：交给上层退避重试。
      if (res.status === 429 || (res.status >= 500 && res.status < 600)) {
        throw new RetryableError(`HTTP ${res.status}`)
      }
      return { status: 'fail', message: `AI 接口错误 ${res.status}：${text.slice(0, 300)}` }
    }
  } catch (err) {
    if (err instanceof RetryableError) throw err
    // 其它（网络异常、超时 abort、JSON 解析异常等）
    const aborted = (err as { name?: string })?.name === 'AbortError'
    const msg = aborted
      ? `请求超时（>${AI_TIMEOUT_MS / 1000}s），请检查网络或换用更轻量的图片`
      : err instanceof Error
        ? err.message
        : '未知错误'
    return { status: 'fail', message: '请求 AI 接口失败：' + msg }
  }

  // 成功响应，解析结果
  let responseJson: unknown
  try {
    responseJson = JSON.parse(responseText)
  } catch {
    return { status: 'fail', message: 'AI 接口返回非 JSON：' + responseText.slice(0, 300) }
  }

  const data = responseJson as Record<string, unknown>
  const choices = data.choices as Array<Record<string, unknown>> | undefined
  const first = choices && choices[0]
  const message = first?.message as Record<string, unknown> | undefined
  const content = message?.content
  if (typeof content !== 'string') {
    return { status: 'fail', message: 'AI 接口返回内容为空或格式异常' }
  }

  let parsed: unknown
  try {
    parsed = extractJson(content)
  } catch (err) {
    return {
      status: 'fail',
      message: 'AI 返回无法解析为 JSON：' + (err instanceof Error ? err.message : '') + '\n原始内容：' + content.slice(0, 400),
    }
  }

  const rows = normalizeRows(parsed)
  if (rows.length === 0) {
    return { status: 'fail', message: 'AI 未识别出任何小票记录，请检查图片清晰度或提示词。' }
  }

  // 输出被 token 上限截断：返回结果但给出提示，避免用户误以为已全识别。
  const finish = first?.finish_reason
  const note =
    finish === 'length'
      ? '识别结果可能被截断（输出达到 token 上限），建议拆分图片或降低单图小票数量。'
      : undefined

  return { status: 'ok', rows, note }
}

// 带超时 + 退避重试地执行一次 doRequest。
async function callWithRetry(
  fn: (signal: AbortSignal) => Promise<ReqResult>,
): Promise<ReqResult> {
  let lastErr: unknown
  for (let attempt = 0; attempt <= RETRY_COUNT; attempt++) {
    const controller = new AbortController()
    const timer = setTimeout(() => controller.abort(), AI_TIMEOUT_MS)
    try {
      const r = await fn(controller.signal)
      clearTimeout(timer)
      return r
    } catch (err) {
      clearTimeout(timer)
      lastErr = err
      if (err instanceof RetryableError) {
        if (attempt < RETRY_COUNT) {
          await sleep(800 * (attempt + 1))
          continue
        }
        return { status: 'fail', message: '请求 AI 接口失败（服务异常/限流）：' + (err.message || '未知错误') }
      }
      // 非可重试异常（如 AbortError）：直接失败，不再重试。
      return { status: 'fail', message: '请求 AI 接口失败：' + (err instanceof Error ? err.message : '未知错误') }
    }
  }
  return { status: 'fail', message: '请求 AI 接口失败：' + (lastErr instanceof Error ? lastErr.message : '未知错误') }
}

export async function recognizeReceipt(
  imageBase64: string,
  config: AIConfig,
  prompt = DEFAULT_PROMPT,
  opts?: { maxTokens?: number },
): Promise<{ error?: boolean; message?: string; rows?: AIRecognizedRow[] }> {
  if (!isValidConfig(config)) {
    return { error: true, message: 'AI 接口未配置：请在设置中填写 Base URL、API Key 和模型。' }
  }

  const baseURL = config.baseURL.replace(/\/$/, '')
  const url = `${baseURL}/chat/completions`

  // 部分 OpenAI 兼容 API 不支持标准 image_url 内容格式（如：unknown variant `image_url`, expected `text`）。
  // 按兼容性从高到低尝试多种图片内容格式，服务端 400 拒绝后自动重试下一种。
  const dataUri = `data:image/jpeg;base64,${imageBase64}`
  const imageContentVariants = [
    { type: 'image_url', image_url: { url: dataUri } }, // 标准 OpenAI 格式
    { type: 'image_url', image_url: dataUri }, // 部分 API 期望 image_url 为字符串
    { type: 'image', image: dataUri }, // 部分本地/代理 API 使用 image 类型
    { type: 'image', image: imageBase64 }, // 部分 API 使用 image 类型 + 纯 base64
  ]

  const temperature = typeof config.temperature === 'number' ? config.temperature : 0.2
  const maxTokens = opts?.maxTokens ?? MAX_TOKENS
  // 快速模式（默认开启）：关闭模型「思考/推理」，显著加快首 token 与整体响应、大幅省 token。
  // - OpenRouter 系：reasoning:{enabled:false}
  // - 阿里云百炼（DashScope）等 Qwen 系：enable_thinking:false（qwen3.8-max 等默认开启思考，
  //   是慢且耗 token 的主因）。两者同时带上；若接口拒绝其中某个参数，下方会检测后去掉重试。
  const useReasoning = config.fastMode !== false
  let lastError = ''

  for (let v = 0; v < imageContentVariants.length; v++) {
    const variant = imageContentVariants[v]
    // 构造请求体；withReasoning 控制是否附带关闭思考的参数
    const buildBody = (withReasoning: boolean): Record<string, unknown> => {
      const body: Record<string, unknown> = {
        model: config.model,
        temperature,
        max_tokens: maxTokens,
        messages: [
          {
            role: 'system',
            content: '你是一位严谨的中文票据识别助手，只输出合法 JSON，且顶层必须是一个数组。',
          },
          {
            role: 'user',
            content: [{ type: 'text', text: prompt || DEFAULT_PROMPT }, variant],
          },
        ],
      }
      if (withReasoning) {
        body.reasoning = { enabled: false }
        body.enable_thinking = false
      }
      return body
    }

    // 快速模式：先带关闭思考参数试；若服务端拒绝（报错含 reasoning/thinking 字样），去掉后重试一次
    let result = await callWithRetry((signal) => doRequest(url, config, buildBody(useReasoning), signal))
    if (useReasoning && result.status === 'fail' && /reasoning|thinking|enable_thinking/i.test(result.message || '')) {
      result = await callWithRetry((signal) => doRequest(url, config, buildBody(false), signal))
    }

    if (result.status === 'imageVariant') {
      lastError = result.error
      break // 该图片格式被拒，试下一种格式
    }
    if (result.status === 'jsonMode') {
      lastError = result.error
      continue // 图片格式可用，但 JSON 模式不支持，退回纯 prompt 重试同格式
    }
    if (result.status === 'fail') {
      return { error: true, message: result.message }
    }
    // status === 'ok'
    return result.note ? { rows: result.rows, message: result.note } : { rows: result.rows }
  }

  // 所有格式都被服务端拒绝
  return {
    error: true,
    message: `AI 接口不支持图片输入（尝试了 ${imageContentVariants.length} 种格式均被拒绝）。\n${lastError}\n请检查：① 使用支持 vision 的模型；② 确认 Base URL 为正确的 OpenAI 兼容视觉接口。`,
  }
}

// 有限并发执行：对 crops 逐张调用 AI，避免一次性打爆接口，也不把单张失败拖垮整体。
async function mapLimit<T, R>(
  items: T[],
  limit: number,
  fn: (item: T, index: number) => Promise<R>,
): Promise<R[]> {
  const results: R[] = new Array(items.length)
  let cursor = 0
  async function worker(): Promise<void> {
    while (cursor < items.length) {
      const i = cursor++
      try {
        results[i] = await fn(items[i], i)
      } catch {
        results[i] = undefined as unknown as R
      }
    }
  }
  const workers = Array.from({ length: Math.min(limit, items.length) }, () => worker())
  await Promise.all(workers)
  return results
}

// 取图片文件名（不含路径）作为「来源」标记
function baseNameOf(p: string): string {
  const idx = Math.max(p.lastIndexOf('/'), p.lastIndexOf('\\'))
  return idx >= 0 ? p.slice(idx + 1) : p
}

export interface DetectRecognizeResult {
  error?: boolean
  message?: string
  rows?: AIRecognizedRow[]
  boxes?: DetectedBox[]
  imageWidth?: number
  imageHeight?: number
  /** 是否走了检测裁剪流程（false = 模型不可用，已回退整图识别） */
  detected?: boolean
  modelAvailable?: boolean
  /** 逐张拆分出的小票：含裁剪图（已自动旋转）、边界框、自动旋转角度、各自识别结果 */
  tickets?: RecognizedTicket[]
}

// 单张拆分出来的小票（用于逐个放大查看 / 录入）
export interface RecognizedTicket {
  index: number
  box: DetectedBox
  crop: string // base64 jpeg，已自动旋转为正向
  angle: number // 自动旋转角度（度）
  rows: AIRecognizedRow[]
}

// 检测增强识别：先用 YOLOv8 把每张小票框出来并裁剪，再逐张调用视觉 AI 识别人名与内容，
// 最后合并结果并用人名清单做模糊修正。
// 当模型/Python 不可用时，自动回退到「整图一次性识别」，保证功能不中断。
export async function recognizeTicketsWithDetection(
  imagePath: string,
  config: AIConfig,
  prompt: string,
  nameList: string[],
  opts?: { modelPath?: string; pythonPath?: string; conf?: number; enableDetect?: boolean; imageBase64?: string },
): Promise<DetectRecognizeResult> {
  if (!isValidConfig(config)) {
    return { error: true, message: 'AI 接口未配置：请在设置中填写 Base URL、API Key 和模型。' }
  }

  const useDetect = opts?.enableDetect !== false
  const det = useDetect
    ? await detectTickets({
        imagePath,
        imageBase64: opts?.imageBase64,
        modelPath: opts?.modelPath,
        pythonPath: opts?.pythonPath,
        conf: opts?.conf,
        crops: true,
      })
    : { ok: false, modelAvailable: false, boxes: [], crops: [], imageWidth: 0, imageHeight: 0 }

  // 模型/Python 不可用 → 回退整图识别
  if (!det.modelAvailable) {
    const base = opts?.imageBase64
      ? { error: false, base64: opts.imageBase64, message: '' }
      : await readImageBase64(imagePath)
    if (base.error || !base.base64) {
      return { error: true, message: base.message || '读取图片失败', detected: false, modelAvailable: false }
    }
    const whole = await recognizeReceipt(
      base.base64,
      config,
      buildAugmentedPrompt(prompt, nameList),
    )
    const msg = (whole.message ? whole.message + '\n' : '') +
      '未检测到本地检测模型（YOLOv8），已自动回退为整图识别。可在设置中配置 Python 路径与模型。'
    return { ...whole, message: msg, detected: false, modelAvailable: false }
  }

  if (!det.ok || det.boxes.length === 0) {
    return {
      error: false,
      message: det.message || '未在图中检测到小票区域，请确认图片清晰度或检测模型是否匹配。',
      rows: [],
      boxes: det.boxes,
      imageWidth: det.imageWidth,
      imageHeight: det.imageHeight,
      detected: true,
      modelAvailable: true,
    }
  }

  const source = baseNameOf(imagePath)
  // 逐张裁剪图调用视觉 AI（单票提示词，减少背景干扰、提高识别率）
  const perTicket = await mapLimit(det.crops, 3, async (cropB64: string, idx: number) => {
    const scaled = downscaleImageBase64(cropB64, CROP_MAX_DIM, CROP_QUALITY)
    const res = await recognizeReceipt(
      scaled,
      config,
      buildAugmentedPrompt(SINGLE_TICKET_PROMPT, nameList),
      { maxTokens: MAX_TOKENS_CROP },
    )
    // 备注(remark)按需求留空，不在备注里追加「#序号」；小票序号由 tickets[].index 承载
    const rows = (res.rows || []).map((r) => ({
      ...r,
      source,
    }))
    return { rows, note: res.message }
  })

  const notes: string[] = []
  for (const t of perTicket) {
    if (t?.note) notes.push(t.note)
  }

  // 逐张组织为「小票」结构
  const ticketsRaw: RecognizedTicket[] = det.boxes.map((box, i) => ({
    index: i + 1,
    box,
    crop: det.crops[i] || '',
    angle: box.angle || 0,
    rows: perTicket[i]?.rows || [],
  }))
  // 日期修正：以第一张有日期的小票为参照，统一后面小票的日期
  const tickets = normalizeTicketDates(ticketsRaw)

  // 人名清单模糊修正（基于逐张小票的合并结果）
  const merged = tickets.flatMap((t) => t.rows)
  const { rows, corrected } = correctPersonNames(merged, nameList)

  // 统计被日期修正的条数（与原始对比）
  let dateCorr = 0
  tickets.forEach((t, i) => {
    t.rows.forEach((r, j) => {
      const orig = ticketsRaw[i]?.rows[j]?.date || ''
      if ((orig || '') !== (r.date || '')) dateCorr++
    })
  })

  let message = `检测到 ${det.boxes.length} 张小票并逐张识别，共 ${rows.length} 条记录。`
  if (corrected > 0) message += ` 已按人名清单自动修正 ${corrected} 处人名。`
  if (dateCorr > 0) message += ` 已按首张日期统一 ${dateCorr} 处日期。`
  if (notes.length) message += '\n' + notes.join('\n')

  return {
    rows,
    boxes: det.boxes,
    imageWidth: det.imageWidth,
    imageHeight: det.imageHeight,
    detected: true,
    modelAvailable: true,
    tickets,
    message,
  }
}

// 仅对单张裁剪图（已拆分出来的小票）做 AI 识别，用于「先框出、再逐张识别」的流程。
// 返回该小票识别出的结构化行（已做人名清单模糊修正）。
export async function recognizeSingleCrop(
  cropBase64: string,
  config: AIConfig,
  prompt: string,
  nameList: string[],
): Promise<{ error?: boolean; message?: string; rows?: AIRecognizedRow[] }> {
  if (!isValidConfig(config)) {
    return { error: true, message: 'AI 接口未配置：请在设置中填写 Base URL、API Key 和模型。' }
  }
  const res = await recognizeReceipt(
    downscaleImageBase64(cropBase64, CROP_MAX_DIM, CROP_QUALITY),
    config,
    buildAugmentedPrompt(SINGLE_TICKET_PROMPT, nameList),
    { maxTokens: MAX_TOKENS_CROP },
  )
  if (res.error) return res
  if (res.rows && res.rows.length) {
    const { rows, corrected } = correctPersonNames(res.rows, nameList)
    res.rows = rows
    if (corrected > 0) {
      res.message = (res.message ? res.message + '\n' : '') + `已按人名清单自动修正 ${corrected} 处人名`
    }
  }
  return res
}
