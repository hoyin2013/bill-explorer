// AI 小票识别服务：调用 OpenAI 兼容的 chat/completions 接口，
// 把图片里的手写/印刷小票内容识别成结构化账单字段。

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
}

export const DEFAULT_PROMPT = `你是一位票据录入助手，专门识别手写销售小票。请识别图片中的全部小票并按条提取为 JSON。

【小票版式说明】（手写销售小票）
- 每张小票顶部有"销售小票"字样。
- 人名（手写）：位于"销售小票"字样的右上方，是手写的人名，对应账单中的"调货人/人名"。
- 年月日：位于"销售小票"字样的紧邻正下方，即日期。
- 小票主体列有：货品名称、单位、数量、单价、金额。

【输出字段】（对应账单 9 列，顺序一致）
{
  "receipts": [
    {
      "no": "序号，可留空（程序会自动编号）",
      "date": "日期，取自'销售小票'正下方，输出 yyyy-mm-dd，缺年则补当前年",
      "name": "货品名称（小票主体上的货品，注意：不是人名！）",
      "unit": "单位（如件、个、斤、套、箱）",
      "qty": "数量",
      "price": "单价",
      "amount": "金额（总价）",
      "person": "人名/调货人，取自'销售小票'右上方手写的名字",
      "remark": "备注"
    }
  ]
}

【识别规则】
1. 一张图片包含多张并排的手写小票，请逐张识别，每张小票输出一条记录。
2. 若单张小票列了多种货品，则每种货品各输出一条记录，并共享同一 person 与 date。
3. 人名务必从"销售小票"字样的右上方手写名字读取，严禁把货品名称当作人名。
4. 日期务必从"销售小票"紧邻正下方读取；若只写月日则补当前年。
5. 数量、单价、金额只保留数字，不要带"件/元/￥"等单位。
6. 若小票只有金额没有数量/单价，则金额必填，数量与单价可留空。
7. 手写可能潦草，请结合上下文（货品、单位、金额）合理推断，不确定字段留空。
8. 不要输出任何解释文字，只输出 JSON。`

// 单次请求超时（毫秒）。小票图识别通常较慢，给足 120s。
const AI_TIMEOUT_MS = 120_000
// 输出 token 上限，单图小票较多时避免被截断。
const MAX_TOKENS = 8192
// 网络抖动 / 5xx / 429 的退避重试次数（不含首次）。
const RETRY_COUNT = 2

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

function normalizeRows(parsed: unknown): AIRecognizedRow[] {
  if (!parsed || typeof parsed !== 'object') return []
  const data = parsed as Record<string, unknown>
  const raw = Array.isArray(data.receipts) ? data.receipts : Array.isArray(data.rows) ? data.rows : Array.isArray(data) ? data : []
  return raw
    .filter((r): r is Record<string, unknown> => r && typeof r === 'object')
    .map((r) => ({
      no: r.no != null ? String(r.no) : '',
      date: r.date != null ? String(r.date) : '',
      name: r.name != null ? String(r.name) : '',
      unit: r.unit != null ? String(r.unit) : '',
      qty: r.qty != null ? String(r.qty) : '',
      price: r.price != null ? String(r.price) : '',
      amount: r.amount != null ? String(r.amount) : '',
      person: r.person != null ? String(r.person) : '',
      remark: r.remark != null ? String(r.remark) : '',
      source: r.source != null ? String(r.source) : '',
    }))
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
  let lastError = ''

  for (let v = 0; v < imageContentVariants.length; v++) {
    const variant = imageContentVariants[v]
    // 先尝试带 JSON 模式（强制 JSON 输出），被服务端拒绝则退回纯 prompt。
    const jsonModes: boolean[] = [true, false]
    for (const useJson of jsonModes) {
      const body: Record<string, unknown> = {
        model: config.model,
        temperature,
        max_tokens: MAX_TOKENS,
        messages: [
          {
            role: 'system',
            content: '你是一位严谨的中文票据识别助手，只输出合法 JSON。',
          },
          {
            role: 'user',
            content: [{ type: 'text', text: prompt || DEFAULT_PROMPT }, variant],
          },
        ],
      }
      if (useJson) {
        body.response_format = { type: 'json_object' }
      }

      const result = await callWithRetry((signal) => doRequest(url, config, body, signal))

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
      return result.note
        ? { rows: result.rows, message: result.note }
        : { rows: result.rows }
    }
  }

  // 所有格式都被服务端拒绝
  return {
    error: true,
    message: `AI 接口不支持图片输入（尝试了 ${imageContentVariants.length} 种格式均被拒绝）。\n${lastError}\n请检查：① 使用支持 vision 的模型；② 确认 Base URL 为正确的 OpenAI 兼容视觉接口。`,
  }
}
