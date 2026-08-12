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

function isValidConfig(c?: AIConfig): boolean {
  return !!c && typeof c.baseURL === 'string' && c.baseURL.trim() !== '' &&
    typeof c.apiKey === 'string' && c.apiKey.trim() !== '' &&
    typeof c.model === 'string' && c.model.trim() !== ''
}

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
    }))
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
  const body = {
    model: config.model,
    temperature: typeof config.temperature === 'number' ? config.temperature : 0.2,
    max_tokens: 4096,
    messages: [
      {
        role: 'system',
        content: '你是一位严谨的中文票据识别助手，只输出合法 JSON。',
      },
      {
        role: 'user',
        content: [
          { type: 'text', text: prompt || DEFAULT_PROMPT },
          {
            type: 'image_url',
            image_url: { url: `data:image/jpeg;base64,${imageBase64}` },
          },
        ],
      },
    ],
  }

  let responseText = ''
  try {
    const res = (await getFetch()(url, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        Authorization: `Bearer ${config.apiKey}`,
      },
      body: JSON.stringify(body),
    })) as { ok: boolean; status: number; statusText: string; text: () => Promise<string> }
    responseText = await res.text()
    if (!res.ok) {
      return { error: true, message: `AI 接口错误 ${res.status}：${responseText.slice(0, 300)}` }
    }
  } catch (err) {
    return { error: true, message: '请求 AI 接口失败：' + (err instanceof Error ? err.message : '未知错误') }
  }

  let responseJson: unknown
  try {
    responseJson = JSON.parse(responseText)
  } catch {
    return { error: true, message: 'AI 接口返回非 JSON：' + responseText.slice(0, 300) }
  }

  const data = responseJson as Record<string, unknown>
  const choices = data.choices as Array<Record<string, unknown>> | undefined
  const first = choices && choices[0]
  const message = first?.message as Record<string, unknown> | undefined
  const content = message?.content
  if (typeof content !== 'string') {
    return { error: true, message: 'AI 接口返回内容为空或格式异常' }
  }

  let parsed: unknown
  try {
    parsed = extractJson(content)
  } catch (err) {
    return {
      error: true,
      message: 'AI 返回无法解析为 JSON：' + (err instanceof Error ? err.message : '') + '\n原始内容：' + content.slice(0, 400),
    }
  }

  const rows = normalizeRows(parsed)
  if (rows.length === 0) {
    return { error: true, message: 'AI 未识别出任何小票记录，请检查图片清晰度或提示词。' }
  }
  return { rows }
}
