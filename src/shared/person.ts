// 共享：人名解析与相似度匹配（主进程 / 渲染进程共用）

// 常见中文姓氏（用于判断一段中文是否为"人名"而非地名/行业词）
const SURNAMES = new Set(
  '王李张刘陈杨黄赵周吴徐孙朱马胡郭何高林罗郑梁谢宋唐许韩冯邓曹彭曾肖田董袁潘于蒋蔡余杜叶程苏魏吕丁任沈姚卢傅钟姜崔谭廖范汪陆金石贾夏邱方侯邹熊孟秦白江阎薛尹段雷黎史龙陶贺顾毛郝龚邵万钱严覃武戚莫孔向常石骆樊兰殷'.split(
    '',
  ),
)

// 明显不是人名的词（地名 / 行业 / 机构 / 后缀等），命中即剔除
const NON_NAME_WORDS = [
  '北京', '上海', '广州', '深圳', '天津', '重庆', '杭州', '南京', '成都', '武汉', '西安', '苏州', '郑州', '长沙', '青岛', '宁波', '东莞', '佛山', '无锡',
  '电脑', '维修', '公司', '有限', '责任', '股份', '集团', '店', '超市', '市场', '批发', '商城', '中心', '科技', '贸易', '工厂', '商行', '经营', '门诊', '医院',
  '学校', '银行', '酒店', '宾馆', '餐厅', '饭店', '美食', '美容', '理发', '五金', '建材', '食品', '电子', '通信', '物流', '机械', '化工', '服装', '鞋业', '百货',
  '网络', '软件', '装饰', '安装', '加工', '打印', '复印', '批发部', '经营部', '工作室', '有限公司',
]

// 是否像"英文/拼音人名"（纯字母，可含空格/点/连字符/撇号，不含数字）
function isLatinName(s: string): boolean {
  return /^[a-zA-Z][a-zA-Z.\s'\-]*$/.test(s)
}

// 一段文本是否像"人名"：中文按姓氏/昵称前缀/长度判断；拉丁按英文名判断；其余（地名/行业/符号）剔除
function looksLikePersonName(s: string): boolean {
  if (!s) return false
  if (/^[\d\s]+$/.test(s)) return false // 纯数字 / 空白
  for (const w of NON_NAME_WORDS) if (s.includes(w)) return false // 地名 / 行业词
  if (isLatinName(s)) return true // 英文 / 拼音人名
  if (!/[一-龥]/.test(s)) return false // 其余非中文也非拉丁 → 不是人名
  if (SURNAMES.has(s[0])) return true // 首字为常见姓氏（张三 / 范文华）
  if (['小', '老', '阿', '大'].includes(s[0]) && s.length >= 2) return true // 小袁 / 老李 / 阿强
  if (s.length >= 2 && s.length <= 4) return true // 兜底：2~4 字中文且非黑名单
  return false
}

// 从 Excel 文件名里解析"人名"：
// - 过滤 Excel 临时锁文件（~$ 开头）
// - 去掉末尾的 年月日 / 编号 / 单据字样
// - 按 空格/下划线/连字符 切词，取第一个"像人名"的词（支持「北京 小袁」这类前缀+人名、英文人名）
// 例：「张三.xlsx」→ 张三；「张三_2026-08-12.xlsx」→ 张三；「北京 小袁.xlsx」→ 小袁；
//     「801 A6 范文华.xlsx」→ 范文华；「Tom Smith.xlsx」→ Tom Smith；「~$801 A6 范文华.xlsx」→ ''
export function parsePersonFromFilename(filePath: string): string {
  const base = (filePath.split(/[\\/]/).pop() || filePath).replace(/\.[^.]+$/, '')
  // Excel 临时锁文件（~$ 开头）：直接忽略，不应进入人名清单
  if (/^~\$/u.test(base)) return ''
  let name = base
  // 去掉末尾的 年月日 / 编号 / 单据字样 等修饰词
  name = name.replace(/(?:[_\- ])+\d{4}\D.*$/u, '') // 2026、2026年、2026-...
  name = name.replace(/(?:[_\- ])+\d{1,2}[_\-]\d{1,2}(?:[_\-]\d{1,2})?.*$/u, '') // 8-11 / 8-11-1
  name = name.replace(/(?:[_\- ])+\d+$/u, '') // 末尾纯数字编号
  name = name.replace(/(?:[_\- ])*(?:账单|小票|明细|记录|汇总|统计|单|表).*$/iu, '') // 末尾单据字样
  name = name.replace(/[_\- ]+/g, ' ').trim()
  if (!name) return ''
  // 按 空格 / 下划线 / 连字符 切词，逐个找第一个"像人名"的词
  const tokens = name.split(/[\s_\-]+/).filter(Boolean)
  for (let i = 0; i < tokens.length; i++) {
    const t = tokens[i]
    if (isLatinName(t)) {
      // 英文 / 拼音人名：合并后续连续的拉丁词（Tom Smith / Fan Hua）
      let full = t
      let j = i + 1
      while (j < tokens.length && isLatinName(tokens[j])) {
        full += ' ' + tokens[j]
        j++
      }
      return full
    }
    if (looksLikePersonName(t)) return t
  }
  return ''
}

// 归一化人名：去空白 + 转小写，便于比较
export function normName(s: string): string {
  return String(s || '').replace(/\s+/g, '').toLowerCase()
}

// 文件名人名 与 识别人名 是否匹配（双向包含，容错大小写/空格）
export function personMatches(filePerson: string, recPerson: string): boolean {
  const a = normName(filePerson)
  const b = normName(recPerson)
  if (!a || !b) return false
  return a.includes(b) || b.includes(a)
}

// 字符级 Levenshtein 相似度（0~1）：用于把识别到的人名映射到清单中最接近的一项
export function nameSimilarity(a: string, b: string): number {
  const s = (a || '').trim()
  const t = (b || '').trim()
  if (!s && !t) return 1
  if (!s || !t) return 0
  const m = s.length
  const n = t.length
  // 一维滚动数组，省内存
  const dp: number[] = new Array(n + 1)
  for (let j = 0; j <= n; j++) dp[j] = j
  for (let i = 1; i <= m; i++) {
    let prev = dp[0]
    dp[0] = i
    for (let j = 1; j <= n; j++) {
      const tmp = dp[j]
      dp[j] = s[i - 1] === t[j - 1] ? prev : 1 + Math.min(dp[j], dp[j - 1], prev)
      prev = tmp
    }
  }
  const dist = dp[n]
  return 1 - dist / Math.max(m, n)
}

// 在已知人名清单中找到与输入最相似的一项；低于阈值返回 null（视为清单外的新客户）
export function closestName(input: string, list: string[], threshold = 0.5): string | null {
  const a = (input || '').trim()
  if (!a || !list || list.length === 0) return null
  let best: string | null = null
  let bestScore = 0
  for (const cand of list) {
    const c = (cand || '').trim()
    if (!c) continue
    const score = nameSimilarity(a, c)
    if (score > bestScore) {
      bestScore = score
      best = c
    }
  }
  return bestScore >= threshold ? best : null
}

// 在清单（完整文件名，如 "DB210 袁文志"）中为识别出的名字（如 "袁文志"）找最接近的匹配。
// 综合三种判断：① 归一化相等；② 双向包含（识别名是清单项的一部分，或反之）；
// ③ 字符级相似度（分别与清单项整体、清单项提取出的人名部分比较，取高者）。
// 识别名不必与清单项字面完全一致——例如 "袁文志" 应匹配到 "DB210 袁文志"。
// 返回匹配结果与得分（0~1），低于阈值视为清单外的新客户。
export function fuzzyMatchName(
  input: string,
  list: string[],
  threshold = 0.34,
): { matched: string | null; score: number } {
  const a = normName(input)
  if (!a || !list || list.length === 0) return { matched: null, score: 0 }
  let best: string | null = null
  let bestScore = 0
  for (const cand of list) {
    const c = normName(cand)
    if (!c) continue
    let score = 0
    if (a === c) score = 1
    else if (a.includes(c) || c.includes(a)) score = 0.92
    else {
      const s1 = nameSimilarity(a, c)
      // 清单项可能带前缀编号/地点（"DB210 袁文志"），提取其中的人名部分再比较
      const extracted = normName(parsePersonFromFilename(cand))
      const s2 = extracted ? nameSimilarity(a, extracted) : 0
      score = Math.max(s1, s2)
    }
    if (score > bestScore) {
      bestScore = score
      best = cand
    }
  }
  return { matched: bestScore >= threshold ? best : null, score: bestScore }
}
