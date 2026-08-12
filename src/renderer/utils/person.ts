// 从 Excel 文件名里解析"人名"：如「张三.xlsx」→ 张三；「张三_2026-08-12.xlsx」→ 张三
export function parsePersonFromFilename(filePath: string): string {
  const base = (filePath.split(/[\\/]/).pop() || filePath).replace(/\.[^.]+$/, '')
  let name = base
  // 去掉末尾的 年月日 / 编号 / 单据字样 等修饰词，只保留人名部分
  name = name.replace(/(?:[_\- ])+\d{4}\D.*$/u, '') // 2026、2026年、2026-...
  name = name.replace(/(?:[_\- ])+\d{1,2}[_\-]\d{1,2}(?:[_\-]\d{1,2})?.*$/u, '') // 8-11 / 8-11-1
  name = name.replace(/(?:[_\- ])+\d+$/u, '') // 末尾纯数字编号
  name = name.replace(/(?:[_\- ])*(?:账单|小票|明细|记录|汇总|统计|单|表).*$/iu, '') // 末尾单据字样
  name = name.replace(/[_\- ]+/g, ' ').trim()
  return name
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
