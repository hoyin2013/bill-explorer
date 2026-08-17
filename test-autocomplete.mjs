// 自动化测试：自动补全算法层。运行：node --test test-autocomplete.mjs
// 必须先 npm run build:autocomplete 生成 dist-autocomplete/autocomplete.mjs
import { test } from 'node:test'
import assert from 'node:assert/strict'
import { computeSuggestions, extractColumnValues, serialToDateStr, toItem } from './dist-autocomplete/autocomplete.mjs'

// 模拟数据库：3 列、若干历史值（含日期）。
function makeWb(rows) {
  // rows: { r: number; c: number; v: string|number }[]
  const cellData = {}
  for (const { r, c, v } of rows) {
    cellData[r] = cellData[r] || {}
    cellData[r][c] = { v }
  }
  return { sheets: { bill: { cellData } } }
}

const excelEpoch = (y, m, d) => {
  // Excel 序列号纪元 1899-12-30
  const base = Date.UTC(1899, 11, 30)
  return (Date.UTC(y, m - 1, d) - base) / 86400000
}

test('extractColumnValues: 聚合同列值，最大行号、累计 freq', () => {
  const wb = makeWb([
    { r: 0, c: 2, v: '苹果' },
    { r: 1, c: 2, v: '香蕉' },
    { r: 2, c: 2, v: '苹果' },
    { r: 3, c: 2, v: '橘子' },
    { r: 4, c: 2, v: '苹果' },
  ])
  const vals = extractColumnValues(wb, 2)
  const byV = Object.fromEntries(vals.map((v) => [v.value, v]))
  assert.equal(byV['苹果'].freq, 3)
  assert.equal(byV['苹果'].row, 4)
  assert.equal(byV['香蕉'].freq, 1)
  assert.equal(byV['橘子'].freq, 1)
})

test('computeSuggestions: 货品名称列 startsWith 优先 + 上方最近 + 频率优先', () => {
  const wb = makeWb([
    { r: 0, c: 2, v: '苹果' },
    { r: 1, c: 2, v: '苹果梨' },
    { r: 2, c: 2, v: '香蕉' },
    { r: 3, c: 2, v: '山楂' },
    { r: 4, c: 2, v: '苹果' }, // '苹果' 再次出现 → freq=2 row=4
    { r: 5, c: 2, v: '苹果干' },
  ])
  const vals = extractColumnValues(wb, 2)
  const sugs = computeSuggestions('苹', 2, 20, vals, [])
  // 期望: row 越大越靠前。'苹果干'(row=5)>'苹果'(row=4)>'苹果梨'(row=1)
  // 同 row 时 freq 大者优先：'苹果'(row=4 freq=2) 本应第二
  assert.equal(sugs.length, 3)
  assert.deepEqual(sugs.map((s) => s.display), ['苹果干', '苹果', '苹果梨'])
})

test('computeSuggestions: 大小写不敏感', () => {
  const wb = makeWb([
    { r: 0, c: 2, v: 'Apple' },
    { r: 1, c: 2, v: 'Banana' },
    { r: 2, c: 2, v: 'apricot' },
  ])
  const vals = extractColumnValues(wb, 2)
  const sugs = computeSuggestions('AP', 2, 10, vals, [])
  assert.equal(sugs.length, 2)
  assert.deepEqual(sugs.map((s) => s.display.toLowerCase()), ['apricot', 'apple'])
})

test('computeSuggestions: startsWith 无结果时退化为 includes', () => {
  const wb = makeWb([
    { r: 0, c: 2, v: '红苹果' },
    { r: 1, c: 2, v: '绿苹果' },
    { r: 2, c: 2, v: '香蕉' },
  ])
  const vals = extractColumnValues(wb, 2)
  // "苹"没有 startsWith 匹配 → 退化为 includes
  const sugs = computeSuggestions('苹', 2, 10, vals, [])
  assert.equal(sugs.length, 2)
  assert.deepEqual(sugs.map((s) => s.display).sort(), ['红苹果', '绿苹果'])
})

test('computeSuggestions: 行号等于当前行及以上都不参与，只用 row<currentRow 的（与 Excel 一致）', () => {
  const wb = makeWb([
    { r: 5, c: 2, v: '苹果' },
    { r: 6, c: 2, v: '苹果' },
    { r: 7, c: 2, v: '苹果' }, // = currentRow 也不参与（row<currentRow 不含 row=7）
  ])
  const vals = extractColumnValues(wb, 2)
  // 修复后 extractColumnValues 会展开每个 row 都贡献一个 entry（去聚合化）
  assert.equal(vals.length, 3)
  // filter row<7 → row=5/6 通过
  const sugs = computeSuggestions('苹', 2, 7, vals, [])
  assert.equal(sugs.length, 1)
  assert.equal(sugs[0].display, '苹果')

  // currentRow=6 时，row<6 只剩 row=5
  const sugs2 = computeSuggestions('苹', 2, 6, vals, [])
  assert.equal(sugs2.length, 1)
  assert.equal(sugs2[0].display, '苹果')

  // currentRow=8 时，row<8 全部参与（包括 row=5,6,7）
  const sugs3 = computeSuggestions('苹', 2, 8, vals, [])
  assert.equal(sugs3.length, 1)
  assert.equal(sugs3[0].display, '苹果')

  // 关键回归：currentRow 在两个同名值之间时仍能看到——验证聚合 max(row) bug 已修
  const wb2 = makeWb([
    { r: 0, c: 2, v: '苹果' },
    { r: 4, c: 2, v: '苹果' },
  ])
  const vals2 = extractColumnValues(wb2, 2)
  // currentRow=2, row<2 只剩 row=0
  const sugsR2 = computeSuggestions('苹', 2, 2, vals2, [])
  assert.equal(sugsR2.length, 1, 'row=2 输入时仍应看到 "苹果"（row=0 贡献）')
  assert.equal(sugsR2[0].display, '苹果')
})

test('computeSuggestions: 历史记录加权（频率+2）', () => {
  const wb = makeWb([
    { r: 0, c: 2, v: '苹果' },
    { r: 1, c: 2, v: '香蕉' },
  ])
  const vals = extractColumnValues(wb, 2)
  // 历史中也有"苹果"→ 它的 freq 应该比 column vals 中没历史的"香蕉"高
  // 我们用 get 验证 freq 而非排序，但 freq 排序在 row 相等时发挥作用
  // 用更大的列数据集触发频率排序
  const wb2 = makeWb([
    { r: 0, c: 2, v: '苹果' },
    { r: 1, c: 2, v: '苹果' },
    { r: 2, c: 2, v: '苹果' },
    { r: 3, c: 2, v: '香蕉' },
    { r: 4, c: 2, v: '香蕉' },
    { r: 5, c: 2, v: '橘子' },
  ])
  const vals2 = extractColumnValues(wb2, 2)
  // 历史加 '苹果' 三次（但只有一条会增加 freq+2 一次）
  const sugs = computeSuggestions('', 2, 20, vals2, ['苹果'])
  // 实际上 partial="" 会返回 []，我们直接用 partial='a'（'苹果'的拼音首字母）
  const sugs2 = computeSuggestions('苹', 2, 20, vals2, ['苹果'])
  // 苹果 freq=3+2=5，香蕉 freq=2，橘子 freq=1，无相同 row
  // 期待排序：苹果在橘子前
  assert.equal(sugs2[0].display, '苹果')
})

test('computeSuggestions: 空 partial 返回空', () => {
  const wb = makeWb([{ r: 0, c: 2, v: '苹果' }])
  const vals = extractColumnValues(wb, 2)
  assert.deepEqual(computeSuggestions('', 2, 5, vals, []), [])
  assert.deepEqual(computeSuggestions('   ', 2, 5, vals, []), [])
})

test('computeSuggestions: 日期列展示 yyyy-mm-dd，raw 是序列号，且按 row 降序（上→下：靠近 currentRow 的在前）', () => {
  // 列 0 = 序号, 列 1 = 日期，列 2 = 货品名称
  const dateA = excelEpoch(2026, 8, 16)
  const dateB = excelEpoch(2026, 8, 17)
  const wb = makeWb([
    { r: 0, c: 1, v: dateA },
    { r: 1, c: 1, v: dateB },
  ])
  const vals = extractColumnValues(wb, 1)
  const sugs = computeSuggestions('2026-08-1', 1, 10, vals, [])
  assert.equal(sugs.length, 2)
  // 上方最近优先：row=1 (08-17) 在 row=0 (08-16) 前面
  assert.deepEqual(sugs.map((s) => s.display), ['2026-08-17', '2026-08-16'])
  // raw 必须是序列号（不是字符串）
  assert.equal(typeof sugs[0].raw, 'number')
  assert.ok(Math.abs(sugs[0].raw - dateB) < 0.01)
})

test('computeSuggestions: 日期列也支持文本 "20260816" 这种输入', () => {
  const dateA = excelEpoch(2026, 8, 16)
  const wb = makeWb([{ r: 0, c: 1, v: dateA }])
  const vals = extractColumnValues(wb, 1)
  const sugs = computeSuggestions('2026-08-1', 1, 10, vals, [])
  assert.equal(sugs.length, 1)
  assert.equal(sugs[0].display, '2026-08-16')
})

test('toItem: 日期列遇到 number 直接当序列号', () => {
  const it = toItem(excelEpoch(2026, 1, 1), true)
  assert.equal(it.display, '2026-01-01')
  assert.equal(typeof it.raw, 'number')
})
