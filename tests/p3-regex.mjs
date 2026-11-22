/**
 * P3 回归（office_find 正则开关）：比对报告 §4.2 第 29 条表格结构读取的替代方案——
 * 「`office_find` 的正则开关」在进度文档里被列为 P3 第三项。
 *
 * 覆盖：
 * 1. 字面量模式未被改动（归一化、命中计数、截图）；
 * 2. regex=true 走原文匹配——大小写敏感、空白敏感、`(?i)` 与 `\s` 必须自己写；
 * 3. 捕获组按命中序回传，未参与的组为 null；
 * 4. 零宽模式（`^`/`\b`/`a*`）不挂死且计数正确；
 * 5. 非法正则、超长 pattern 在扫描任何单元之前就报错；
 * 6. 每单元计数上限 capped 语义与渲染文案；
 * 7. 结果满足宿主边界：lossless JSON + 自己声明的输出 schema。
 *
 * 用法：node tests/p3-regex.mjs（在 dsh-office-reader 目录下）
 */
import { mkdirSync, writeFileSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { loadReaderTools, makeChecker } from './harness.mjs'

const here = dirname(fileURLToPath(import.meta.url))
const editorLib = join(here, '..', '..', 'dsh-office-editor', 'lib', 'docx.js')
const editor = await import(new URL(`file://${editorLib.replace(/\\/gu, '/')}`).href)
const editorXlsx = join(here, '..', '..', 'dsh-office-editor', 'lib', 'xlsx.js')
const xlsx = await import(new URL(`file://${editorXlsx.replace(/\\/gu, '/')}`).href)

const { check, done } = makeChecker('office_find 正则开关')
const { call, render, tools } = await loadReaderTools()

const writeTemp = (name, bytes) => {
  const dir = join(here, 'tmp')
  mkdirSync(dir, { recursive: true })
  const out = join(dir, name)
  writeFileSync(out, bytes)
  return out
}

// 夹具刻意混入：大小写、全角数字、段内换行、行内空格、括号年份。
const docxPath = writeTemp('p3-regex.docx', editor.createDocx([
  { type: 'paragraph', text: 'Alpha 项目于 2024 年立项，Beta 于 2025 年结题。' },
  { type: 'paragraph', text: 'alpha 小写开头，另有一行 ＡＢＣ 全角。' },
  { type: 'paragraph', text: '联系电话 010-1234 5678，备用 010-8765 4321。' },
  { type: 'paragraph', text: '第三段用于测试零宽匹配与捕获组。' },
], {}))

const xlsxPath = writeTemp('p3-regex.xlsx', xlsx.createXlsx([{
  name: '明细',
  rows: [
    ['编号', '名称', '金额'],
    ['A-001', '甲项目', 120],
    ['A-002', '乙项目', 340],
    ['B-001', '丙项目', 560],
  ],
}], 100))

console.log('== 1. 字面量模式保持原语义（归一化 + 空白不敏感） ==')
{
  const result = await call('office_find', { source: docxPath, query: 'alpha' })
  check('默认不走正则', result.regex === false, String(result.regex))
  check('大小写不敏感：两段都命中', result.matches.length === 2, JSON.stringify(result.matches.map((m) => m.unit)))
  check('全半角容错：ABC 命中 ＡＢＣ', (await call('office_find', { source: docxPath, query: 'ABC' })).total_matches >= 1)
  check('空白容错：去掉空格仍命中电话段', (await call('office_find', { source: docxPath, query: '010-12345678' })).total_matches === 1)
  const literal = render('office_find', result, { query: 'alpha' })[0].text
  check('渲染文案标出匹配规则', literal.includes('大小写不敏感'), literal.slice(0, 120))
}

console.log('== 2. regex=true 走原文匹配：不再归一化 ==')
{
  const insensitive = await call('office_find', { source: docxPath, query: 'alpha', regex: true })
  check('正则模式默认忽略大小写（两段都命中）', insensitive.total_matches === 2, JSON.stringify(insensitive.matches.map((m) => [m.unit, m.count])))
  check('回传 case_sensitive=false', insensitive.case_sensitive === false)

  const upper = await call('office_find', { source: docxPath, query: 'Alpha', regex: true, case_sensitive: true })
  check('case_sensitive=true 时区分大小写（只命中 1 处）', upper.total_matches === 1, JSON.stringify(upper.matches.map((m) => [m.unit, m.count])))
  check('回传 case_sensitive=true', upper.case_sensitive === true)
  const lower = await call('office_find', { source: docxPath, query: 'alpha', regex: true, case_sensitive: true })
  check('小写 alpha 命中另一段', lower.matches[0]?.unit === '段落 2', JSON.stringify(lower.matches.map((m) => m.unit)))
  const renderedCase = render('office_find', upper, { query: 'Alpha' })[0].text
  check('渲染文案标出正则与标志', renderedCase.includes('正则搜索：/Alpha/g'), renderedCase.slice(0, 80))

  const inlineFlag = await call('office_find', { source: docxPath, query: '(?i)alpha', regex: true }).then(() => '').catch((error) => error.message)
  check('(?i) 内联标志明确报错（JS 不支持，别让模型猜）', inlineFlag.includes('正则表达式无效'), inlineFlag)

  const fullwidth = await call('office_find', { source: docxPath, query: 'ABC', regex: true })
  check('正则模式不做全半角折叠（ABC 不命中 ＡＢＣ）', fullwidth.total_matches === 0, JSON.stringify(fullwidth.matches))

  const spaced = await call('office_find', { source: docxPath, query: '010-12345678', regex: true })
  check('正则模式不剥空白（连写不命中）', spaced.total_matches === 0, JSON.stringify(spaced.matches))
  const pattern = await call('office_find', { source: docxPath, query: '010-\\d{4}\\s?\\d{4}', regex: true })
  check('用 \\s? 显式表达空白时命中两处', pattern.total_matches === 2, JSON.stringify(pattern.matches.map((m) => m.count)))
  check('正则命中计数正确', pattern.matches[0]?.count === 2, JSON.stringify(pattern.matches[0]))
}

console.log('== 3. 字符类与锚点（字面量模式表达不了的能力） ==')
{
  const years = await call('office_find', { source: docxPath, query: '(20\\d{2}) 年', regex: true })
  check('模式命中两年份', years.total_matches === 2, JSON.stringify(years.matches.map((m) => m.count)))
  // xlsx 的命中单元是「一整行」（Sheet!第N行，单元格以制表符分隔），所以 ^ 锚的是行首而不是单元格。
  const rowAnchor = await call('office_find', { source: xlsxPath, query: '^A-\\d{3}\\t', regex: true, sheet: '明细' })
  check('锚点按行单元文本匹配（^ 锚行首）', rowAnchor.total_matches === 2, JSON.stringify(rowAnchor.matches.map((m) => [m.unit, m.count])))
  check('命中单元标为数据行', rowAnchor.matches[0]?.unit === '明细!第2行' || rowAnchor.matches[0]?.unit.startsWith('明细!第'), JSON.stringify(rowAnchor.matches.map((m) => m.unit)))
  const notB = await call('office_find', { source: xlsxPath, query: '^(?!B-)\\w+-\\d{3}\\t', regex: true, sheet: '明细' })
  check('否定前瞻可用（排除 B 开头）', notB.total_matches === 2, JSON.stringify(notB.matches.map((m) => m.count)))
  const range = await call('office_find', { source: xlsxPath, query: '[0-9]{3}', regex: true, sheet: '明细' })
  check('数字类命中编号与金额', range.total_matches >= 3, JSON.stringify(range.matches.map((m) => m.count)))
}

console.log('== 4. 捕获组按命中序回传，未参与的组为 null ==')
{
  const captured = await call('office_find', { source: docxPath, query: '(\\d{4}) 年(\\d+)?', regex: true })
  check('captured 存在', Array.isArray(captured.matches[0]?.captured), JSON.stringify(captured.matches[0]))
  check('第一组是年份', captured.matches[0].captured[0][0] === '2024', JSON.stringify(captured.matches[0].captured))
  check('未参与的可选组为 null', captured.matches[0].captured[0][1] === null, JSON.stringify(captured.matches[0].captured))
  check('两组命中按序排列', captured.matches[0].captured.length === 2 && captured.matches[0].captured[1][0] === '2025', JSON.stringify(captured.matches[0].captured))
  const rendered = render('office_find', captured, { query: 'x' })[0].text
  check('渲染文案列出捕获组', rendered.includes('捕获组'), rendered.slice(-160))

  const noGroups = await call('office_find', { source: docxPath, query: 'alpha', regex: true })
  check('无分组时不产生 captured 键', noGroups.matches.every((m) => m.captured === undefined), JSON.stringify(noGroups.matches.map((m) => Object.keys(m))))
  const noGroupRender = render('office_find', noGroups, { query: 'alpha' })[0].text
  check('无分组时渲染不出现捕获组行', !noGroupRender.includes('捕获组'))
}

console.log('== 5. 零宽模式不挂死且计数正确 ==')
{
  const anchors = await call('office_find', { source: docxPath, query: '^', regex: true })
  check('^ 每个非空单元记 1 次', anchors.total_matches === 4, JSON.stringify(anchors.matches.map((m) => [m.unit, m.count])))
  check('零宽命中不产生片段', anchors.matches.every((m) => m.snippets.length === 0))
  const stars = await call('office_find', { source: docxPath, query: 'a*', regex: true })
  check('a* 正常返回（不无限循环）', stars.total_matches > 0, String(stars.total_matches))
  const word = await call('office_find', { source: docxPath, query: '\\b[Aa]lpha\\b', regex: true })
  check('\\b 词界命中「Alpha」与「alpha」两处', word.total_matches === 2, JSON.stringify(word.matches.map((m) => m.count)))
  const wordExact = await call('office_find', { source: docxPath, query: '\\bAlpha\\b', regex: true, case_sensitive: true })
  check('区分大小写时 \\bAlpha\\b 只命中 1 处', wordExact.total_matches === 1, JSON.stringify(wordExact.matches.map((m) => m.count)))
}

console.log('== 6. 非法与超长 pattern 在扫描前报错 ==')
{
  let invalid = ''
  try { await call('office_find', { source: docxPath, query: '(未闭合', regex: true }) } catch (error) { invalid = error.message }
  check('非法正则报错', invalid.includes('正则表达式无效'), invalid)

  let tooLong = ''
  try { await call('office_find', { source: docxPath, query: 'a'.repeat(600), regex: true }) } catch (error) { tooLong = error.message }
  check('超长 pattern 报错', tooLong.includes('过长'), tooLong)

  let empty = ''
  try { await call('office_find', { source: docxPath, query: '  ' }) } catch (error) { empty = error.message }
  check('字面量空白 query 仍报错', empty.includes('query'), empty)
}

console.log('== 7. 每单元计数上限与 truncate 语义 ==')
{
  const many = writeTemp('p3-regex-many.docx', editor.createDocx([
    { type: 'paragraph', text: 'x'.repeat(300) },
    { type: 'paragraph', text: '普通段落' },
  ], {}))
  // 300 个 x < 默认上限 1000，所以先确认「计数不被上限截断」（返回条数仍受 max_matches=50 约束）。
  const underCap = await call('office_find', { source: many, query: 'x', regex: true })
  check('计数完整（300）而不被返回上限截断', underCap.matches[0].count === 300 && underCap.total_matches === 300, JSON.stringify({ n: underCap.matches[0].count, total: underCap.total_matches }))
  check('片段只给 5 条、明细标截断', underCap.matches[0].snippets.length === 5 && underCap.truncated === true, JSON.stringify({ s: underCap.matches[0].snippets.length, t: underCap.truncated }))

  // ……再用直接调引擎的方式钉住「撞上限即停」，不依赖 1000 这个默认值。
  const { searchUnitRegex, compilePattern } = await import('../lib/search.js')
  const capped = searchUnitRegex('x'.repeat(300), compilePattern('x'), 10, 5, 7)
  check('计数停在上限 7', capped.count === 7, String(capped.count))
  check('撞上限时 capped=true', capped.capped === true)
  check('片段仍受 maxSnippets 限制', capped.snippets.length === 5, String(capped.snippets.length))

  const uncapped = searchUnitRegex('x'.repeat(300), compilePattern('x'), 10, 5, 1000)
  check('上限足够时计数完整且 capped=false', uncapped.count === 300 && uncapped.capped === false, `${uncapped.count}/${uncapped.capped}`)

  const zeroWidth = searchUnitRegex('abc', compilePattern(''), 1, 5, 100)
  check('空模式在每个位置命中且不挂死', zeroWidth.count === 4 && zeroWidth.snippets.length === 0, JSON.stringify(zeroWidth))
}

console.log('== 8. 宿主边界：lossless JSON 与输出 schema ==')
{
  const value = await call('office_find', { source: docxPath, query: '(\\d{4})', regex: true })
  let lossless = true
  try { JSON.stringify(value) } catch { lossless = false }
  check('结果是 lossless JSON', lossless && !JSON.stringify(value).includes('undefined'))
  check('regex 字段为 true', value.regex === true)
  check('matches 项只含声明过的键', value.matches.every((m) => Object.keys(m).every((key) => ['unit', 'count', 'snippets', 'captured'].includes(key))), JSON.stringify(Object.keys(value.matches[0])))

  // 按宿主同一套子集规则核对输出 schema（additionalProperties:false + required + items）
  const schema = tools.get('office_find').output.schema
  const violations = []
  const visit = (node, val, path) => {
    if (node.type === 'object') {
      for (const key of node.required ?? []) if (!Object.hasOwn(val, key)) violations.push(`${path}.${key} 缺失`)
      if (node.additionalProperties === false) {
        for (const key of Object.keys(val)) if (!Object.hasOwn(node.properties ?? {}, key)) violations.push(`${path}.${key} 未声明`)
      }
      for (const [key, child] of Object.entries(node.properties ?? {})) {
        if (Object.hasOwn(val, key)) visit(child, val[key], `${path}.${key}`)
      }
      return
    }
    if (node.type === 'array') {
      if (!Array.isArray(val)) { violations.push(`${path} 应为数组`); return }
      for (const [index, item] of val.entries()) visit(node.items, item, `${path}[${index}]`)
      return
    }
    if (node.type === 'integer' && !Number.isInteger(val)) violations.push(`${path} 应为整数`)
    if (node.type === 'string' && typeof val !== 'string') violations.push(`${path} 应为字符串`)
    if (node.type === 'boolean' && typeof val !== 'boolean') violations.push(`${path} 应为布尔`)
  }
  visit(schema, value, '$')
  check('结果满足自己声明的输出 schema', violations.length === 0, violations.join('；'))
}

process.exit(done())
