/**
 * 对比基准：Office 工具读取 vs 直接读等内容 MD 文件。
 * 测量两类指标：① 解析耗时 ② 输出字符量（token 与字符近似成正比）。
 * 用法：node tests/bench.mjs
 */
import { readFile, writeFile } from 'node:fs/promises'
import { performance } from 'node:perf_hooks'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { zipSync, strToU8 } from 'fflate'
import XLSX from 'xlsx'
import { readDocx } from '../lib/docx.js'
import { readXlsx } from '../lib/xlsx.js'

const root = dirname(dirname(fileURLToPath(import.meta.url)))
const fixtures = join(root, 'tests', 'fixtures')
const median = (arr) => arr.sort((a, b) => a - b)[Math.floor(arr.length / 2)]

// ---------- 生成大样例 ----------
const PARA = 2000
const ROWS = 20000
const docParas = []
let docXmlBody = ''
for (let i = 1; i <= PARA; i++) {
  const heading = i % 20 === 1
  const text = heading ? `第 ${Math.ceil(i / 20)} 章 标题` : `段落 ${i}：这是一段大约五十个字符的正文内容，用于测量 docx 段落解析与重建的耗时和输出体积。`
  docParas.push({ heading, text, chars: text.length })
  docXmlBody += heading
    ? `<w:p><w:pPr><w:pStyle w:val="Heading1"/><w:outlineLvl w:val="0"/></w:pPr><w:r><w:t>${text}</w:t></w:r></w:p>`
    : `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`
}
await writeFile(join(fixtures, 'big.docx'), zipSync({
  'word/document.xml': strToU8(`<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${docXmlBody}</w:body></w:document>`),
}))
console.log(`big.docx 生成：${PARA} 段`)

const header = ['日期', '产品', '渠道', '金额', '备注']
const aoa = [header]
for (let i = 0; i < ROWS; i++) {
  aoa.push([`2026-01-${String((i % 28) + 1).padStart(2, '0')}`, `产品-${i % 50}`, ['线上', '线下'][i % 2], 100 + (i % 900), i % 100 === 0 ? `重点行 ${i}` : ''])
}
const wb = XLSX.utils.book_new()
XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(aoa), '数据')
await writeFile(join(fixtures, 'big.xlsx'), XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' }))
console.log(`big.xlsx 生成：${ROWS} 行 × 5 列`)

// 等内容 MD（docx 内容 → Markdown；xlsx 内容 → Markdown 表格）
const mdDocx = docParas.map((p) => (p.heading ? `# ${p.text}` : p.text)).join('\n\n')
await writeFile(join(fixtures, 'big-docx.md'), mdDocx, 'utf8')
const mdXlsx = [`| ${header.join(' | ')} |`, `| ${header.map(() => '---').join(' | ')} |`,
  ...aoa.slice(1).map((r) => `| ${r.join(' | ')} |`)].join('\n')
await writeFile(join(fixtures, 'big-xlsx.md'), mdXlsx, 'utf8')
console.log('等内容 MD 已生成')

// ---------- 耗时 ----------
async function timeIt(fn, times = 5) {
  const samples = []
  for (let i = 0; i < times; i++) {
    const t0 = performance.now()
    await fn()
    samples.push(performance.now() - t0)
  }
  return median(samples)
}

const smallDocx = await readFile(join(fixtures, 'sample.docx'))
const smallXlsx = await readFile(join(fixtures, 'sample.xlsx'))
const bigDocxBytes = await readFile(join(fixtures, 'big.docx'))
const bigXlsxBytes = await readFile(join(fixtures, 'big.xlsx'))

const tSmallDocx = await timeIt(() => readDocx(smallDocx))
const tSmallXlsx = await timeIt(() => readXlsx(smallXlsx))
const tBigDocx = await timeIt(() => readDocx(bigDocxBytes))
const tBigXlsx = await timeIt(() => readXlsx(bigXlsxBytes))
const tMdRead = await timeIt(async () => { await readFile(join(fixtures, 'big-docx.md'), 'utf8') })

// ---------- 输出体积 ----------
const strip = (s) => s.replace(/\s+/gu, '')
// office_read 的渲染体积 ≈ 头两行 + 每单元标头 + 正文（复刻 renderer 关键结构）
const officeDocxRendered = `头两行约120\n${docParas.map((p, i) => `\n—— 段落 ${i + 1}（${p.chars} 字符）——\n${p.text}`).join('')}`
const docx = readDocx(bigDocxBytes)
const bigXlsxDoc = readXlsx(bigXlsxBytes)
const xlsxRendered = `头两行约120\n${bigXlsxDoc.rows.map((r) => `\n—— ${r.sheet}!第${r.row}行（${r.chars} 字符）——\n${r.text}`).join('')}`
const mdDocxChars = mdDocx.length
const mdXlsxChars = mdXlsx.length
// MD 经 read 工具读取会带行号（约 8 字符/行）
const mdDocxWithLineNums = mdDocxChars + mdDocx.split('\n').length * 8
const mdXlsxWithLineNums = mdXlsxChars + mdXlsx.split('\n').length * 8

const fmt = (n) => (n / 1024).toFixed(1) + 'KB'
console.log(`
===== 解析耗时（中位数，5 次）=====
sample.docx（7 段）          ${tSmallDocx.toFixed(1)} ms
sample.xlsx（7 行）          ${tSmallXlsx.toFixed(1)} ms
big.docx（${PARA} 段）       ${tBigDocx.toFixed(1)} ms
big.xlsx（${ROWS} 行 × 5 列） ${tBigXlsx.toFixed(1)} ms
read(big-docx.md) 文件 IO     ${tMdRead.toFixed(2)} ms

===== 输出体积（同内容）=====
docx 正文净字符        ${fmt(strip(JSON.stringify(docParas.map((p) => p.text))).length)}
docx → MD（含行号）    ${fmt(mdDocxWithLineNums)}
docx → office_read     ${fmt(officeDocxRendered.length)}（≈ MD 的 ${(officeDocxRendered.length / mdDocxWithLineNums * 100).toFixed(0)}%）
xlsx 表格净字符        ${fmt(strip(JSON.stringify(bigXlsxDoc.rows.map((r) => r.text))).length)}
xlsx → MD 表格（行号） ${fmt(mdXlsxWithLineNums)}
xlsx → office_read     ${fmt(xlsxRendered.length)}（≈ MD 的 ${(xlsxRendered.length / mdXlsxWithLineNums * 100).toFixed(0)}%）
`)
