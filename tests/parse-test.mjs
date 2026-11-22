/**
 * 解析层单测：对 fixtures 三件套跑 docx/pptx/xlsx 读取器并断言关键内容。
 * 用法：node tests/parse-test.mjs
 */
import { readFile } from 'node:fs/promises'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { readDocx } from '../lib/docx.js'
import { readPptx } from '../lib/pptx.js'
import { readXlsx, readXlsxRangeRows } from '../lib/xlsx.js'
import { openDocument, selectUnits } from '../lib/office.js'
import { normalizeForMatch } from '../lib/search.js'

const root = dirname(dirname(fileURLToPath(import.meta.url)))
const fixture = (name) => join(root, 'tests', 'fixtures', name)
let failures = 0
const check = (name, condition, detail = '') => {
  if (condition) console.log(`  ok - ${name}`)
  else { failures++; console.log(`  FAIL - ${name} ${detail}`) }
}

const docxBytes = await readFile(fixture('sample.docx'))
const pptxBytes = await readFile(fixture('sample.pptx'))
const xlsxBytes = await readFile(fixture('sample.xlsx'))

console.log('== docx ==')
const docx = readDocx(docxBytes)
check('段落数 7', docx.paragraphs.length === 7, `got ${docx.paragraphs.length}`)
check('元数据标题', docx.metadata.title === 'Office 阅读器验收文档', JSON.stringify(docx.metadata))
check('outlineLvl 标题1', docx.paragraphs[0].level === 1)
check('pStyle 标题2', docx.paragraphs[2].level === 2)
check('tab 转换', docx.paragraphs[4].text.includes('\t'), JSON.stringify(docx.paragraphs[4]))
check('实体解码', docx.paragraphs[4].text.includes('&'), JSON.stringify(docx.paragraphs[4]))

console.log('== pptx ==')
const pptx = readPptx(pptxBytes)
check('幻灯片数 2', pptx.slides.length === 2, `got ${pptx.slides.length}`)
// 夹具刻意让 sldIdLst 顺序（slide2 → slide1）与文件名数字顺序相反：
// 第 1 页必须是 sldIdLst 的第一项（预算与排期，部件 slide2.xml），不是 slide1.xml。
check('按 sldIdLst 读顺序（第 1 页 = 部件 slide2.xml）', pptx.slides[0].part === 'ppt/slides/slide2.xml'
  && pptx.slides[0].title === '预算与排期', JSON.stringify({ part: pptx.slides[0].part, title: pptx.slides[0].title }))
check('标题占位符', pptx.slides[1].title === '2026 年度规划', pptx.slides[1].title)
check('备注走 slide 自己的 rels（第 1 页 = notesSlide1）', pptx.slides[0].notes === '发布前核对数据来源。', JSON.stringify(pptx.slides[0].notes))
check('备注并入正文', pptx.slides[0].text.includes('[备注]'))
check('模板页脚占位符不进入正文', !pptx.slides[0].text.includes('模板页脚不该被当成正文'), pptx.slides[0].text)
check('备注中的页码域不进入备注', !/(^|\n)\s*2\s*($|\n)/u.test(pptx.slides[0].notes ?? ''), JSON.stringify(pptx.slides[0].notes))
check('元数据标题', pptx.metadata.title === '年度规划演示')

console.log('== xlsx ==')
const xlsx = readXlsx(xlsxBytes)
const xlsxRows = xlsx.rows()
check('工作表 2 个', xlsx.sheets.length === 2, JSON.stringify(xlsx.sheets.map((s) => s.name)))
check('销售数据 4 行', xlsxRows.filter((r) => r.sheet === '销售数据').length === 4)
check('表头行', xlsxRows[0].text === '日期\t产品\t金额', xlsxRows[0].text)
check('数字格式化', xlsxRows[1].cells.includes('1200'), JSON.stringify(xlsxRows[1]))
check('日期格式化', /2026|26/.test(xlsxRows[1].cells[0]), xlsxRows[1].cells[0])
// 概览聚合（P2 补齐）：表头、列类型、公式数、样本行都在一次扫描里得出。
const sales = xlsx.sheets[0]
check('概览：表头单元', JSON.stringify(sales.headerCells) === JSON.stringify(['日期', '产品', '金额']), JSON.stringify(sales.headerCells))
check('概览：列类型推断', JSON.stringify(sales.columnTypes) === JSON.stringify(['datetime', 'str', 'int']), JSON.stringify(sales.columnTypes))
check('概览：样本行有界（≤3）', sales.sampleRows.length <= 3, JSON.stringify(sales.sampleRows))
check('概览：charCount 无需物化即可得', xlsx.charCount === xlsxRows.reduce((t, r) => t + r.chars, 0), `${xlsx.charCount}`)
check('概览：可见性字段', sales.visibility === 'visible' && sales.hidden === false, sales.visibility)

console.log('== office.selectUnits ==')
const doc = openDocument(xlsxBytes, fixture('sample.xlsx'))
const sheetSel = selectUnits(doc, { sheet: '销售数据' }, 200)
check('sheet 选择', sheetSel.units.length === 4, `got ${sheetSel.units.length}`)
const rangeSel = selectUnits(doc, { sheet: '备注', range: 'A2:A3' }, 200)
check('range 选择', rangeSel.units.length === 2 && rangeSel.units[0].text === '含关键词验收', JSON.stringify(rangeSel.units))
let threw = false
try { selectUnits(doc, { slides: '1' }, 200) } catch { threw = true }
check('xlsx 拒绝 slides 选择器', threw)

const docDoc = openDocument(docxBytes, fixture('sample.docx'))
const secSel = selectUnits(docDoc, { section: '背景' }, 200)
check('section 章节切片', secSel.units.length === 3 && secSel.units[0].text === '背景' && secSel.units[2].text.includes('制表符'), JSON.stringify(secSel.units.map((u) => u.text)))
const paraSel = selectUnits(docDoc, { paragraphs: '2-3' }, 200)
check('paragraphs 范围', paraSel.units[0].text.includes('验收要点'), paraSel.units[0].text)

const pptDoc = openDocument(pptxBytes, fixture('sample.pptx'))
const slideSel = selectUnits(pptDoc, { slides: '2' }, 200)
check('slides 范围按演示顺序取第 2 页', slideSel.units.length === 1 && slideSel.units[0].text.includes('2026 年度规划'), JSON.stringify(slideSel.units.map((u) => u.text)))

console.log('== find (search 层) ==')
const queryNorm = normalizeForMatch('关键词 验收')
const hit = selectUnits(doc, { sheet: '备注' }, 200).units.filter((u) => normalizeForMatch(u.text).includes(queryNorm))
check('空白容错命中', hit.length === 1 && hit[0].label === '备注!第2行', JSON.stringify(hit))

console.log(failures === 0 ? '\nALL PASS' : `\n${failures} FAILURES`)
process.exit(failures === 0 ? 0 : 1)
