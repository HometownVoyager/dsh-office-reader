/**
 * P2 回归（xlsx 读取面）：比对报告 §6 P2「xlsx：CSV 导出；overview 升级（公式数/表头/列类型）；
 * audit 工具；冻结窗格 XML 补丁；合并单元格操作 + 区间位移正确性；公式检视模式」中的**读侧**。
 *
 * 写侧用 dsh-office-editor 的构建产物，读侧用本插件的 readXlsx + office_inspect/office_read/
 * office_find/office_audit，形成闭链。
 * 用法：node tests/p2-xlsx.mjs（在 dsh-office-reader 目录下）
 */
import { mkdirSync, writeFileSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate'
import XLSX from 'xlsx'
import { loadReaderTools, makeChecker } from './harness.mjs'
import { readXlsx } from '../lib/xlsx.js'
import { openDocument, selectUnits } from '../lib/office.js'

const here = dirname(fileURLToPath(import.meta.url))
const editorLib = join(here, '..', '..', 'dsh-office-editor', 'lib', 'xlsx.js')
const editor = await import(new URL(`file://${editorLib.replace(/\\/gu, '/')}`).href)

const { check, done } = makeChecker('xlsx 读取面')
const { call, render } = await loadReaderTools()

const writeTemp = (name, bytes) => {
  const dir = join(here, 'tmp')
  mkdirSync(dir, { recursive: true })
  const out = join(dir, name)
  writeFileSync(out, bytes)
  return out
}

console.log('== 1. 造一份带公式/批注/超链接/定义名/隐藏表/冻结窗格的工作簿 ==')
const base = editor.createXlsx([
  {
    name: '销售',
    rows: [
      ['日期', '产品', '数量', '金额'],
      ['2026-01-01', '甲', 2, '=C2*10'],
      ['2026-01-02', '乙', 3, '=C3*10'],
      ['2026-01-03', '丙', 4, '=C4*10'],
    ],
    formulas: { F1: '=SUM(D2:D4)' },
  },
  { name: '备注', rows: [['说明'], ['隐藏表用于字典']] },
], 1000)
const edited = editor.editXlsx(base, [
  { op: 'set_cell_comment', sheet: '销售', cell: 'D1', text: '金额列按数量×单价计算', author: '李四' },
  { op: 'set_hyperlink', sheet: '销售', cell: 'B2', url: 'https://例子.测试/产品/甲', tooltip: '产品页' },
  { op: 'set_defined_name', name: '单价', refers_to: '销售!$C$1' },
  { op: 'set_sheet_state', name: '备注', state: 'hidden' },
  { op: 'freeze_panes', sheet: '销售', cell: 'A2' },
  { op: 'merge_cells', sheet: '销售', range: 'A5:C5' },
], 1000)
const bytes = new Uint8Array(edited.bytes)
const path = writeTemp('p2-sample.xlsx', bytes)

console.log('== 2. 概览聚合（公式数/表头/列类型/可见性） ==')
const parsed = readXlsx(bytes)
const sales = parsed.sheets.find((s) => s.name === '销售')
const notes = parsed.sheets.find((s) => s.name === '备注')
check('工作表顺序与索引', parsed.sheets.map((s) => s.name).join(',') === '销售,备注' && notes.index === 1, JSON.stringify(parsed.sheets.map((s) => s.name)))
check('公式计数（4 个）', sales.formulaCount === 4, String(sales.formulaCount))
check('表头单元', JSON.stringify(sales.headerCells) === JSON.stringify(['日期', '产品', '数量', '金额']), JSON.stringify(sales.headerCells))
// 金额列只有「无缓存值的公式」，因此单列类型是 formula 而不是 empty/mixed。
check('列类型推断（含 formula 列）', JSON.stringify(sales.columnTypes) === JSON.stringify(['str', 'str', 'int', 'formula']), JSON.stringify(sales.columnTypes))
check('样本行有界', sales.sampleRows.length === 3, String(sales.sampleRows.length))
check('隐藏表标为 hidden', notes.visibility === 'hidden' && notes.hidden === true, notes.visibility)
check('可见表标为 visible', sales.visibility === 'visible' && sales.hidden === false)
check('批注计入所属表', sales.commentCount === 1 && notes.commentCount === 0, `${sales.commentCount}/${notes.commentCount}`)
check('超链接计入所属表', sales.hyperlinkCount === 1, String(sales.hyperlinkCount))
check('charCount 与物化行一致', parsed.charCount === parsed.rows().reduce((t, r) => t + r.chars, 0), `${parsed.charCount}`)
check('公式清单拿到 A1 地址与公式', parsed.formulas.some((f) => f.addr === 'D2' && f.formula === 'C2*10'), JSON.stringify(parsed.formulas))
check('公式按行列序（先第 1 行的 F1）', parsed.formulas[0].addr === 'F1' && parsed.formulas[1].addr === 'D2', JSON.stringify(parsed.formulas.map((f) => f.addr)))
check('定义名可读', parsed.definedNames.some((n) => n.name === '单价' && n.refersTo === '销售!$C$1'), JSON.stringify(parsed.definedNames))

console.log('== 3. 批注与超链接按 UTF-8 读（绕开 SheetJS 的 latin1 解码） ==')
check('批注作者中文正确', parsed.comments[0]?.author === '李四', JSON.stringify(parsed.comments))
check('批注文本正确', parsed.comments[0]?.text === '金额列按数量×单价计算', JSON.stringify(parsed.comments))
check('批注挂在正确单元格', parsed.comments[0]?.addr === 'D1' && parsed.comments[0]?.sheet === '销售', JSON.stringify(parsed.comments[0]))
check('超链接目标中文正确', parsed.hyperlinks[0]?.url === 'https://例子.测试/产品/甲', JSON.stringify(parsed.hyperlinks))
check('超链接提示保留', parsed.hyperlinks[0]?.tooltip === '产品页', JSON.stringify(parsed.hyperlinks[0]))

console.log('== 4. 公式检视模式与选择器 ==')
const doc = openDocument(bytes, 'p2-sample.xlsx')
const formulaSel = selectUnits(doc, { formulas: true }, 500)
check('formulas=true 逐个公式成单元', formulaSel.units.length === 4 && formulaSel.units[1].label === '销售!D2', JSON.stringify(formulaSel.units.map((u) => u.label)))
check('公式单元文本是 =公式', formulaSel.units[1].text === '=C2*10', formulaSel.units[1].text)
const perSheet = selectUnits(doc, { formulas: true, sheet: '销售' }, 500)
check('formulas+sheet 限定工作表', perSheet.units.length === 4, String(perSheet.units.length))
let noFormula = ''
try { selectUnits(doc, { formulas: true, sheet: '备注' }, 500) } catch (error) { noFormula = error.message }
check('该表没有公式时明确报错', noFormula.includes('没有公式'), noFormula)
const sheetSel = selectUnits(doc, { sheet: '销售' }, 500)
check('sheet 选择仍是数据行', sheetSel.units.length === 4 && sheetSel.units[0].text.startsWith('日期\t产品'), JSON.stringify(sheetSel.units.map((u) => u.text)))
check('隐藏表也能按名读取', selectUnits(doc, { sheet: '备注' }, 500).units.length === 2)

console.log('== 5. 工具层：inspect / read / find ==')
const inspected = await call('office_inspect', { source: path })
check('inspect 报每表公式数', inspected.sheets[0].formulas === 4, JSON.stringify(inspected.sheets))
check('inspect 报表头与列类型', inspected.sheets[0].header[0] === '日期' && inspected.sheets[0].column_types[2] === 'int', JSON.stringify(inspected.sheets[0]))
check('inspect 报可见性', inspected.sheets[1].visibility === 'hidden', JSON.stringify(inspected.sheets[1]))
check('inspect 报定义名', inspected.defined_names.some((n) => n.name === '单价'), JSON.stringify(inspected.defined_names))
check('inspect 渲染含工作表属性与定义名', /工作表属性：/u.test(render('office_inspect', inspected)[0].text) && /定义名（1）/u.test(render('office_inspect', inspected)[0].text))

const readFormulas = await call('office_read', { source: path, formulas: true })
check('office_read formulas 模式', readFormulas.mode === 'formulas' && readFormulas.units.length === 4, JSON.stringify(readFormulas.mode))
check('公式渲染带 =公式', render('office_read', readFormulas)[0].text.includes('=C2*10'), render('office_read', readFormulas)[0].text.slice(0, 300))

const found = await call('office_find', { source: path, query: 'SUM', formulas: true })
check('office_find 可只搜公式', found.total_matches === 1 && found.matches[0].unit === '销售!F1', JSON.stringify(found.matches))

console.log('== 6. office_audit：公式缓存值与错误值 ==')
const cleanReport = await call('office_audit', { source: path })
check('经 office_edit 写出的工作簿有 calcPr（无 INFO）', !cleanReport.issues.some((i) => i.message.includes('calcPr')), JSON.stringify(cleanReport.issues))
check('报公式数与缺缓存值数', cleanReport.counts.formulas === 4 && cleanReport.counts.formulas_without_value === 4, JSON.stringify(cleanReport.counts))
check('公式无缓存值报 WARN', cleanReport.issues.some((i) => i.severity === 'WARN' && i.message.includes('没有缓存值')), JSON.stringify(cleanReport.issues))
check('批注/超链接计数进入 audit', cleanReport.counts.comments === 1 && cleanReport.counts.hyperlinks === 1, JSON.stringify(cleanReport.counts))

// 一张由 SheetJS 直接写出、没有 calcPr 的工作簿，并在里面放一个错误值单元格。
const wb = XLSX.utils.book_new()
const ws = XLSX.utils.aoa_to_sheet([['a', 1], ['b', '=1/0']])
ws['!ref'] = 'A1:C2' // C1 必须在已用区域内，否则不会被写进文件
ws.C1 = { t: 'e', v: 0x07, w: '#DIV/0!' }
XLSX.utils.book_append_sheet(wb, ws, 'S')
const rawBytes = new Uint8Array(XLSX.write(wb, { bookType: 'xlsx', type: 'buffer' }))
const rawReport = await call('office_audit', { source: writeTemp('p2-raw.xlsx', rawBytes) })
check('缺 calcPr 报 INFO', rawReport.issues.some((i) => i.severity === 'INFO' && i.message.includes('calcPr')), JSON.stringify(rawReport.issues))
check('错误值单元格报 WARN 并给地址', rawReport.issues.some((i) => i.severity === 'WARN' && i.message.includes('#DIV/0!') && i.message.includes('C1')), JSON.stringify(rawReport.issues))
check('audit 渲染按严重度排序', (() => {
  const text = render('office_audit', rawReport)[0].text
  const firstError = text.indexOf('[ERROR]')
  const firstInfo = text.indexOf('[INFO]')
  return firstError === -1 || firstInfo === -1 || firstError < firstInfo
})())

// 关系解析必须处理 `..` 与包根：批注关系是 `../comments2.xml`（相对工作表部件），
// 根关系是 `xl/workbook.xml`（相对包根）。直接拼接会把它们全报成「部件不存在」。
check('批注关系按 ../ 正确解析', !cleanReport.issues.some((i) => i.message.includes('关系指向的部件不存在')), JSON.stringify(cleanReport.issues))
{
  // 手工断开一条关系：把工作表的 comments 关系指向不存在的部件，确认体检仍会报出来。
  const files = unzipSync(bytes)
  const relsName = 'xl/worksheets/_rels/sheet1.xml.rels'
  files[relsName] = strToU8(strFromU8(files[relsName]).replace(/(Target="\.\.\/)comments\d*\.xml"/u, '$1commentsX.xml"'))
  const brokenReport = await call('office_audit', { source: writeTemp('p2-dangling.xlsx', zipSync(files)) })
  check('真的断链关系仍报 WARN', brokenReport.issues.some((i) => i.message.includes('关系指向的部件不存在') && i.message.includes('commentsX.xml')), JSON.stringify(brokenReport.issues))
}

console.log('== 7. 冻结窗格与合并区在读取侧可见 ==')
const frozenSheet = readXlsx(bytes).sheets.find((s) => s.name === '销售')
check('冻结窗格不破坏解析（表仍可读）', frozenSheet.rowCount === 4, String(frozenSheet.rowCount))
const merged = XLSX.read(bytes, { type: 'buffer' })
check('合并区往返保留', JSON.stringify(merged.Sheets['销售']['!merges']) === JSON.stringify([{ s: { c: 0, r: 4 }, e: { c: 2, r: 4 } }]), JSON.stringify(merged.Sheets['销售']['!merges']))

process.exit(done())
