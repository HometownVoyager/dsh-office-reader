/**
 * P2 回归（docx 读取面）：比对报告 §6 P2「docx：修订检测 → 接受全部；批注；读取器部件覆盖
 * （页眉页脚/脚注/尾注/批注）；行单元表格读取」。
 *
 * 写侧用 dsh-office-editor 的构建产物（add_comment / createDocx 的 header/footer/table），
 * 读侧用本插件的 readDocx + office_read/office_audit 工具，形成闭链。
 * 用法：node tests/p2-docx.mjs（在 dsh-office-reader 目录下）
 */
import { mkdirSync, writeFileSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate'
import { loadReaderTools, makeChecker } from './harness.mjs'
import { readDocx } from '../lib/docx.js'
import { openDocument, selectUnits } from '../lib/office.js'

const here = dirname(fileURLToPath(import.meta.url))
const editorLib = join(here, '..', '..', 'dsh-office-editor', 'lib', 'docx.js')
const editor = await import(new URL(`file://${editorLib.replace(/\\/gu, '/')}`).href)

const { check, done } = makeChecker('docx 读取面')
const { call, render } = await loadReaderTools()

const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
const R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
const CT_NS = 'http://schemas.openxmlformats.org/package/2006/content-types'
const REL_NS = 'http://schemas.openxmlformats.org/package/2006/relationships'
const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'

/** Minimal docx wrapper for fixtures the editor cannot write (footnotes/endnotes, revisions). */
function rawDocx(body, parts = {}, contentTypes = []) {
  return zipSync({
    '[Content_Types].xml': strToU8(`${XML_DECL}<Types xmlns="${CT_NS}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>${contentTypes.join('')}</Types>`),
    '_rels/.rels': strToU8(`${XML_DECL}<Relationships xmlns="${REL_NS}"><Relationship Id="rId1" Type="${R_NS}/officeDocument" Target="word/document.xml"/></Relationships>`),
    'word/document.xml': strToU8(`${XML_DECL}<w:document xmlns:w="${W_NS}" xmlns:r="${R_NS}"><w:body>${body}<w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr></w:body></w:document>`),
    ...parts,
  })
}

console.log('== 1. add_comment：批注部件、锚点与读回 ==')
{
  const base = editor.createDocx([
    { type: 'paragraph', text: '第一章 研究背景' },
    { type: 'paragraph', text: '本文研究的问题是样本量不足。' },
    { type: 'paragraph', text: '结尾段落。' },
  ])
  const out = await editor.editDocx(base, [
    { op: 'add_comment', index: 2, text: '这里的样本量需要补一个说明。', author: '导师' },
    { op: 'add_comment', index: 2, text: '第二条批注', author: '张三', anchor: '样本量不足' },
  ])
  const files = unzipSync(out.bytes)
  const docXml = strFromU8(files['word/document.xml'])
  const commentsXml = strFromU8(files['word/comments.xml'])
  check('生成 word/comments.xml', files['word/comments.xml'] !== undefined)
  check('content-type 登记 comments', /PartName="\/word\/comments.xml"/u.test(strFromU8(files['[Content_Types].xml'])))
  check('document.xml.rels 登记 comments 关系', new RegExp(`${R_NS}/comments`).test(strFromU8(files['word/_rels/document.xml.rels'])))
  check('批注正文写入部件', commentsXml.includes('这里的样本量需要补一个说明。') && commentsXml.includes('第二条批注'), commentsXml)
  check('批注作者写入部件', commentsXml.includes('w:author="导师"') && commentsXml.includes('w:author="张三"'), commentsXml)
  check('批注 id 递增（0、1）', /w:id="0"/u.test(commentsXml) && /w:id="1"/u.test(commentsXml), commentsXml)
  check('commentRangeStart/End 成对', (docXml.match(/<w:commentRangeStart\b/gu) ?? []).length === 2 && (docXml.match(/<w:commentRangeEnd\b/gu) ?? []).length === 2, docXml)
  check('引用 run 用 CommentReference 样式', (docXml.match(/<w:rStyle w:val="CommentReference"\/>/gu) ?? []).length === 2)
  check('styles.xml 补上批注样式', /w:styleId="CommentReference"/u.test(strFromU8(files['word/styles.xml'])) && /w:styleId="CommentText"/u.test(strFromU8(files['word/styles.xml'])))
  check('带 anchor 的批注不改动 run 结构（区间扩到整段所在 run）', /<w:commentRangeStart w:id="1"\/><w:r><w:t xml:space="preserve">本文研究的问题是样本量不足。<\/w:t><\/w:r><w:commentRangeEnd w:id="1"\/>/u.test(docXml), docXml)
  check('改批注不改正文文本', editor.scanParas(docXml)[1].text === '本文研究的问题是样本量不足。', editor.scanParas(docXml)[1].text)
  check('verifyDocx 通过', (() => { try { editor.verifyDocx(out.bytes); return true } catch { return false } })())

  const parsed = readDocx(out.bytes)
  const comments = parsed.parts.filter((part) => part.kind === 'comment')
  check('读取器读到 2 条批注', comments.length === 2, JSON.stringify(comments))
  check('批注作者可读（UTF-8 正确）', comments[1].author === '张三', JSON.stringify(comments[1]))
  check('批注文本可读', comments[0].text === '这里的样本量需要补一个说明。', JSON.stringify(comments[0]))
  check('批注不影响段落编号契约', parsed.paragraphs.length === 3, String(parsed.paragraphs.length))

  const doc = openDocument(out.bytes, 'comments.docx')
  const sel = selectUnits(doc, { parts: 'comments' }, 200)
  check('parts=comments 选中 2 个单元', sel.units.length === 2, JSON.stringify(sel.units.map((u) => u.label)))
  check('批注单元标签带作者', sel.units[1].label.includes('张三'), sel.units[1].label)
  let noParts = ''
  try { selectUnits(doc, { parts: 'footnotes' }, 200) } catch (error) { noParts = error.message }
  check('无该类部件时明确报错', noParts.includes('没有匹配'), noParts)

  let anchorMiss = ''
  try { await editor.editDocx(base, [{ op: 'add_comment', index: 2, text: 'x', anchor: '不存在的片段' }]) } catch (error) { anchorMiss = error.message }
  check('anchor 未命中时明确报错', anchorMiss.includes('批注锚点未在段落中找到'), anchorMiss)

  // 多 run 段落：锚点落在第 2 个 run 上时，只有那个 run 被包进批注区间。
  const multiRun = rawDocx('<w:p><w:r><w:t xml:space="preserve">前半 </w:t></w:r><w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">加粗的锚点</w:t></w:r><w:r><w:t xml:space="preserve"> 后半</w:t></w:r></w:p>')
  const anchored = await editor.editDocx(multiRun, [{ op: 'add_comment', index: 1, text: '只批注中间那段', anchor: '加粗的锚点' }])
  const anchoredXml = strFromU8(unzipSync(anchored.bytes)['word/document.xml'])
  check('多 run 段落：只包住命中 run', /<w:commentRangeStart w:id="0"\/><w:r><w:rPr><w:b\/><\/w:rPr><w:t xml:space="preserve">加粗的锚点<\/w:t><\/w:r><w:commentRangeEnd w:id="0"\/>/u.test(anchoredXml), anchoredXml)
  check('多 run 段落：其余 run 原样', anchoredXml.includes('>前半 </w:t>') && anchoredXml.includes('> 后半</w:t>'))
  check('多 run 段落：段落文本未变', readDocx(anchored.bytes).paragraphs[0].text === '前半 加粗的锚点 后半', readDocx(anchored.bytes).paragraphs[0].text)
}

console.log('== 2. 读取器部件覆盖：页眉页脚 ==')
{
  const bytes = editor.createDocx([{ type: 'paragraph', text: '正文一段。' }], {
    header: '<w:p><w:r><w:t xml:space="preserve">学位论文页眉</w:t></w:r></w:p>',
    footer: '<w:p><w:r><w:t xml:space="preserve">第 1 页</w:t></w:r></w:p>',
  })
  const parsed = readDocx(bytes)
  const header = parsed.parts.find((part) => part.kind === 'header')
  const footer = parsed.parts.find((part) => part.kind === 'footer')
  check('读到页眉部件', header?.text === '学位论文页眉' && header.label === '页眉 1', JSON.stringify(header))
  check('读到页脚部件', footer?.text === '第 1 页', JSON.stringify(footer))
  const doc = openDocument(bytes, 'hf.docx')
  check('units 仍是段落（页眉不混入编号）', doc.units.length === 1 && doc.units[0].label === '段落 1', JSON.stringify(doc.units.map((u) => u.label)))
  const all = selectUnits(doc, { parts: 'all' }, 200)
  check('parts=all 返回两个部件', all.units.length === 2, JSON.stringify(all.units.map((u) => u.label)))
  const headers = selectUnits(doc, { parts: 'headers' }, 200)
  check('parts=headers 只返回页眉', headers.units.length === 1 && headers.units[0].text === '学位论文页眉')
  let conflict = ''
  try { selectUnits(doc, { parts: 'all', tables: true }, 200) } catch (error) { conflict = error.message }
  check('parts 与 tables 互斥', conflict.includes('不能同时指定'), conflict)
  let withParagraphs = ''
  try { selectUnits(doc, { parts: 'all', paragraphs: '1' }, 200) } catch (error) { withParagraphs = error.message }
  check('parts 与 paragraphs 互斥', withParagraphs.includes('互斥'), withParagraphs)
}

console.log('== 3. 读取器部件覆盖：脚注与尾注 ==')
{
  const notes = (tag, label) => `${XML_DECL}<w:${tag}s xmlns:w="${W_NS}"><w:${tag} w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:${tag}><w:${tag} w:id="0"><w:p><w:r><w:t>分隔符占位</w:t></w:r></w:p></w:${tag}><w:${tag} w:id="3"><w:p><w:r><w:t>${label}</w:t></w:r></w:p></w:${tag}></w:${tag}s>`
  const body = '<w:p><w:r><w:t>正文带脚注引用</w:t></w:r><w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteReference w:id="3"/></w:r></w:p>'
  const bytes = rawDocx(body, {
    'word/footnotes.xml': strToU8(notes('footnote', '这是脚注正文')),
    'word/endnotes.xml': strToU8(notes('endnote', '这是尾注正文')),
  }, [
    `<Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/>`,
    `<Override PartName="/word/endnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml"/>`,
  ])
  const parsed = readDocx(bytes)
  const footnote = parsed.parts.find((part) => part.kind === 'footnote')
  const endnote = parsed.parts.find((part) => part.kind === 'endnote')
  check('脚注按 id 命名并读到正文', footnote?.label === '脚注 3' && footnote.text === '这是脚注正文', JSON.stringify(footnote))
  check('尾注同理', endnote?.label === '尾注 3' && endnote.text === '这是尾注正文', JSON.stringify(endnote))
  check('分隔符脚注（id -1/0）被跳过', parsed.parts.filter((part) => part.kind === 'footnote').length === 1)
  const doc = openDocument(bytes, 'notes.docx')
  const sel = selectUnits(doc, { parts: 'footnotes' }, 200)
  check('parts=footnotes 选中脚注', sel.units.length === 1 && sel.units[0].text === '这是脚注正文', JSON.stringify(sel.units))
  const result = await call('office_read', { source: writeTemp('notes.docx', bytes), parts: 'all' })
  check('office_read parts=all 返回 2 个单元', result.units.length === 2, JSON.stringify(result.units.map((u) => u.unit)))
}

console.log('== 4. 表格行单元读取 ==')
{
  const bytes = editor.createDocx([
    { type: 'paragraph', text: '表前一段' },
    { type: 'table', header: true, rows: [['指标', '基线', '本方法'], ['准确率', '81.2', '88.7'], ['耗时(s)', '12.4', '9.1']] },
    { type: 'paragraph', text: '表后一段' },
  ])
  const parsed = readDocx(bytes)
  check('读到 3 行表格', parsed.tables.length === 3, JSON.stringify(parsed.tables.map((r) => r.text)))
  check('单元格以制表符分隔', parsed.tables[1].text === '准确率\t81.2\t88.7', parsed.tables[1].text)
  check('首行标为表头', parsed.tables[0].header === true && parsed.tables[1].header === false)
  check('行号与表号', parsed.tables[0].table === 1 && parsed.tables[0].row === 1 && parsed.tables[2].row === 3)
  // 段落契约：表格单元格内段落与正文段落同级编号，所以 2 段正文 + 3 行 × 3 单元格 = 11。
  check('表格单元格段落仍计入段落编号（契约）', parsed.paragraphs.length === 11, String(parsed.paragraphs.length))
  check('正文段落是第 1 段与第 11 段', parsed.paragraphs[0].text === '表前一段' && parsed.paragraphs[10].text === '表后一段', JSON.stringify(parsed.paragraphs.map((p) => p.text)))
  const doc = openDocument(bytes, 'table.docx')
  const sel = selectUnits(doc, { tables: true }, 200)
  check('tables=true 选中全部行', sel.units.length === 3 && sel.units[1].label === '表格 1 第 2 行', JSON.stringify(sel.units.map((u) => u.label)))
  const read = await call('office_read', { source: writeTemp('table.docx', bytes), tables: true })
  check('office_read tables=true 走行单元模式', read.mode === 'tables' && read.units.length === 3, JSON.stringify(read.mode))
  check('渲染结果含行标签', render('office_read', read)[0].text.includes('表格 1 第 2 行'))
}

console.log('== 5. 修订检测 ==')
{
  const body = '<w:p><w:r><w:t>原文</w:t></w:r><w:ins w:id="1" w:author="甲"><w:r><w:t>插入的内容</w:t></w:r></w:ins><w:del w:id="2" w:author="甲"><w:r><w:delText>删掉的内容</w:delText></w:r></w:del></w:p>'
    + '<w:p><w:pPr><w:rPr><w:rPrChange w:id="3" w:author="甲"><w:rPr><w:b/></w:rPr></w:rPrChange></w:rPr></w:pPr><w:r><w:t>带格式修订的段落</w:t></w:r></w:p>'
  const bytes = rawDocx(body)
  const parsed = readDocx(bytes)
  check('检出插入一项', parsed.revisions.insertions === 1, JSON.stringify(parsed.revisions))
  check('检出删除一项', parsed.revisions.deletions === 1, JSON.stringify(parsed.revisions))
  check('检出格式变更一项', parsed.revisions.formatChanges === 1, JSON.stringify(parsed.revisions))
  check('present 为真且列出部件', parsed.revisions.present === true && parsed.revisions.parts.includes('word/document.xml'))
  const read = await call('office_read', { source: writeTemp('revisions.docx', bytes) })
  check('office_read 回传修订计数', read.revisions.present === true && read.revisions.insertions === 1, JSON.stringify(read.revisions))
  check('渲染带修订警告', render('office_read', read)[0].text.includes('未接受的修订'), render('office_read', read)[0].text.slice(0, 200))
  const accept = await editor.editDocx(bytes, [{ op: 'accept_revisions' }])
  const after = readDocx(accept.bytes)
  check('accept_revisions 后读取器不再报修订', after.revisions.present === false, JSON.stringify(after.revisions))
  check('接受后插入内容保留、删除内容消失', after.paragraphs[0].text === '原文插入的内容', after.paragraphs[0].text)
  const inspected = await call('office_inspect', { source: writeTemp('revisions2.docx', bytes) })
  check('office_inspect 暴露修订计数', inspected.docx.revisions.present === true && inspected.docx.revisions.format_changes === 1, JSON.stringify(inspected.docx.revisions))
}

console.log('== 6. 部件与表格计数进入 office_inspect ==')
{
  const bytes = editor.createDocx([
    { type: 'paragraph', text: '正文' },
    { type: 'table', rows: [['a', 'b']] },
  ], { header: '<w:p><w:r><w:t>页眉</w:t></w:r></w:p>' })
  const inspected = await call('office_inspect', { source: writeTemp('inspect.docx', bytes) })
  check('inspect 列出页眉部件', inspected.docx.parts.length === 1 && inspected.docx.parts[0].kind === 'header', JSON.stringify(inspected.docx.parts))
  check('inspect 报表格数与行数', inspected.docx.table_count === 1 && inspected.docx.table_rows === 1, JSON.stringify(inspected.docx))
  check('inspect 渲染含页眉与表格行', /正文之外：页眉 1/u.test(render('office_inspect', inspected)[0].text))
}

console.log('== 7. office_audit：结构体检与内容信号 ==')
{
  const good = editor.createDocx([{ type: 'paragraph', text: '干净文档' }], { toc: true })
  const goodReport = await call('office_audit', { source: writeTemp('audit-good.docx', good) })
  check('干净文档无 ERROR', goodReport.ok === true && !goodReport.issues.some((i) => i.severity === 'ERROR'), JSON.stringify(goodReport.issues))
  check('报出条数与段落数', goodReport.entry_count > 5 && goodReport.counts.paragraphs >= 1, JSON.stringify(goodReport.counts))

  const damaged = unzipSync(good)
  delete damaged['word/document.xml']
  const missingReport = await call('office_audit', { source: writeTemp('audit-missing.docx', zipSync(damaged)) })
  check('缺必需部件报 ERROR', missingReport.ok === false && missingReport.issues.some((i) => i.severity === 'ERROR' && i.message.includes('缺少必需部件')), JSON.stringify(missingReport.issues))

  const broken = unzipSync(good)
  broken['word/styles.xml'] = strToU8('<w:styles><w:style></w:styles>')
  const brokenReport = await call('office_audit', { source: writeTemp('audit-broken.docx', zipSync(broken)) })
  check('XML 良构问题报 ERROR', brokenReport.issues.some((i) => i.severity === 'ERROR' && i.message.includes('标签不匹配')), JSON.stringify(brokenReport.issues))

  const leftover = editor.createDocx([{ type: 'paragraph', text: '研究缺口：{{gap}}；TODO 补充' }])
  const leftoverReport = await call('office_audit', { source: writeTemp('audit-leftover.docx', leftover) })
  check('残留占位符报 WARN', leftoverReport.issues.some((i) => i.severity === 'WARN' && i.message.includes('占位符')), JSON.stringify(leftoverReport.issues))
  check('草稿标记报 INFO', leftoverReport.issues.some((i) => i.severity === 'INFO'), JSON.stringify(leftoverReport.issues))

  const withRevision = rawDocx('<w:p><w:ins w:id="1" w:author="甲"><w:r><w:t>新增</w:t></w:r></w:ins><w:r><w:t>正文</w:t></w:r></w:p>')
  const revReport = await call('office_audit', { source: writeTemp('audit-rev.docx', withRevision) })
  check('修订痕迹报 WARN 并给处置建议', revReport.issues.some((i) => i.severity === 'WARN' && i.message.includes('accept_revisions')), JSON.stringify(revReport.issues))
  check('audit 渲染带严重度前缀', /\[(WARN|ERROR|INFO)\]/u.test(render('office_audit', revReport)[0].text))

  const notZip = strToU8('这不是 OOXML')
  let notZipMessage = ''
  try { await call('office_audit', { source: writeTemp('audit-notzip.docx', notZip) }) } catch (error) { notZipMessage = error.message }
  check('非 zip 输入明确报错', notZipMessage.includes('无法审计'), notZipMessage)
}

process.exit(done())

/** Write bytes into tests/tmp/ and return the path (the tools take a path, not bytes). */
function writeTemp(name, bytes) {
  const dir = join(here, 'tmp')
  mkdirSync(dir, { recursive: true })
  const out = join(dir, name)
  writeFileSync(out, bytes)
  return out
}
