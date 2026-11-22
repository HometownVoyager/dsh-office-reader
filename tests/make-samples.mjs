/**
 * 生成三种格式的最小测试样例（不依赖 Office 软件）。
 * 用法：node tests/make-samples.mjs
 */
import { zipSync, strToU8 } from 'fflate'
import { mkdirSync, writeFileSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import XLSX from 'xlsx'

const root = dirname(dirname(fileURLToPath(import.meta.url)))
const fixtures = join(root, 'tests', 'fixtures')
mkdirSync(fixtures, { recursive: true })

// ---------- sample.docx ----------
const docxParagraph = (text, { heading } = {}) =>
  `<w:p>${heading !== undefined
    ? `<w:pPr><w:pStyle w:val="Heading${heading}"/><w:outlineLvl w:val="${heading - 1}"/></w:pPr>`
    : ''}<w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`

const docxDocument = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:body>
${docxParagraph('项目说明', { heading: 1 })}
${docxParagraph('本文档描述 Office 阅读器的验收要点。')}
${docxParagraph('背景', { heading: 2 })}
${docxParagraph('此前阅读 docx 需要先解压再提取，工具化之后一步到位。')}
${docxParagraph('段落中包含\t制表符与 &amp; 实体转义。')}
${docxParagraph('风险与对策', { heading: 1 })}
${docxParagraph('旧版二进制格式不受支持，需在 Office 中另存为 OOXML。')}
</w:body>
</w:document>`

const docxCore = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/">
<dc:title>Office 阅读器验收文档</dc:title>
<dc:creator>测试员</dc:creator>
<cp:lastModifiedBy>测试员二号</cp:lastModifiedBy>
<dcterms:created>2026-09-10T00:00:00Z</dcterms:created>
</cp:coreProperties>`

const docxContentTypes = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
</Types>`

const docxRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`

writeFileSync(join(fixtures, 'sample.docx'), zipSync({
  '[Content_Types].xml': strToU8(docxContentTypes),
  '_rels/.rels': strToU8(docxRels),
  'docProps/core.xml': strToU8(docxCore),
  'word/document.xml': strToU8(docxDocument),
}))

// ---------- sample.pptx ----------
// 真实 deck 的顺序由 ppt/presentation.xml 的 <p:sldIdLst> 决定，部件编号只是创建时分配的。
// 本夹具刻意让 sldIdLst 的顺序（slide2 → slide1）与文件名数字顺序**相反**，
// 这样「按文件名排序」的实现会在这里读错页；备注走 slide 自己的 rels 解析，
// 且 slide2 的备注部件编号（notesSlide1）与它自己的编号不同，用来钉住「不许按文件名猜备注」。
const slideXml = (title, bullets, extra = '') => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
<p:cSld><p:spTree>
<p:sp><p:nvSpPr><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr><p:txBody><a:bodyPr/><a:p><a:r><a:t>${title}</a:t></a:r></a:p></p:txBody></p:sp>
<p:sp><p:nvSpPr><p:nvPr><p:ph type="body"/></p:nvPr></p:nvSpPr><p:txBody><a:bodyPr/>${bullets.map((b) => `<a:p><a:r><a:t>${b}</a:t></a:r></a:p>`).join('')}</p:txBody></p:sp>
${extra}
</p:spTree></p:cSld></p:sld>`

const notesXml = (text) => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:notes xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
<p:cSld><p:spTree>
<p:sp><p:nvSpPr><p:nvPr><p:ph type="body"/></p:nvPr></p:nvSpPr><p:txBody><a:bodyPr/><a:p><a:r><a:t>${text}</a:t></a:r></a:p></p:txBody></p:sp>
<p:sp><p:nvSpPr><p:nvPr><p:ph type="sldNum"/></p:nvPr></p:nvSpPr><p:txBody><a:bodyPr/><a:p><a:fld id="{X}" type="slidenum"><a:t>2</a:t></a:fld></a:p></p:txBody></p:sp>
</p:spTree></p:cSld></p:notes>`

const pptxCore = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/">
<dc:title>年度规划演示</dc:title>
<dc:creator>演示者</dc:creator>
</cp:coreProperties>`

const pptxContentTypes = docxContentTypes
  .replace('/word/document.xml', '/ppt/presentation.xml')
  .replace('wordprocessingml.document.main', 'presentationml.presentation.main')

const relsXml = (entries) => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
${entries.map(([id, type, target]) => `<Relationship Id="${id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/${type}" Target="${target}"/>`).join('\n')}
</Relationships>`

writeFileSync(join(fixtures, 'sample.pptx'), zipSync({
  '[Content_Types].xml': strToU8(pptxContentTypes),
  '_rels/.rels': strToU8(docxRels),
  'docProps/core.xml': strToU8(pptxCore),
  'ppt/presentation.xml': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<p:sldIdLst><p:sldId id="257" r:id="rId2"/><p:sldId id="256" r:id="rId1"/></p:sldIdLst>
</p:presentation>`),
  'ppt/_rels/presentation.xml.rels': strToU8(relsXml([['rId1', 'slide', 'slides/slide1.xml'], ['rId2', 'slide', 'slides/slide2.xml']])),
  'ppt/slides/slide1.xml': strToU8(slideXml('2026 年度规划', ['目标：工具化阅读 Office 文档', '路径：克隆 pdf-reader 模式'])),
  'ppt/slides/slide2.xml': strToU8(slideXml('预算与排期', ['预算：一台咖啡机', '排期：本周五前上线'], '<p:sp><p:nvSpPr><p:nvPr><p:ph type="ftr"/></p:nvPr></p:nvSpPr><p:txBody><a:bodyPr/><a:p><a:r><a:t>模板页脚不该被当成正文</a:t></a:r></a:p></p:txBody></p:sp>')),
  'ppt/slides/_rels/slide1.xml.rels': strToU8(relsXml([['rId1', 'notesSlide', '../notesSlides/notesSlide2.xml']])),
  'ppt/slides/_rels/slide2.xml.rels': strToU8(relsXml([['rId1', 'notesSlide', '../notesSlides/notesSlide1.xml']])),
  'ppt/notesSlides/notesSlide1.xml': strToU8(notesXml('发布前核对数据来源。')),
  'ppt/notesSlides/notesSlide2.xml': strToU8(notesXml('年度规划页的讲稿。')),
}))

// ---------- sample.xlsx ----------
const workbook = XLSX.utils.book_new()
const sales = XLSX.utils.aoa_to_sheet([
  ['日期', '产品', '金额'],
  [new Date('2026-01-05T00:00:00Z'), '咖啡机', 1200],
  [new Date('2026-02-01T00:00:00Z'), '显示器', 980],
  [null, '合计', 2180],
])
sales['!ref'] = 'A1:C4'
XLSX.utils.book_append_sheet(workbook, sales, '销售数据')
const memo = XLSX.utils.aoa_to_sheet([['备注'], ['含关键词验收'], ['旧格式 .doc 不支持']])
XLSX.utils.book_append_sheet(workbook, memo, '备注')
XLSX.writeFile(workbook, join(fixtures, 'sample.xlsx'), { bookType: 'xlsx', type: 'buffer', cellDates: true })

console.log('fixtures written to', fixtures)
