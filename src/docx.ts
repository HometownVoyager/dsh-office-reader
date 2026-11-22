/**
 * .docx 解析：段落切分（嵌套感知）、标题层级（outlineLvl / pStyle）、正文提取，以及 P2 补齐的
 * 三块读取面——正文之外的部件（页眉页脚/脚注尾注/批注）、表格行单元、修订痕迹信号。
 *
 * 段落单元 = document.xml 中的**顶层** `<w:p>`，含空段落（文本框内嵌段落不是独立单元）。
 * 这一编号即 `office_read` 的 `paragraphs` 选择器与 `office_edit` 的 `index` 共用的契约。
 *
 * 正文之外的部件**不进入** `units` 序列：它们另成命名空间（`页眉 1`/`脚注 3`/`批注 2`），由
 * `parts` 选择器读取。混进段落序列会让 `段落 N` 偏移，那正是契约要防的事。
 *
 * @module dsh-office-reader/src/docx
 */

import {
  attribute,
  openOoxml,
  readCoreProps,
  splitElements,
  splitWordParagraphs,
  wordHeadingLevel,
  type CoreProps,
  type OoxmlPackage,
} from './ooxml.js'

/** One body paragraph with its 1-based index and optional heading level. */
export interface DocxParagraph {
  /** 1-based index among all top-level paragraphs (empty ones included, so the index is stable). */
  index: number
  /** Heading level (1-9) when the paragraph carries one. */
  level?: number
  /** Extracted text; empty for blank and page-break-only paragraphs. */
  text: string
  /** Text length in characters. */
  chars: number
  /** True when the paragraph carries no visible text (blank line, page break, section break). */
  empty: boolean
}

/** Kinds of non-body parts this reader exposes. */
export type DocxPartKind = 'header' | 'footer' | 'footnote' | 'endnote' | 'comment'

/** One unit of a non-body part (a header/footer as a whole, or one note/comment). */
export interface DocxPart {
  kind: DocxPartKind
  /** Part file name (e.g. `word/header1.xml`). */
  part: string
  /** User-facing label, e.g. `页眉 1` / `脚注 3` / `批注 0`. */
  label: string
  /** Plain text (paragraphs joined by newline; notes/comments keep their own paragraphs). */
  text: string
  chars: number
  /** Comment/note author, when the part carries one. */
  author?: string
  /** Comment/note timestamp (`w:date`), when present. */
  date?: string
  /** Note/comment id inside its part (`w:id`). */
  id?: number
}

/** One table row in the table-row unit mode. */
export interface DocxTableRow {
  /** 1-based table index in document order. */
  table: number
  /** 1-based row index inside that table. */
  row: number
  /** Cell texts joined by tab ("列1\t列2"). */
  text: string
  chars: number
  /** Cell texts, for callers that want the columns separately. */
  cells: string[]
  /** True for the first row of the table (usually the header row). */
  header: boolean
}

/** Revision markers found in the package. */
export interface DocxRevisions {
  /** `w:ins` insertions. */
  insertions: number
  /** `w:del` deletions. */
  deletions: number
  /** Formatting changes (`w:rPrChange`, `w:pPrChange`, …). */
  formatChanges: number
  /** Parts that carry at least one marker. */
  parts: string[]
  /** True when any marker exists. */
  present: boolean
}

/** Parsed .docx document. */
export interface DocxDocument {
  format: 'docx'
  paragraphs: DocxParagraph[]
  /** Non-body parts, in a stable order (headers, footers, footnotes, endnotes, comments). */
  parts: DocxPart[]
  /** Every table row, in document order. */
  tables: DocxTableRow[]
  /** Revision markers across the package. */
  revisions: DocxRevisions
  metadata: CoreProps
  /** Total visible characters of all paragraphs (lets the parse cache size the doc without a re-walk). */
  charCount: number
}

/**
 * Parse a .docx document into numbered paragraphs, non-body parts, table rows and revision counts.
 *
 * Every top-level paragraph becomes a unit, empty ones included: skipping them (as this reader
 * used to) makes `段落 N` disagree with the editor's paragraph index, so a model that reads a
 * document and then edits "paragraph 5" rewrites a different paragraph.
 * @param bytes - raw document bytes.
 * @throws an Error with a user-facing Chinese message when not a valid docx.
 */
export function readDocx(bytes: Uint8Array): DocxDocument {
  const pkg = openOoxml(bytes)
  const documentXml = pkg.text('word/document.xml')
  if (documentXml === null) throw new Error('不是有效的 .docx 文档（缺少 word/document.xml）')
  const paragraphs: DocxParagraph[] = splitWordParagraphs(documentXml).map((span, i) => ({
    index: i + 1,
    level: wordHeadingLevel(span.xml),
    text: span.text,
    chars: span.text.length,
    empty: span.empty,
  }))
  return {
    format: 'docx',
    paragraphs,
    parts: readParts(pkg),
    tables: readTables(documentXml),
    revisions: countRevisions(pkg, documentXml),
    metadata: readCoreProps(pkg),
    charCount: paragraphs.reduce((total, p) => total + p.chars, 0),
  }
}

/** Number of empty (text-free) paragraphs, for reader summaries. */
export function countEmptyParagraphs(paragraphs: DocxParagraph[]): number {
  return paragraphs.reduce((total, p) => total + (p.empty ? 1 : 0), 0)
}

/** Text of one part's paragraphs, joined by newline; empty paragraphs drop out. */
function partText(xml: string): string {
  return splitWordParagraphs(xml)
    .map((span) => span.text)
    .filter((text) => text !== '')
    .join('\n')
}

/**
 * Read every non-body part the package carries.
 *
 * Footnotes and endnotes are split per `w:footnote`/`w:endnote`: Word's separator entries
 * (`w:id="-1"`/`"0"`) are structural, not content, so they are skipped. Comments keep their author
 * and date, and are numbered by `w:id` so the ids match what `office_edit`'s `add_comment` writes.
 */
function readParts(pkg: OoxmlPackage): DocxPart[] {
  const out: DocxPart[] = []
  const names = Object.keys(pkg.files).sort()
  for (const kind of ['header', 'footer'] as const) {
    const parts = names.filter((name) => new RegExp(`^word/${kind}\\d+\\.xml$`, 'u').test(name))
    parts.forEach((name, i) => {
      const text = partText(pkg.text(name) as string)
      out.push({ kind, part: name, label: `${kind === 'header' ? '页眉' : '页脚'} ${i + 1}`, text, chars: text.length })
    })
  }
  for (const [kind, file, label] of [['footnote', 'word/footnotes.xml', '脚注'], ['endnote', 'word/endnotes.xml', '尾注']] as const) {
    const xml = pkg.text(file)
    if (xml === null) continue
    for (const note of splitElements(xml, kind)) {
      const id = Number(attribute(note.xml, 'w:id') ?? 'NaN')
      if (id === -1 || id === 0 || !Number.isFinite(id)) continue
      const text = partText(note.xml)
      out.push({ kind, part: file, label: `${label} ${id}`, text, chars: text.length, id })
    }
  }
  const commentsXml = pkg.text('word/comments.xml')
  if (commentsXml !== null) {
    for (const comment of splitElements(commentsXml, 'comment')) {
      const id = Number(attribute(comment.xml, 'w:id') ?? 'NaN')
      const author = attribute(comment.xml, 'w:author')
      const date = attribute(comment.xml, 'w:date')
      const text = partText(comment.xml)
      out.push({
        kind: 'comment',
        part: 'word/comments.xml',
        label: `批注 ${Number.isFinite(id) ? id : out.length}`,
        text,
        chars: text.length,
        ...(Number.isFinite(id) ? { id } : {}),
        ...(author !== undefined && author !== '' ? { author } : {}),
        ...(date !== undefined && date !== '' ? { date } : {}),
      })
    }
  }
  return out
}

/**
 * Flatten every top-level table into row units.
 *
 * Cells are joined by a tab (the same separator `office_read` uses for xlsx rows), so a table read
 * in this mode lines up column-for-column with the spreadsheet reading of the same data. A nested
 * table's rows belong to the inner table and are not repeated here.
 */
function readTables(documentXml: string): DocxTableRow[] {
  const out: DocxTableRow[] = []
  splitElements(documentXml, 'tbl').forEach((table, tableIndex) => {
    splitElements(table.xml, 'tr').forEach((row, rowIndex) => {
      const cells = splitElements(row.xml, 'tc').map((cell) => partText(cell.xml).replace(/\t/gu, ' ').replace(/\r?\n/gu, ' '))
      const text = cells.join('\t')
      out.push({ table: tableIndex + 1, row: rowIndex + 1, text, chars: text.length, cells, header: rowIndex === 0 })
    })
  })
  return out
}

/** Count revision markers across every content part. */
function countRevisions(pkg: OoxmlPackage, documentXml: string): DocxRevisions {
  let insertions = 0
  let deletions = 0
  let formatChanges = 0
  const parts: string[] = []
  for (const name of Object.keys(pkg.files)) {
    if (!/\.xml$/u.test(name) || name.startsWith('docProps/') || name === '[Content_Types].xml') continue
    const xml = name === 'word/document.xml' ? documentXml : pkg.text(name)
    if (xml === null) continue
    const ins = (xml.match(/<w:ins(?=[\s/>])/gu) ?? []).length
    const del = (xml.match(/<w:del(?=[\s/>])/gu) ?? []).length
    const fmt = (xml.match(/<w:(?:rPr|pPr|tblPr|trPr|tcPr|sectPr|numbering)Change(?=[\s/>])/gu) ?? []).length
    if (ins + del + fmt === 0) continue
    insertions += ins
    deletions += del
    formatChanges += fmt
    parts.push(name)
  }
  return { insertions, deletions, formatChanges, parts, present: insertions + deletions + formatChanges > 0 }
}
