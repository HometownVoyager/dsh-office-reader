/**
 * 统一文档模型：格式识别、单元（unit）构建与读取/搜索范围选择。
 * 单元粒度：docx = 段落（另有 `parts`/`tables` 两个旁路模式）；xlsx = 工作表数据行
 * （另有 `formulas` 公式检视模式）；pptx = 幻灯片。
 * @module dsh-office-reader/src/office
 */

import { notOoxmlMessage, openOoxml, cleanCoreProps, type CoreProps, type OoxmlPackage } from './ooxml.js'
import { readDocx, type DocxParagraph, type DocxPart, type DocxTableRow, type DocxRevisions } from './docx.js'
import { readPptx, type PptxSlide } from './pptx.js'
import { readXlsx, readXlsxRangeRows, type XlsxSheetInfo, type XlsxRow, type XlsxDocument } from './xlsx.js'
import { normalizeForMatch } from './search.js'

export type OfficeFormat = 'docx' | 'xlsx' | 'pptx'

/** One readable/searchable unit with a user-facing label. */
export interface OfficeUnit {
  label: string
  text: string
  chars: number
  /** True for a structural but text-free unit (a blank/page-break docx paragraph). */
  empty?: boolean
}

/** Outline entry pointing at a unit. */
export interface OutlineEntry {
  /** Unit label, e.g. "段落 3" / "Sheet1" / "幻灯片 2". */
  unit: string
  /** 1-based index into the format's own unit list. */
  index: number
  /** Heading level (docx headings); 1 for flat lists. */
  level: number
  /** Title text. */
  title: string
}

/** xlsx-specific data kept for sheet/range/formula selection. */
export interface XlsxData {
  workbook: unknown
  sheets: XlsxSheetInfo[]
  /** Materialize every data row (cached by the parser). */
  rows: () => XlsxRow[]
  formulas: XlsxDocument['formulas']
  formulasTruncated: boolean
  definedNames: XlsxDocument['definedNames']
  comments: XlsxDocument['comments']
  hyperlinks: XlsxDocument['hyperlinks']
}

/** Unified parsed document. */
export interface OfficeDocument {
  format: OfficeFormat
  displayName: string
  metadata: CoreProps
  /**
   * Read/search granularity units, in document order.
   *
   * For xlsx this is a lazy getter: `office_inspect` never touches it, so a large workbook is not
   * materialized into row strings just to print a table of contents.
   */
  units: OfficeUnit[]
  outline: OutlineEntry[]
  /** Total unit characters, available without materializing `units` (used by the parse cache). */
  charCount: number
  paragraphs?: DocxParagraph[]
  /** docx non-body parts (headers/footers/footnotes/endnotes/comments). */
  parts?: DocxPart[]
  /** docx table rows, for the table-row unit mode. */
  tables?: DocxTableRow[]
  /** docx revision markers. */
  revisions?: DocxRevisions
  slides?: PptxSlide[]
  xlsx?: XlsxData
}

/**
 * Open a document by extension, falling back to zip content sniffing.
 * .xls/.xlsm ride the SheetJS path alongside .xlsx.
 * @throws an Error with a user-facing Chinese message for unknown/legacy/encrypted files.
 */
export function openDocument(bytes: Uint8Array, displayName: string): OfficeDocument {
  const ext = extensionOf(displayName)
  if (ext === 'xlsx' || ext === 'xlsm' || ext === 'xls') return fromXlsx(readXlsx(bytes, zipOrNull(bytes, ext)), displayName)
  if (ext === 'docx') return fromDocx(readDocx(bytes), displayName)
  if (ext === 'pptx') return fromPptx(readPptx(bytes), displayName)
  const sniffed = sniffFormat(bytes)
  if (sniffed === 'docx') return fromDocx(readDocx(bytes), displayName)
  if (sniffed === 'pptx') return fromPptx(readPptx(bytes), displayName)
  return fromXlsx(readXlsx(bytes, zipOrNull(bytes, 'xlsx')), displayName)
}

/** Unpack an OOXML container when the format is a zip; legacy CFB files have no parts to read. */
function zipOrNull(bytes: Uint8Array, ext: string): OoxmlPackage | null {
  if (ext === 'xls') return null
  if (!(bytes[0] === 0x50 && bytes[1] === 0x4b)) return null
  try {
    return openOoxml(bytes)
  } catch {
    return null
  }
}

function extensionOf(displayName: string): string {
  const base = displayName.split(/[?#]/u)[0]
  return base.split('.').at(-1)?.toLowerCase() ?? ''
}

function sniffFormat(bytes: Uint8Array): OfficeFormat {
  if (!(bytes[0] === 0x50 && bytes[1] === 0x4b)) {
    throw new Error(notOoxmlMessage(bytes, '无法识别的文档'))
  }
  const pkg = openOoxml(bytes)
  if (pkg.has('word/document.xml')) return 'docx'
  if (pkg.has('xl/workbook.xml')) return 'xlsx'
  if (pkg.has('ppt/presentation.xml')) return 'pptx'
  throw new Error('无法识别的 Office 文档：缺少 word/document.xml / xl/workbook.xml / ppt/presentation.xml 核心部件')
}

function fromDocx(parsed: ReturnType<typeof readDocx>, displayName: string): OfficeDocument {
  // Every top-level paragraph is a unit, empty ones included, so `段落 N` is exactly the
  // `index` that office_edit accepts (see dsh-office-reader/src/docx.ts).
  const units: OfficeUnit[] = parsed.paragraphs.map((p) => ({ label: `段落 ${p.index}`, text: p.text, chars: p.chars, empty: p.empty }))
  const outline: OutlineEntry[] = parsed.paragraphs
    .filter((p) => p.level !== undefined)
    .map((p) => ({ unit: `段落 ${p.index}`, index: p.index, level: p.level ?? 1, title: p.text.slice(0, 80) }))
  return {
    format: 'docx',
    displayName,
    metadata: parsed.metadata,
    units,
    outline,
    charCount: parsed.charCount,
    paragraphs: parsed.paragraphs,
    parts: parsed.parts,
    tables: parsed.tables,
    revisions: parsed.revisions,
  }
}

function fromPptx(parsed: ReturnType<typeof readPptx>, displayName: string): OfficeDocument {
  const units: OfficeUnit[] = parsed.slides.map((s) => ({ label: `幻灯片 ${s.index}`, text: s.text, chars: s.chars }))
  const outline: OutlineEntry[] = parsed.slides.map((s) => ({
    unit: `幻灯片 ${s.index}`,
    index: s.index,
    level: 1,
    title: s.title === '' ? '（无标题）' : s.title,
  }))
  return {
    format: 'pptx',
    displayName,
    metadata: parsed.metadata,
    units,
    outline,
    charCount: units.reduce((total, unit) => total + unit.chars, 0),
    slides: parsed.slides,
  }
}

function fromXlsx(parsed: XlsxDocument, displayName: string): OfficeDocument {
  const outline: OutlineEntry[] = parsed.sheets.map((s, i) => ({
    unit: s.name,
    index: i + 1,
    level: 1,
    title: `${s.name}（${s.rowCount} 行数据 × ${s.colCount} 列，使用区域 ${s.range}${s.hidden ? '，隐藏' : ''}）`,
  }))
  let cache: OfficeUnit[] | null = null
  const doc: OfficeDocument = {
    format: 'xlsx',
    displayName,
    metadata: parsed.metadata,
    outline,
    charCount: parsed.charCount,
    xlsx: {
      workbook: parsed.workbook,
      sheets: parsed.sheets,
      rows: parsed.rows,
      formulas: parsed.formulas,
      formulasTruncated: parsed.formulasTruncated,
      definedNames: parsed.definedNames,
      comments: parsed.comments,
      hyperlinks: parsed.hyperlinks,
    },
    // Lazy: `office_inspect` reads the aggregates and never materializes a single row string.
    get units(): OfficeUnit[] {
      cache ??= parsed.rows().map(rowUnit)
      return cache
    },
  }
  return doc
}

/** One xlsx data row as a read/search unit (shared by the full and per-sheet selections). */
function rowUnit(row: XlsxRow): OfficeUnit {
  return { label: `${row.sheet}!第${row.row}行`, text: row.text, chars: row.chars }
}

// ---------- unit selection ----------

/** Kinds `parts` can select from a docx package. */
export const DOCX_PART_SELECTORS = ['all', 'headers', 'footers', 'footnotes', 'endnotes', 'comments'] as const
export type DocxPartSelector = typeof DOCX_PART_SELECTORS[number]

/** Model-supplied selectors (validated against the document format). */
export interface SelectorArgs {
  sheet?: string
  range?: string
  slides?: string
  paragraphs?: string
  section?: string
  /** docx: read the non-body parts instead of the paragraphs. */
  parts?: DocxPartSelector | ''
  /** docx: read table rows instead of the paragraphs. */
  tables?: boolean
  /** xlsx: read formula cells instead of data rows. */
  formulas?: boolean
}

/** A validated selection of units. */
export interface UnitSelection {
  mode: 'full' | 'sheet' | 'range' | 'slides' | 'paragraphs' | 'section' | 'parts' | 'tables' | 'formulas'
  requested: string
  units: OfficeUnit[]
}

/** Select units for read/find from format-specific selector args. */
export function selectUnits(doc: OfficeDocument, sel: SelectorArgs, maxUnits: number): UnitSelection {
  const provided = new Set(
    Object.entries(sel)
      .filter(([, value]) => value !== undefined && value !== false && String(value).trim() !== '')
      .map(([key]) => key),
  )
  const pick = (name: keyof SelectorArgs): string => String(sel[name] ?? '').trim()
  switch (doc.format) {
    case 'docx': {
      for (const bad of ['sheet', 'range', 'slides', 'formulas'] as const) {
        if (provided.has(bad)) throw new Error(`docx 文档不支持 ${bad} 选择器（可用：section、paragraphs、parts、tables）`)
      }
      if (provided.has('parts') && provided.has('tables')) throw new Error('parts 与 tables 不能同时指定')
      const exclusive = ['section', 'paragraphs'].filter((name) => provided.has(name))
      if (exclusive.length > 1 || (exclusive.length === 1 && (provided.has('parts') || provided.has('tables')))) {
        throw new Error('section、paragraphs、parts、tables 四者互斥，请只指定一个')
      }
      if (provided.has('parts')) return selectParts(doc, pick('parts') as DocxPartSelector, maxUnits)
      if (provided.has('tables')) return selectTables(doc, maxUnits)
      if (provided.has('section')) return selectSection(doc, pick('section'), maxUnits)
      if (provided.has('paragraphs')) {
        const [first, last] = parseIndexRange(pick('paragraphs'), doc.units.length, '段落')
        return rangeSelection(doc.units, first, last, 'paragraphs', pick('paragraphs'), maxUnits)
      }
      return fullSelection(doc, maxUnits, '请用 paragraphs="开始-结束"（如 "12-35"）、section="标题"、parts="headers|footers|footnotes|endnotes|comments|all" 或 tables=true 分段读取')
    }
    case 'pptx': {
      for (const bad of ['sheet', 'range', 'paragraphs', 'section', 'parts', 'tables', 'formulas'] as const) {
        if (provided.has(bad)) throw new Error(`pptx 文档不支持 ${bad} 选择器（可用：slides）`)
      }
      if (provided.has('slides')) {
        const [first, last] = parseIndexRange(pick('slides'), doc.units.length, '幻灯片')
        return rangeSelection(doc.units, first, last, 'slides', pick('slides'), maxUnits)
      }
      return fullSelection(doc, maxUnits, '请用 slides="开始-结束"（如 "12-35"）分段读取')
    }
    case 'xlsx': {
      for (const bad of ['slides', 'paragraphs', 'section', 'parts', 'tables'] as const) {
        if (provided.has(bad)) throw new Error(`xlsx 文档不支持 ${bad} 选择器（可用：sheet、range、formulas）`)
      }
      if (provided.has('range') && !provided.has('sheet')) throw new Error('range 需要与 sheet 一起使用（先选定工作表）')
      if (provided.has('formulas')) return selectFormulas(doc, provided.has('sheet') ? pick('sheet') : undefined, maxUnits)
      if (provided.has('sheet')) {
        const sheet = resolveSheet(doc, pick('sheet'))
        if (provided.has('range')) return readSheetRange(doc, sheet.name, pick('range'), maxUnits)
        return sheetRows(doc, sheet, maxUnits)
      }
      return fullSelection(doc, maxUnits, '请用 sheet="工作表名或序号"（配合 range="A1:C10"）分段读取，或用 formulas=true 检视公式')
    }
  }
}

function parseIndexRange(selector: string, count: number, kind: string): [number, number] {
  const match = /^(\d+)(?:-(\d+))?$/u.exec(selector)
  if (match === null) throw new Error(`${kind}范围格式无效：${selector}（应为如 12，或范围如 12-35）`)
  const first = Number(match[1])
  const last = match[2] === undefined ? first : Number(match[2])
  if (first < 1 || last < first || last > count) {
    throw new Error(`${kind}范围 ${selector} 超出文档范围（共 ${count} 个）`)
  }
  return [first, last]
}

function rangeSelection(units: OfficeUnit[], first: number, last: number, mode: UnitSelection['mode'], requested: string, maxUnits: number): UnitSelection {
  const selected = units.slice(first - 1, last)
  if (selected.length > maxUnits) {
    throw new Error(`本次选取 ${selected.length} 个单元超过单次上限 ${maxUnits}，请缩小范围分段读取`)
  }
  return { mode, requested, units: selected }
}

function fullSelection(doc: OfficeDocument, maxUnits: number, hint: string): UnitSelection {
  if (doc.units.length === 0) return { mode: 'full', requested: '全文', units: [] }
  if (doc.units.length > maxUnits) {
    throw new Error(`文档共有 ${doc.units.length} 个单元，超过单次上限 ${maxUnits}。${hint}`)
  }
  return { mode: 'full', requested: '全文', units: doc.units }
}

/** Section selection: from a heading match to just before the next heading of the same-or-higher level. */
function selectSection(doc: OfficeDocument, query: string, maxUnits: number): UnitSelection {
  const normalized = normalizeForMatch(query)
  const at = doc.outline.findIndex((entry) => {
    const title = normalizeForMatch(entry.title)
    if (title === '') return false
    return title.includes(normalized) || normalized.includes(title)
  })
  if (at === -1) {
    const available = doc.outline
      .map((entry) => `${'  '.repeat(Math.max(0, entry.level - 1))}- ${entry.title}（${entry.unit}）`)
      .join('\n')
      .slice(0, 2000)
    throw new Error(`未找到标题匹配"${query}"的章节。${available === '' ? '该文档没有标题大纲。' : `可用标题：\n${available}`}`)
  }
  const entry = doc.outline[at]
  let end = doc.units.length
  for (let i = at + 1; i < doc.outline.length; i++) {
    if (doc.outline[i].level <= entry.level) {
      end = doc.outline[i].index - 1
      break
    }
  }
  if (end < entry.index) end = entry.index
  return rangeSelection(doc.units, entry.index, end, 'section', query, maxUnits)
}

/** docx non-body parts ("页眉 1" / "脚注 3" / "批注 2"), selected by kind. */
function selectParts(doc: OfficeDocument, selector: DocxPartSelector, maxUnits: number): UnitSelection {
  const parts = doc.parts ?? []
  if (!DOCX_PART_SELECTORS.includes(selector)) {
    throw new Error(`parts 选择器无效："${selector}"（可用：${DOCX_PART_SELECTORS.join('/')}）`)
  }
  const wanted = selector === 'all' ? parts : parts.filter((part) => `${part.kind}s` === selector)
  if (wanted.length === 0) {
    const available = parts.length === 0 ? '该文档没有任何页眉页脚/脚注尾注/批注部件' : `现有：${parts.map((p) => p.label).join('、')}`
    throw new Error(`没有匹配 parts="${selector}" 的内容。${available}`)
  }
  if (wanted.length > maxUnits) {
    throw new Error(`parts="${selector}" 选中 ${wanted.length} 个单元，超过单次上限 ${maxUnits}，请缩小范围`)
  }
  const units: OfficeUnit[] = wanted.map((part) => ({
    label: part.author === undefined ? part.label : `${part.label}（${part.author}）`,
    text: part.text,
    chars: part.chars,
  }))
  return { mode: 'parts', requested: `parts=${selector}（${units.length} 个）`, units }
}

/** docx table rows, one unit per `w:tr` with tab-separated cells. */
function selectTables(doc: OfficeDocument, maxUnits: number): UnitSelection {
  const rows = doc.tables ?? []
  if (rows.length === 0) throw new Error('该文档没有表格（未找到 w:tbl）')
  if (rows.length > maxUnits) {
    const tables = new Set(rows.map((row) => row.table)).size
    throw new Error(`该文档有 ${tables} 个表格共 ${rows.length} 行，超过单次上限 ${maxUnits}；请改用 office_read 的 paragraphs/section 选择器定位后逐段读取`)
  }
  const units: OfficeUnit[] = rows.map((row) => ({ label: `表格 ${row.table} 第 ${row.row} 行`, text: row.text, chars: row.chars }))
  return { mode: 'tables', requested: `tables（${new Set(rows.map((row) => row.table)).size} 个表格，${rows.length} 行）`, units }
}

/** xlsx formula cells, one unit per formula (`Sheet!D2` → `=SUM(B2:C2)`). */
function selectFormulas(doc: OfficeDocument, sheetSelector: string | undefined, maxUnits: number): UnitSelection {
  const all = doc.xlsx?.formulas ?? []
  let wanted = all
  let requested = 'formulas=true（整簿）'
  if (sheetSelector !== undefined) {
    const sheet = resolveSheet(doc, sheetSelector)
    wanted = all.filter((cell) => cell.sheetIndex === sheet.index)
    requested = `formulas=true（工作表 ${sheet.name}）`
  }
  if (wanted.length === 0) throw new Error(`没有公式${sheetSelector === undefined ? '' : `（工作表 ${sheetSelector}）`}：该工作簿的单元格都是静态值`)
  if (wanted.length > maxUnits) {
    throw new Error(`共有 ${wanted.length} 个公式单元格，超过单次上限 ${maxUnits}；请用 sheet 选择器缩小到单个工作表`)
  }
  const units: OfficeUnit[] = wanted.map((cell) => ({ label: `${cell.sheet}!${cell.addr}`, text: `=${cell.formula}`, chars: cell.formula.length + 1 }))
  return { mode: 'formulas', requested: `${requested}，${doc.xlsx?.formulasTruncated === true ? '公式清单已达采集上限，仅含前部分' : `共 ${units.length} 个`}`, units }
}

function resolveSheet(doc: OfficeDocument, selector: string): { name: string; index: number } {
  const sheets = doc.xlsx?.sheets ?? []
  const names = sheets.map((s) => s.name).join('、')
  if (/^\d+$/u.test(selector)) {
    const index = Number(selector) - 1
    if (index < 0 || index >= sheets.length) {
      throw new Error(`工作表序号 ${selector} 超出范围（共 ${sheets.length} 个：${names}）`)
    }
    return { name: sheets[index].name, index }
  }
  const at = sheets.findIndex((s) => s.name.toLowerCase() === selector.toLowerCase())
  if (at === -1) throw new Error(`找不到工作表"${selector}"（现有：${names}）`)
  return { name: sheets[at].name, index: at }
}

function sheetRows(doc: OfficeDocument, sheet: { name: string; index: number }, maxUnits: number): UnitSelection {
  const units: OfficeUnit[] = (doc.xlsx?.rows() ?? []).filter((row) => row.sheetIndex === sheet.index).map(rowUnit)
  if (units.length === 0) throw new Error(`工作表"${sheet.name}"没有可读取的文本内容`)
  if (units.length > maxUnits) {
    throw new Error(`工作表"${sheet.name}"有 ${units.length} 行数据，超过单次上限 ${maxUnits}。请用 range="A1:C10" 分段读取`)
  }
  return { mode: 'sheet', requested: sheet.name, units }
}

function readSheetRange(doc: OfficeDocument, sheetName: string, range: string, maxUnits: number): UnitSelection {
  const rows = readXlsxRangeRows(doc.xlsx?.workbook, sheetName, range)
  if (rows.length === 0) throw new Error(`区域 ${range}（工作表"${sheetName}"）没有可读取的文本内容`)
  if (rows.length > maxUnits) {
    throw new Error(`区域 ${range} 含 ${rows.length} 行，超过单次上限 ${maxUnits}，请缩小区域`)
  }
  const units = rows.map((r) => ({ label: `${sheetName}!第${r.row}行`, text: r.text, chars: r.text.length }))
  return { mode: 'range', requested: `${sheetName}!${range}`, units }
}

export { cleanCoreProps }
