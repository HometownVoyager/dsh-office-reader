/**
 * .xlsx/.xlsm/.xls 解析：基于 SheetJS，工作表枚举、行级文本重建与区域读取。
 *
 * 两条纪律：
 *
 * 1. **概览不物化整表**——一次扫描同时算出每表的行数/列数/公式数/表头/列类型/样本行/总字数，
 *    只保留有界的预览；行文本另走一次「按需物化」的扫描（`XlsxDocument.rows()`），因此
 *    `office_inspect` 不再为大工作簿付全量文本拷贝的代价（报告 §3.2.4 第 9/10 条）。
 * 2. **批注与超链接自己解包读**——SheetJS 的关系目标/作者名走 latin1 解码，非 ASCII 会变乱码
 *    （`https://example.com/甲` → `.../ç²`，`张三` → `å¼ ä¸`）。这两项直接从部件 XML 取。
 *
 * @module dsh-office-reader/src/xlsx
 */

import XLSX from 'xlsx'
import { attribute, decodeXmlEntities, openOoxml, parseRels, resolveRelTarget, type CoreProps, type OoxmlPackage } from './ooxml.js'

/** Column type tags used by the per-sheet overview. */
export type ColumnType = 'int' | 'float' | 'bool' | 'datetime' | 'str' | 'formula' | 'mixed' | 'empty'

/** Sheet summary derived from one aggregate pass over the used range. */
export interface XlsxSheetInfo {
  /** Sheet name. */
  name: string
  /** 0-based position in the workbook. */
  index: number
  /** Used range in A1 notation, e.g. "A1:C12". */
  range: string
  /** Number of non-empty data rows. */
  rowCount: number
  /** Widest row in cells. */
  colCount: number
  /** Cells carrying a formula. */
  formulaCount: number
  /** First non-empty row's cells (the header row, by convention). */
  headerCells: string[]
  /** Inferred type per column over the data rows below the header. */
  columnTypes: ColumnType[]
  /** Up to three row texts, for the overview preview. */
  sampleRows: string[]
  /** Visibility as stored in `workbook.xml` (`state` attribute). */
  visibility: 'visible' | 'hidden' | 'veryHidden'
  /** Convenience flag: `visibility !== 'visible'`. */
  hidden: boolean
  /** Legacy cell comments attached to this sheet. */
  commentCount: number
  /** Hyperlinks attached to this sheet. */
  hyperlinkCount: number
}

/** One non-empty data row flattened to text. */
export interface XlsxRow {
  /** Sheet name. */
  sheet: string
  /** 0-based sheet index. */
  sheetIndex: number
  /** 1-based sheet row number (position in the used range). */
  row: number
  /** Cell texts with trailing empties trimmed. */
  cells: string[]
  /** Cells joined by tab. */
  text: string
  /** Text length in characters. */
  chars: number
}

/** One formula cell, for formula-inspection reads. */
export interface XlsxFormulaCell {
  sheet: string
  sheetIndex: number
  /** A1 address, e.g. `D2`. */
  addr: string
  /** 1-based row of the cell, so callers can order/group by row. */
  row: number
  /** Formula text without the leading `=`. */
  formula: string
}

/** One cell comment (legacy note), read from the comments part. */
export interface XlsxComment {
  sheet: string
  sheetIndex: number
  /** A1 address the note is attached to. */
  addr: string
  author: string
  text: string
}

/** One hyperlink, read from the sheet part plus its relationships. */
export interface XlsxHyperlink {
  sheet: string
  sheetIndex: number
  addr: string
  /** External URL, or `#Sheet1!A1` for an internal link. */
  url: string
  tooltip?: string
}

/** One defined name. */
export interface XlsxDefinedName {
  name: string
  refersTo: string
  /** Sheet the name is scoped to; absent means workbook scope. */
  scope?: string
}

/** Parsed workbook. */
export interface XlsxDocument {
  format: 'xlsx'
  workbook: unknown
  sheets: XlsxSheetInfo[]
  metadata: CoreProps
  /** Workbook- and sheet-scoped defined names. */
  definedNames: XlsxDefinedName[]
  comments: XlsxComment[]
  hyperlinks: XlsxHyperlink[]
  /** Every formula cell in the workbook (bounded by Config.maxFormulas). */
  formulas: XlsxFormulaCell[]
  /** True when the formula list hit its cap. */
  formulasTruncated: boolean
  /** Total characters of all row texts, computed without materializing them. */
  charCount: number
  /** Materialize every data row (cached after the first call). */
  rows(): XlsxRow[]
}

/** Cap on collected formula cells, so a formula-heavy workbook cannot blow up the cache. */
const MAX_FORMULAS = 20_000

/**
 * Parse a workbook (.xlsx/.xlsm/.xls) via SheetJS.
 * @param bytes - raw file bytes.
 * @param pkg - unpacked OOXML package for xlsx/xlsm parts (comments/hyperlinks). Omit to let the
 *   parser unpack the container itself; pass `null` for legacy `.xls`, which has no parts.
 * @throws an Error with a user-facing Chinese message when parsing fails.
 */
export function readXlsx(bytes: Uint8Array, pkg?: OoxmlPackage | null): XlsxDocument {
  let workbook: XLSX.WorkBook
  try {
    workbook = XLSX.read(bytes, {
      type: 'buffer',
      cellDates: true,
      // Formula cells that carry no cached `<v>` are dropped without this flag, which made every
      // freshly written formula invisible to the formula-inspection mode.
      cellFormula: true,
      sheetStubs: true,
    })
  } catch (error) {
    throw new Error(`无法解析工作簿：${error instanceof Error ? error.message : String(error)}（加密工作簿不受支持）`)
  }
  const unpacked = pkg === undefined ? autoPackage(bytes) : pkg
  const parts = unpacked === null ? null : readPartSurfaces(unpacked)
  const comments = parts?.comments ?? []
  const hyperlinks = parts?.hyperlinks ?? []
  const sheets: XlsxSheetInfo[] = []
  const formulas: XlsxFormulaCell[] = []
  let formulasTruncated = false
  let charCount = 0
  workbook.SheetNames.forEach((name, index) => {
    const state = workbook.Workbook?.Sheets?.[index]?.Hidden ?? 0
    const visibility = state === 2 ? 'veryHidden' : (state === 1 ? 'hidden' : 'visible')
    const scan = scanSheet(workbook, name, index, {
      visibility,
      commentCount: comments.filter((c) => c.sheetIndex === index).length,
      hyperlinkCount: hyperlinks.filter((h) => h.sheetIndex === index).length,
      collectRows: false,
      formulaSink: (cell) => {
        if (formulas.length >= MAX_FORMULAS) { formulasTruncated = true; return }
        formulas.push(cell)
      },
    })
    charCount += scan.charCount
    sheets.push(scan.info)
  })
  let cache: XlsxRow[] | null = null
  return {
    format: 'xlsx',
    workbook,
    sheets,
    metadata: workbookProps(workbook),
    definedNames: readDefinedNames(workbook),
    comments,
    hyperlinks,
    formulas,
    formulasTruncated,
    charCount,
    rows: () => {
      cache ??= workbook.SheetNames.flatMap((name, index) => scanSheet(workbook, name, index, {
        visibility: 'visible',
        commentCount: 0,
        hyperlinkCount: 0,
        collectRows: true,
      }).rows)
      return cache
    },
  }
}

/** Per-sheet aggregate scan (and, optionally, row materialization) in a single pass. */
function scanSheet(
  workbook: XLSX.WorkBook,
  name: string,
  index: number,
  options: {
    visibility: XlsxSheetInfo['visibility']
    commentCount: number
    hyperlinkCount: number
    collectRows: boolean
    formulaSink?: (cell: XlsxFormulaCell) => void
  },
): { info: XlsxSheetInfo; rows: XlsxRow[]; charCount: number } {
  const ws = workbook.Sheets[name]
  const rows: XlsxRow[] = []
  if (ws === undefined) {
    return {
      info: { name, index, range: 'A1', rowCount: 0, colCount: 0, formulaCount: 0, headerCells: [], columnTypes: [], sampleRows: [], visibility: options.visibility, hidden: options.visibility !== 'visible', commentCount: options.commentCount, hyperlinkCount: options.hyperlinkCount },
      rows,
      charCount: 0,
    }
  }
  const ref = typeof ws['!ref'] === 'string' && /^[A-Z]+\d+(?::[A-Z]+\d+)?$/u.test(ws['!ref']) ? ws['!ref'] : 'A1'
  const range = XLSX.utils.decode_range(ref)
  const typeSets: Array<Set<ColumnType>> = []
  /** Columns that only ever hold value-less formula cells ("继承来的工作簿" often looks like this). */
  const formulaOnlyColumns = new Set<number>()
  let headerCells: string[] = []
  let rowCount = 0
  let colCount = 0
  let formulaCount = 0
  let charCount = 0
  const sampleRows: string[] = []
  for (let r = range.s.r; r <= range.e.r; r++) {
    const cells: string[] = []
    let hasContent = false
    for (let c = range.s.c; c <= range.e.c; c++) {
      const addr = XLSX.utils.encode_cell({ r, c })
      const cell = ws[addr] as XLSX.CellObject | undefined
      if (cell === undefined) { cells.push(''); continue }
      const formula = typeof cell.f === 'string' && cell.f !== '' ? cell.f : undefined
      // `t: 'z'` is a stub: an "exists but empty" cell. A *formula* cell without a cached value is
      // read as exactly that stub (with a placeholder `v: 0`), so skipping every stub would hide
      // every freshly written formula — and formatting it would show a made-up 0.
      if (cell.t === 'z' && formula === undefined) { cells.push(''); continue }
      if (formula !== undefined) {
        formulaCount++
        options.formulaSink?.({ sheet: name, sheetIndex: index, addr, row: r + 1, formula })
      }
      const text = cell.t === 'z' ? '' : formatCell(cell)
      if (text !== '') hasContent = true
      // A formula whose cached value is missing still makes the row real (Excel shows its result).
      if (formula !== undefined) hasContent = true
      cells.push(text.replace(/\t/gu, ' ').replace(/\r?\n/gu, ' '))
    }
    while (cells.length > 0 && cells.at(-1) === '') cells.pop()
    if (!hasContent || cells.length === 0) continue
    rowCount++
    colCount = Math.max(colCount, cells.length)
    const text = cells.join('\t')
    charCount += text.length
    if (headerCells.length === 0) {
      headerCells = [...cells]
    } else {
      // Types are inferred over the data rows below the header row, walking the *untrimmed* column
      // range: a stub formula cell trims away with the trailing empties, and it still has to be
      // remembered (otherwise a formula-only column reports as `empty`).
      for (let c = range.s.c; c <= range.e.c; c++) {
        const cell = ws[XLSX.utils.encode_cell({ r, c })] as XLSX.CellObject | undefined
        if (cell === undefined) continue
        if (cell.t === 'z') {
          if (typeof cell.f === 'string' && cell.f !== '') formulaOnlyColumns.add(c)
          continue
        }
        const text = formatCell(cell)
        if (text === '') continue
        typeSets[c] ??= new Set<ColumnType>()
        typeSets[c].add(cellType(cell))
      }
    }
    if (sampleRows.length < 3) sampleRows.push(text.slice(0, 200))
    if (options.collectRows) {
      rows.push({ sheet: name, sheetIndex: index, row: r + 1, cells, text, chars: text.length })
    }
  }
  const columnTypes: ColumnType[] = []
  for (let c = 0; c < colCount; c++) {
    const set = typeSets[c]
    if (set === undefined || set.size === 0) columnTypes.push(formulaOnlyColumns.has(c) ? 'formula' : 'empty')
    else if (set.size === 1) columnTypes.push([...set][0])
    else columnTypes.push('mixed')
  }
  return {
    info: {
      name, index, range: ref, rowCount, colCount, formulaCount, headerCells, columnTypes, sampleRows,
      visibility: options.visibility,
      hidden: options.visibility !== 'visible',
      commentCount: options.commentCount,
      hyperlinkCount: options.hyperlinkCount,
    },
    rows,
    charCount,
  }
}

/** Unpack the container when the bytes are a zip (xlsx/xlsm); legacy CFB files have no parts. */
function autoPackage(bytes: Uint8Array): OoxmlPackage | null {
  if (!(bytes[0] === 0x50 && bytes[1] === 0x4b)) return null
  try {
    return openOoxml(bytes)
  } catch {
    return null
  }
}

/** Display text of one cell, via SheetJS's own formatter (number formats included). */
function formatCell(cell: XLSX.CellObject): string {
  if (cell.w !== undefined && cell.w !== '') return cell.w
  try {
    const formatted = XLSX.utils.format_cell(cell)
    return formatted === undefined || formatted === null ? '' : String(formatted)
  } catch {
    return cell.v === undefined || cell.v === null ? '' : String(cell.v)
  }
}

/** Value type tag of one cell, for the overview's per-column inference. */
function cellType(cell: XLSX.CellObject | undefined): ColumnType {
  if (cell === undefined) return 'empty'
  switch (cell.t) {
    case 'n': {
      const value = typeof cell.v === 'number' ? cell.v : Number(cell.v)
      return Number.isInteger(value) ? 'int' : 'float'
    }
    case 'b': return 'bool'
    case 'd': return 'datetime'
    case 'e': return 'str'
    default: return 'str'
  }
}

/** Map SheetJS workbook props to our CoreProps shape. */
function workbookProps(workbook: XLSX.WorkBook): CoreProps {
  const props = workbook.Props
  if (props === undefined) return {}
  const out: CoreProps = {}
  if (typeof props.Title === 'string' && props.Title !== '') out.title = props.Title
  if (typeof props.Subject === 'string' && props.Subject !== '') out.subject = props.Subject
  if (typeof props.Keywords === 'string' && props.Keywords !== '') out.keywords = props.Keywords
  if (typeof props.Author === 'string' && props.Author !== '') out.creator = props.Author
  if (typeof props.LastAuthor === 'string' && props.LastAuthor !== '') out.lastModifiedBy = props.LastAuthor
  if (props.CreatedDate instanceof Date) out.created = props.CreatedDate.toISOString()
  if (props.ModifiedDate instanceof Date) out.modified = props.ModifiedDate.toISOString()
  return out
}

/** Defined names, resolved to their sheet scope. */
function readDefinedNames(workbook: XLSX.WorkBook): XlsxDefinedName[] {
  const names = workbook.Workbook?.Names ?? []
  const out: XlsxDefinedName[] = []
  for (const entry of names) {
    if (typeof entry.Name !== 'string' || entry.Name === '') continue
    const scope = typeof entry.Sheet === 'number' ? workbook.SheetNames[entry.Sheet] : undefined
    out.push({ name: entry.Name, refersTo: String(entry.Ref ?? ''), ...(scope !== undefined ? { scope } : {}) })
  }
  return out
}

/** Surfaces read straight out of the package parts: comments and hyperlinks. */
function readPartSurfaces(pkg: OoxmlPackage): { comments: XlsxComment[]; hyperlinks: XlsxHyperlink[] } {
  const files = pkg.files
  const workbookXml = pkg.text('xl/workbook.xml')
  const workbookRels = parseRels(pkg.text('xl/_rels/workbook.xml.rels'))
  const comments: XlsxComment[] = []
  const hyperlinks: XlsxHyperlink[] = []
  if (workbookXml === null) return { comments, hyperlinks }
  const sheetParts: Array<{ name: string; part: string }> = []
  for (const match of workbookXml.matchAll(/<sheet\b[^>]*\/>/gu)) {
    const name = attribute(match[0], 'name')
    const rid = attribute(match[0], 'r:id') ?? attribute(match[0], 'id')
    if (name === undefined || rid === undefined) continue
    const target = workbookRels.get(rid)?.target
    if (target === undefined) continue
    sheetParts.push({ name, part: normalizePartPath('xl', target) })
  }
  const partToIndex = new Map(sheetParts.map((entry, index) => [entry.part, index]))
  // Hyperlinks: the sheet part names a relationship id; the relationship carries the real (UTF-8)
  // target. Internal links have no relationship and use `location="#Sheet1!A1"` instead.
  sheetParts.forEach((entry, index) => {
    const sheetXml = pkg.text(entry.part)
    if (sheetXml === null) return
    const rels = parseRels(pkg.text(relsPathOf(entry.part)))
    const block = /<hyperlinks>[\s\S]*?<\/hyperlinks>/u.exec(sheetXml)
    if (block === null) return
    for (const link of block[0].matchAll(/<hyperlink\b[^>]*\/>|<hyperlink\b[^>]*>[\s\S]*?<\/hyperlink>/gu)) {
      const addr = attribute(link[0], 'ref')
      if (addr === undefined) continue
      const tooltip = attribute(link[0], 'tooltip')
      const rid = attribute(link[0], 'r:id')
      const location = attribute(link[0], 'location')
      const url = rid !== undefined ? rels.get(rid)?.target : (location !== undefined ? `#${location}` : undefined)
      if (url === undefined) continue
      hyperlinks.push({ sheet: entry.name, sheetIndex: index, addr, url, ...(tooltip !== undefined && tooltip !== '' ? { tooltip } : {}) })
    }
  })
  // Comments: legacy notes live in `xl/commentsN.xml`, referenced from the owning sheet's rels
  // (`../comments1.xml`, i.e. resolved against the sheet part's own directory).
  for (const part of Object.keys(files)) {
    if (!/^xl\/comments\d*\.xml$/u.test(part)) continue
    const owner = [...partToIndex.entries()].find(([sheetPart]) => {
      const rels = parseRels(pkg.text(relsPathOf(sheetPart)))
      return [...rels.values()].some((rel) => resolveRelTarget(sheetPart, rel.target) === part)
    })
    if (owner === undefined) continue
    const [sheetPart, index] = owner
    const sheetName = sheetParts[index]?.name ?? sheetPart
    const xml = pkg.text(part)
    if (xml === null) continue
    const authors = [...(/<authors>[\s\S]*?<\/authors>/u.exec(xml)?.[0] ?? '').matchAll(/<author>([\s\S]*?)<\/author>/gu)]
      .map((m) => decodeXmlEntities(m[1]))
    for (const comment of xml.matchAll(/<comment\b([^>]*)>([\s\S]*?)<\/comment>/gu)) {
      const addr = attribute(`<comment ${comment[1]}>`, 'ref')
      if (addr === undefined) continue
      const authorId = Number(attribute(`<comment ${comment[1]}>`, 'authorId') ?? 'NaN')
      const body = /<text>([\s\S]*?)<\/text>/u.exec(comment[2])?.[1] ?? ''
      const text = [...body.matchAll(/<t(?:\s[^>]*)?>([\s\S]*?)<\/t>/gu)].map((m) => decodeXmlEntities(m[1])).join('')
      comments.push({
        sheet: sheetName,
        sheetIndex: index,
        addr,
        author: Number.isFinite(authorId) ? (authors[authorId] ?? '') : '',
        text,
      })
    }
  }
  return { comments, hyperlinks }
}

/** `xl/worksheets/sheet1.xml` → `xl/worksheets/_rels/sheet1.xml.rels`. */
function relsPathOf(part: string): string {
  const at = part.lastIndexOf('/')
  return `${part.slice(0, at)}/_rels/${part.slice(at + 1)}.rels`
}

/** Resolve a relationship target against its base directory, keeping the package-absolute form. */
function normalizePartPath(base: string, target: string): string {
  if (target.startsWith('/')) return target.slice(1)
  if (target.startsWith('../')) return normalizePartPath(base.slice(0, base.lastIndexOf('/')), target.slice(3))
  return `${base}/${target.replace(/^\.\//u, '')}`
}

const RANGE_RE = /^[A-Za-z]{1,4}\d+(?::[A-Za-z]{1,4}\d+)?$/u

/**
 * Read a cell range from one sheet as text rows with their real row numbers.
 * @param workbook - the parsed workbook (opaque handle from {@link XlsxDocument}).
 * @param sheetName - exact sheet name.
 * @param range - A1 range like "A1:C10", or a start cell like "B2" (reads to the end).
 */
export function readXlsxRangeRows(workbook: unknown, sheetName: string, range: string): Array<{ row: number; text: string }> {
  const wb = workbook as XLSX.WorkBook
  const ws = wb.Sheets[sheetName]
  if (ws === undefined) throw new Error(`找不到工作表"${sheetName}"`)
  const trimmed = range.trim()
  if (!RANGE_RE.test(trimmed)) {
    throw new Error(`单元格区域格式无效：${range}（应为如 A1:C10，或起始单元格如 B2）`)
  }
  let matrix: unknown[][]
  let startRow: number
  try {
    const decoded = XLSX.utils.decode_range(trimmed.includes(':') ? trimmed : `${trimmed}:${trimmed}`)
    startRow = decoded.s.r + 1
    matrix = XLSX.utils.sheet_to_json(ws, { header: 1, range: trimmed, defval: '', raw: false }) as unknown[][]
  } catch (error) {
    throw new Error(`单元格区域读取失败：${range}（${error instanceof Error ? error.message : String(error)}）`)
  }
  const out: Array<{ row: number; text: string }> = []
  matrix.forEach((raw, i) => {
    const cells = raw.map((cell) => (cell === undefined || cell === null ? '' : String(cell)))
    while (cells.length > 0 && cells.at(-1) === '') cells.pop()
    if (cells.length === 0) return
    const text = cells.map((cell) => cell.replace(/\t/gu, ' ').replace(/\r?\n/gu, ' ')).join('\t')
    out.push({ row: startRow + i, text })
  })
  return out
}
