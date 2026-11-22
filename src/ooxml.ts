/**
 * OOXML 基础层：zip 容器解包（fflate）与轻量 XML 文本提取工具。
 * @module dsh-office-reader/src/ooxml
 */

import { unzipSync, strFromU8 } from 'fflate'

/** An unpacked OOXML (zip) package with convenient entry access. */
export interface OoxmlPackage {
  /** Zip entries keyed by their package path. */
  files: Record<string, Uint8Array>
  /** Whether an entry exists. */
  has(name: string): boolean
  /** Entry decoded as UTF-8 text, or null when absent. */
  text(name: string): string | null
}

/**
 * Describe why a file could not be opened as an Office Open XML container.
 *
 * A zero-byte file is the most common cause of an unzip failure, and it is neither a legacy nor an
 * encrypted format — blaming those sends the reader looking in the wrong place. Every entry point
 * builds its message here so the separate copies cannot drift apart.
 *
 * @param bytes - the raw file bytes that failed to parse.
 * @param verb - the entry point's verb, e.g. `无法审计`.
 * @returns the user-facing message.
 */
export function notOoxmlMessage(bytes: Uint8Array, verb: string): string {
  if (bytes.byteLength === 0) {
    return `${verb}：文件是 0 字节的空文件，没有任何内容可解析。它多半是新建后从未保存过，请在 Office 中打开并保存后再试`
  }
  return `${verb}：不是有效的 Office Open XML（zip）容器。旧版二进制格式（.doc/.ppt/.xls）与加密文档请先在 Office 中另存为 .docx/.xlsx/.pptx`
}

/**
 * Unpack an Office Open XML container.
 * @param bytes - raw document bytes.
 * @throws an Error with a user-facing Chinese message when not a valid zip container.
 */
export function openOoxml(bytes: Uint8Array): OoxmlPackage {
  let files: Record<string, Uint8Array>
  try {
    files = unzipSync(bytes)
  } catch {
    throw new Error(notOoxmlMessage(bytes, '无法解析文档'))
  }
  return {
    files,
    has: (name) => files[name] !== undefined,
    text: (name) => {
      const entry = files[name]
      return entry === undefined ? null : strFromU8(entry)
    },
  }
}

/** Dublin Core metadata from docProps/core.xml（全部可选）。 */
export interface CoreProps {
  title?: string
  subject?: string
  keywords?: string
  creator?: string
  lastModifiedBy?: string
  created?: string
  modified?: string
}

/** Keep only defined metadata fields. */
export function cleanCoreProps(props: CoreProps): Record<string, string> {
  const out: Record<string, string> = {}
  for (const [key, value] of Object.entries(props)) {
    if (typeof value === 'string' && value !== '') out[key] = value
  }
  return out
}

/**
 * Extract Dublin Core metadata from docProps/core.xml.
 * @param pkg - unpacked package.
 */
export function readCoreProps(pkg: OoxmlPackage): CoreProps {
  const core = pkg.text('docProps/core.xml')
  if (core === null) return {}
  const firstTag = (local: string): string | undefined => {
    const match = new RegExp(`<(?:[\\w.]+:)?${local}(?:\\s[^>]*)?>([\\s\\S]*?)</(?:[\\w.]+:)?${local}>`, 'u').exec(core)
    if (match === null) return undefined
    const value = decodeXmlEntities(match[1]).trim()
    return value === '' ? undefined : value
  }
  return {
    title: firstTag('title'),
    subject: firstTag('subject'),
    keywords: firstTag('keywords'),
    creator: firstTag('creator'),
    lastModifiedBy: firstTag('lastModifiedBy'),
    created: firstTag('created'),
    modified: firstTag('modified'),
  }
}

/** Decode the five XML predefined entities plus numeric character references. */
export function decodeXmlEntities(input: string): string {
  return input
    .replace(/&#x([0-9a-fA-F]+);/gu, (_, hex: string) => String.fromCodePoint(Number.parseInt(hex, 16)))
    .replace(/&#(\d+);/gu, (_, dec: string) => String.fromCodePoint(Number(dec)))
    .replace(/&lt;/gu, '<')
    .replace(/&gt;/gu, '>')
    .replace(/&quot;/gu, '"')
    .replace(/&apos;/gu, "'")
    .replace(/&amp;/gu, '&')
}

/**
 * Concatenate the text of all `<tag>` runs inside an XML fragment, decoding entities and
 * resolving element-valued breaks (`<w:tab/>` → `\t`, `<w:br/>` → `\n`) **in document order**.
 *
 * Order matters: <w:t>制表</w:t><w:tab/><w:t>符</w:t> is `制表\t符`, which a "replace breaks
 * first, then collect text" pass would silently turn into `制表符`.
 *
 * @param xml - XML fragment.
 * @param tag - run tag local name (e.g. `w:t`, `a:t`).
 * @param breaks - optional element→character replacements.
 */
export function extractRunsText(xml: string, tag: string, breaks?: Array<[RegExp, string]>): string {
  const breakPatterns = breaks ?? []
  const re = new RegExp(`<${tag}(?:\\s[^>]*)?/>|<${tag}(?:\\s[^>]*)?>([\\s\\S]*?)</${tag}>|</?[A-Za-z_][\\w.:-]*(?:\\s[^>]*)?/?>`, 'gu')
  let out = ''
  for (let m = re.exec(xml); m !== null; m = re.exec(xml)) {
    if (m[1] !== undefined) {
      out += decodeXmlEntities(m[1])
      continue
    }
    // A bare element token: only the registered breaks contribute a character.
    for (const [pattern, replacement] of breakPatterns) {
      pattern.lastIndex = 0
      if (pattern.test(m[0])) {
        out += replacement
        break
      }
    }
  }
  return out
}

/** Text contributions inside a Word paragraph: tabs and explicit breaks. */
export const WORD_BREAKS: Array<[RegExp, string]> = [
  [/^<w:tab\b/u, '\t'],
  [/^<w:br\b/u, '\n'],
  [/^<w:cr\b/u, '\n'],
]

/**
 * One top-level Word paragraph: its exact XML slice plus the offsets it occupies in the
 * parent document, so callers can slice and splice without re-scanning.
 */
export interface WordParagraphSpan {
  /** Offset of `<w:p …>` (or `<w:p/>`) in the parent XML. */
  start: number
  /** End offset (exclusive) of `</w:p>` (or of the self-closing tag). */
  end: number
  /** Paragraph XML, i.e. `parent.slice(start, end)`. */
  xml: string
  /** Visible text: `<w:t>` content with tabs and breaks resolved and entities decoded. */
  text: string
  /** True when the paragraph carries no visible text (a page-break-only or blank paragraph). */
  empty: boolean
}

/**
 * Split WordprocessingML into top-level `<w:p>` paragraphs.
 *
 * A paragraph nested inside another paragraph (a textbox drawing inside a run) is *not* a
 * separate unit: Word addresses it through its host paragraph, and a flattened list would make
 * the "paragraph N" numbering shared with dsh-office-editor ambiguous. Paragraphs inside table
 * cells are siblings of body paragraphs at this depth, so they stay addressable.
 *
 * Nesting is resolved by a real element walk: `<w:p>` also occurs inside table cells, whose
 * closers interleave with the host paragraph's in token order, so a depth counter cannot tell
 * "the paragraph ended" from "an inner box ended".
 *
 * This is the single traversal contract for every `段落 N` label the reader emits and every
 * `index` the editor accepts. `dsh-office-editor/src/docx.ts` mirrors it; the two plugins are
 * independent (no build-time link between them), so `tests/paragraph-contract.mjs` asserts both
 * implementations agree on the same bytes.
 *
 * @param xml - word/document.xml content.
 * @throws an Error when the fragment is not balanced.
 */
export function splitWordParagraphs(xml: string): WordParagraphSpan[] {
  return splitElements(xml, 'p').map((span) => makeSpan(xml, span.start, span.end))
}

/** One XML element slice found by {@link splitElements}. */
export interface ElementSpan {
  /** Offset of the open tag in the parent XML. */
  start: number
  /** End offset (exclusive) of the close tag (or of the self-closing tag). */
  end: number
  /** The element XML, i.e. `parent.slice(start, end)`. */
  xml: string
}

/**
 * Split an XML fragment into the spans of one element name, skipping elements of that name nested
 * inside another one.
 *
 * This is the single traversal rule behind both the paragraph contract (`w:p`, where nested
 * textbox paragraphs must not become units) and the table/row/cell extraction added for table-row
 * reading (`w:tbl`/`w:tr`/`w:tc`, where a nested table's rows belong to the inner table).
 *
 * Nesting is resolved by a real element walk rather than a depth counter: `<w:p>` also occurs
 * inside table cells, whose closers interleave with the host paragraph's in token order.
 *
 * @param xml - the fragment to walk.
 * @param local - local element name to collect (namespace prefix ignored).
 * @throws an Error when the fragment is not balanced.
 */
export function splitElements(xml: string, local: string): ElementSpan[] {
  const out: ElementSpan[] = []
  const stack: string[] = []
  let cursor = 0
  const tokenRe = /<([A-Za-z_][\w.:-]*)((?:"[^"]*"|'[^']*'|[^>"'])*?)(\/?)>|<\/([A-Za-z_][\w.:-]*)\s*>|<\?[\s\S]*?\?>|<!--[\s\S]*?-->|<!\[CDATA\[[\s\S]*?\]\]>|<!DOCTYPE[^>]*>/gu
  for (let m = tokenRe.exec(xml); m !== null; m = tokenRe.exec(xml)) {
    if (m[1] !== undefined) {
      const element = localName(m[1])
      if (m[3] === '/') {
        if (element === local && !stack.includes(local)) out.push({ start: m.index, end: m.index + m[0].length, xml: xml.slice(m.index, m.index + m[0].length) })
        continue
      }
      if (element === local && !stack.includes(local)) cursor = m.index
      stack.push(element)
      continue
    }
    if (m[4] === undefined) continue
    const element = localName(m[4])
    const open = stack.pop()
    if (open === undefined) throw new Error(`XML 结构不闭合（出现多余的 </${element}>）`)
    if (open !== element) throw new Error(`XML 标签不匹配（<${open}> 被 </${element}> 关闭）`)
    if (element === local && !stack.includes(local)) out.push({ start: cursor, end: m.index + m[0].length, xml: xml.slice(cursor, m.index + m[0].length) })
  }
  if (stack.length > 0) throw new Error(`XML 结构不闭合（<${stack[stack.length - 1]}> 未闭合）`)
  return out
}

/** Local name of a possibly namespace-prefixed element name. */
export function localName(name: string): string {
  const at = name.indexOf(':')
  return at < 0 ? name : name.slice(at + 1)
}

/** One relationship of a `.rels` part. */
export interface Relationship {
  id: string
  type: string
  target: string
  /** True for `TargetMode="External"` relationships (hyperlinks, external links). */
  external: boolean
}

/**
 * Parse a relationships part into id → relationship.
 *
 * Used instead of SheetJS's own hyperlink parsing, which decodes non-ASCII relationship targets as
 * latin1 and turns `https://例子.测试/甲` into mojibake.
 */
export function parseRels(relsXml: string | null): Map<string, Relationship> {
  const out = new Map<string, Relationship>()
  if (relsXml === null) return out
  for (const match of relsXml.matchAll(/<Relationship\b[^>]*\/>/gu)) {
    const tag = match[0]
    const id = /Id="([^"]*)"/u.exec(tag)?.[1]
    const target = /Target="([^"]*)"/u.exec(tag)?.[1]
    if (id === undefined || target === undefined) continue
    out.set(id, {
      id,
      type: /Type="([^"]*)"/u.exec(tag)?.[1] ?? '',
      target: decodeXmlEntities(target),
      external: /TargetMode="External"/u.test(tag),
    })
  }
  return out
}

/** Attribute value of an element's open tag, entity-decoded. */
export function attribute(openTag: string, name: string): string | undefined {
  const match = new RegExp(`${name}="([^"]*)"`, 'u').exec(openTag)
  return match === null ? undefined : decodeXmlEntities(match[1])
}

/**
 * Resolve a relationship target against the part that declares it.
 *
 * OOXML targets are relative to the declaring part's directory (`../comments1.xml` from
 * `xl/worksheets/sheet1.xml` means `xl/comments1.xml`), may be package-absolute (`/xl/...`), and may
 * carry a fragment. Naive concatenation produced false "missing part" reports for exactly the
 * commentary relationships a sheet declares.
 *
 * @param partName - the declaring part, e.g. `xl/worksheets/sheet1.xml`.
 * @param target - the raw `Target` attribute.
 * @returns the package-relative part name, or null for external targets.
 */
export function resolveRelTarget(partName: string, target: string): string | null {
  if (/^[a-z][a-z0-9+.-]*:/iu.test(target)) return null
  const withoutFragment = target.split('#')[0]
  if (withoutFragment === '') return null
  if (withoutFragment.startsWith('/')) return normalizeSegments(withoutFragment.slice(1))
  const baseDir = partName.includes('/') ? partName.slice(0, partName.lastIndexOf('/')) : ''
  return normalizeSegments(`${baseDir}/${withoutFragment}`)
}

/** Collapse `.`/`..` segments of a package path. */
function normalizeSegments(path: string): string {
  const out: string[] = []
  for (const segment of path.split('/')) {
    if (segment === '' || segment === '.') continue
    if (segment === '..') out.pop()
    else out.push(segment)
  }
  return out.join('/')
}

/**
 * A tag-stack well-formedness check for OOXML parts.
 *
 * Deliberately not a general XML parser: it verifies that elements close in order and that nothing
 * sits outside the root, which is the damage class string-surgery editing produces.
 * @param xml - one XML part.
 * @param partName - part name used in the error message.
 * @returns the number of elements walked, or throws with the imbalance named.
 */
export function assertWellFormed(xml: string, partName: string): number {
  const cleaned = xml
    .replace(/<!--[\s\S]*?-->/gu, ' ')
    .replace(/<!\[CDATA\[[\s\S]*?\]\]>/gu, ' ')
    .replace(/<!DOCTYPE[^>]*>/gu, ' ')
  const tokenRe = /<([A-Za-z_][\w.:-]*)((?:"[^"]*"|'[^']*'|[^>"'])*?)(\/?)>|<\/([A-Za-z_][\w.:-]*)\s*>|<\?[\s\S]*?\?>/gu
  const stack: string[] = []
  let rootClosed = false
  let cursor = 0
  let elements = 0
  for (let m = tokenRe.exec(cleaned); m !== null; m = tokenRe.exec(cleaned)) {
    const stray = cleaned.slice(cursor, m.index)
    cursor = m.index + m[0].length
    if (rootClosed) {
      if (m[1] !== undefined || m[4] !== undefined) throw new Error(`${partName} 在根元素之后还有元素（偏移 ${m.index}）`)
      continue
    }
    if (stack.length === 0) {
      if (stray.trim() !== '') throw new Error(`${partName} 在根元素之前有非法内容：${JSON.stringify(stray.slice(0, 60))}`)
      if (m[1] === undefined) continue
      elements++
      if (m[3] === '/') { rootClosed = true; continue }
      stack.push(m[1])
      continue
    }
    if (m[4] !== undefined) {
      const open = stack.pop()
      if (open === undefined) throw new Error(`${partName} 在偏移 ${m.index} 处出现多余的 </${m[4]}>`)
      if (open !== m[4]) throw new Error(`${partName} 标签不匹配：<${open}> 被 </${m[4]}> 关闭（偏移 ${m.index}）`)
      if (stack.length === 0) rootClosed = true
      continue
    }
    if (m[1] === undefined) continue
    elements++
    if (m[3] !== '/') stack.push(m[1])
  }
  const trailing = cleaned.slice(cursor).trim()
  if (trailing !== '') throw new Error(`${partName} 在根元素之后有非法内容：${JSON.stringify(trailing.slice(0, 60))}`)
  if (stack.length > 0) throw new Error(`${partName} 有未闭合的标签：<${stack[stack.length - 1]}>（共 ${stack.length} 个）`)
  return elements
}


function makeSpan(xml: string, start: number, end: number): WordParagraphSpan {
  const slice = xml.slice(start, end)
  const text = extractRunsText(slice, 'w:t', WORD_BREAKS).replace(/[ \t]*\n[ \t\n]*/gu, '\n').trim()
  return { start, end, xml: slice, text, empty: text === '' }
}

/**
 * Heading level of a paragraph (1-9), from the style's outline level or its `w:pStyle`.
 * The numeric `styleId` fallback only fires for Word's own localized heading ids
 * (`Heading1`/`heading 1`/`标题 1`), never for an arbitrary numeric style id.
 * @param paragraphXml - one `<w:p>` slice.
 */
export function wordHeadingLevel(paragraphXml: string): number | undefined {
  const outline = /<w:outlineLvl\b[^>]*w:val="(\d+)"/u.exec(paragraphXml)
  if (outline !== null) return Number(outline[1]) + 1
  const style = /<w:pStyle\b[^>]*w:val="([^"]*)"/u.exec(paragraphXml)
  if (style === null) return undefined
  // `Heading1`…`Heading9`, `heading 1`, `标题 1` — Word's own heading style ids only.
  // An arbitrary numeric styleId is *not* a heading level.
  const named = /^(?:Heading|heading|标题)\s*([1-9])$/u.exec(style[1])
  return named === null ? undefined : Number(named[1])
}
