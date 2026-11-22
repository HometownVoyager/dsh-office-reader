/**
 * @dsh-external/dsh-office-reader — Office 文档查阅工具集（.docx/.xlsx/.pptx）。
 *
 * 提供三个模型可见工具，与 dsh-pdf-reader 三件套对齐：
 * - `office_inspect`：文档概览（格式、元数据、标题大纲/工作表列表/幻灯片列表与单元内容量预览）；
 * - `office_read`：正文提取——docx 全文/章节（section）/段落范围（paragraphs）；xlsx 全部/指定
 *   工作表（sheet）/单元格区域（range）；pptx 全文/幻灯片范围（slides）；超长自动截断；
 * - `office_find`：关键词搜索（NFKC + 大小写折叠 + 空白不敏感），命中单元为
 *   docx 段落 / xlsx 数据行（Sheet!第N行）/ pptx 幻灯片，附上下文片段。
 *
 * 输入统一为 `source`：本地路径（绝对或相对会话工作目录）、file:// 链接或 http(s) 网址
 * （URL 下载仅入内存，不落盘）。解析结果经进程内缓存复用（本地文件按 size+mtime 校验，
 * 命中不读盘不解析；上限见 Config.cacheMaxEntries / cacheMaxChars）。旧版二进制格式
 * （.doc/.ppt）与加密文档以明确中文错误报告（.xls/.xlsm 例外，经 SheetJS 支持）。
 *
 * 解析引擎：docx/pptx 为内置轻量 OOXML 遍历（fflate 解包 + 结构化 XML 提取），
 * xlsx/xls 为 SheetJS。
 */

import type { Context } from 'cordis'
import { defineTool, type GenericCallView, type ToolExecution } from '@deepseek-ai/dsh-tools'
import z from 'schemastery'
import { type SourceCaps } from './source.js'
import { compilePattern, normalizeForMatch, searchUnitFor, searchUnitRegex } from './search.js'
import { OfficeCache } from './cache.js'
import { selectUnits, cleanCoreProps, DOCX_PART_SELECTORS, type OfficeDocument, type SelectorArgs } from './office.js'
import { auditOffice, auditHeadline, type AuditReport } from './audit.js'

export const name = '@dsh-external/dsh-office-reader'
export const inject = ['tools']

/** Deployment caps and defaults for the Office tools (all changeable from cordis.yml). */
export interface Config {
  /** Maximum characters returned by one `office_read` call; over-length output truncates. */
  maxCharsPerCall: number
  /** Maximum units (paragraphs/rows/slides) selected by one call. */
  maxUnitsPerCall: number
  /** Maximum hits returned by one `office_find` call across units. */
  maxSearchMatches: number
  /** Maximum hits counted per unit in regex mode, bounding the work one pattern can cost. */
  maxRegexMatchesPerUnit: number
  /** Maximum length of a `regex=true` pattern, bounding compile and match cost. */
  maxRegexPatternChars: number
  /** Character length of the per-unit preview in `office_inspect`. */
  previewChars: number
  /** Context characters kept on each side of a search match snippet. */
  contextChars: number
  /** Maximum local document file size in bytes. */
  maxFileBytes: number
  /** Maximum downloaded document body size in bytes. */
  maxDownloadBytes: number
  /** Download timeout in milliseconds. */
  downloadTimeoutMs: number
  /** Regex list of allowed URL hosts; empty allows any host. */
  allowedHosts: string[]
  /** Parsed-document cache: maximum cached documents (LRU). */
  cacheMaxEntries: number
  /** Parsed-document cache: maximum total unit chars across cached documents. */
  cacheMaxChars: number
}

export const Config = z.object({
  maxCharsPerCall: z.number().default(20000),
  maxUnitsPerCall: z.number().default(200),
  maxSearchMatches: z.number().default(50),
  maxRegexMatchesPerUnit: z.number().default(1000),
  maxRegexPatternChars: z.number().default(500),
  previewChars: z.number().default(120),
  contextChars: z.number().default(60),
  maxFileBytes: z.number().default(200 * 1024 * 1024),
  maxDownloadBytes: z.number().default(100 * 1024 * 1024),
  downloadTimeoutMs: z.number().default(30000),
  allowedHosts: z.array(z.string()).default([]),
  cacheMaxEntries: z.number().default(6),
  cacheMaxChars: z.number().default(30_000_000),
})

const SOURCE_DESC = 'Office 文档来源（.docx/.xlsx/.pptx，.xls/.xlsm 亦可）：本地文件路径（绝对路径或相对会话工作目录）、file:// 链接或 http(s) 网址'

/**
 * The slice of JSON a regex capture group can be: a string, or null for a group that did not
 * participate in the match. Declared locally because the tool result channel's own `JsonValue`
 * lives in `@deepseek-ai/dsh-util-values`, which this plugin does not link against.
 */
type JsonLike = string | number | boolean | null | JsonLike[] | { [key: string]: JsonLike }

const FORMAT_LABEL: Record<OfficeDocument['format'], string> = {
  docx: 'Word 文档（docx）',
  xlsx: 'Excel 工作簿（xlsx）',
  pptx: 'PowerPoint 演示文稿（pptx）',
}

/** Open the source through the process-wide parse cache (mtime-validated for local files). */
function openFor(cache: OfficeCache, source: string, exec: ToolExecution, caps: SourceCaps): Promise<OfficeDocument> {
  return cache.open(source, exec, caps)
}

export function apply(ctx: Context, config: Config): void {
  const caps: SourceCaps = {
    maxFileBytes: config.maxFileBytes,
    maxDownloadBytes: config.maxDownloadBytes,
    downloadTimeoutMs: config.downloadTimeoutMs,
    allowedHosts: config.allowedHosts,
  }
  const cache = new OfficeCache({ maxEntries: Math.max(1, config.cacheMaxEntries), maxChars: Math.max(1, config.cacheMaxChars) })

  // ---------- office_inspect：概览 ----------
  ctx.effect(() => ctx.tools.register(defineTool({
    name: 'office_inspect',
    description: [
      '查看 Office 文档概览（.docx/.xlsx/.pptx）：格式、元数据、标题大纲（docx）/ 工作表列表（xlsx）/ 幻灯片列表（pptx）与各单元内容量预览，用于决定如何提取正文。',
      'xlsx 额外给出每表的公式数、表头、列类型推断、样本行、可见性、批注与超链接数，以及定义名清单；docx 额外给出页眉页脚/脚注尾注/批注数量、表格数与修订痕迹。',
    ].join(''),
    parameters: {
      source: { type: 'string', required: true, description: SOURCE_DESC },
    },
    output: {
      schema: {
        type: 'object',
        additionalProperties: false,
        properties: {
          source: { type: 'string', required: true },
          format: { type: 'string', required: true },
          metadata: { type: 'object', additionalProperties: true, required: true, properties: {} },
          outline: {
            type: 'array',
            required: true,
            items: {
              type: 'object',
              additionalProperties: false,
              properties: {
                unit: { type: 'string', required: true },
                level: { type: 'integer', required: true },
                title: { type: 'string', required: true },
              },
            },
          },
          units: {
            type: 'array',
            required: true,
            items: {
              type: 'object',
              additionalProperties: false,
              properties: {
                unit: { type: 'string', required: true },
                chars: { type: 'integer', required: true },
                preview: { type: 'string' },
              },
            },
          },
          units_truncated: { type: 'boolean', required: true },
          sheets: {
            type: 'array',
            required: true,
            description: 'xlsx：逐表概览（其他格式为空数组）',
            items: {
              type: 'object',
              additionalProperties: false,
              properties: {
                name: { type: 'string', required: true },
                range: { type: 'string', required: true },
                rows: { type: 'integer', required: true },
                cols: { type: 'integer', required: true },
                formulas: { type: 'integer', required: true },
                visibility: { type: 'string', required: true },
                header: { type: 'array', required: true, items: { type: 'string' } },
                column_types: { type: 'array', required: true, items: { type: 'string' } },
                comments: { type: 'integer', required: true },
                hyperlinks: { type: 'integer', required: true },
              },
            },
          },
          defined_names: {
            type: 'array',
            required: true,
            items: {
              type: 'object',
              additionalProperties: false,
              properties: {
                name: { type: 'string', required: true },
                refers_to: { type: 'string', required: true },
                scope: { type: 'string' },
              },
            },
          },
          docx: {
            type: 'object',
            additionalProperties: false,
            required: true,
            description: 'docx：正文之外的部件与修订痕迹（其他格式为空对象）',
            properties: {
              parts: {
                type: 'array',
                required: true,
                items: {
                  type: 'object',
                  additionalProperties: false,
                  properties: {
                    label: { type: 'string', required: true },
                    kind: { type: 'string', required: true },
                    chars: { type: 'integer', required: true },
                    author: { type: 'string' },
                  },
                },
              },
              table_count: { type: 'integer', required: true },
              table_rows: { type: 'integer', required: true },
              revisions: {
                type: 'object',
                additionalProperties: false,
                required: true,
                properties: {
                  present: { type: 'boolean', required: true },
                  insertions: { type: 'integer', required: true },
                  deletions: { type: 'integer', required: true },
                  format_changes: { type: 'integer', required: true },
                  parts: { type: 'string' },
                },
              },
            },
          },
        },
      },
      render: (_args, value) => [{ type: 'text', text: renderInspect(value) }],
    },
    isConcurrencySafe: () => true,
    async execute(args, exec) {
      return inspectOffice(cache, args, exec, caps, config)
    },
    presentCall: (args): GenericCallView => ({
      card: 'generic',
      title: `Inspect ${basename(args.source)}`,
      kind: 'read',
      ...isLocalPath(args.source) ? { locations: [{ path: args.source }] } : {},
    }),
  })), 'office_inspect')

  // ---------- office_read：正文提取 ----------
  ctx.effect(() => ctx.tools.register(defineTool({
    name: 'office_read',
    description: [
      '提取 Office 文档正文（.docx/.xlsx/.pptx），超长自动截断并提示如何分段继续：',
      'docx 支持全文 / 章节（section="标题"）/ 段落范围（paragraphs，如 "12-35"）/ 正文之外的部件（parts="headers|footers|footnotes|endnotes|comments|all"）/ 表格行（tables=true，一行一单元、单元格以制表符分隔）；',
      'xlsx 支持全部 / 指定工作表（sheet，名称或序号）/ 单元格区域（range，如 "A1:C10"，需与 sheet 同用）/ 公式检视（formulas=true，逐格返回 =公式，可用 sheet 缩小）；',
      'pptx 支持全文 / 幻灯片范围（slides，如 "12-35"）。',
      'docx 含未接受修订时，结果会带 revisions 计数与提示。',
    ].join(''),
    parameters: {
      source: { type: 'string', required: true, description: SOURCE_DESC },
      sheet: { type: 'string', description: 'xlsx：工作表名称（或序号，如 "2"）；与其他选择器互斥（formulas=true 时可用来限定工作表）' },
      range: { type: 'string', description: 'xlsx：单元格区域（如 "A1:C10" 或起始单元格 "B2"），需与 sheet 同用' },
      slides: { type: 'string', description: 'pptx：幻灯片范围（与 slides 二选一），如 "12-35"；省略则读取全文' },
      paragraphs: { type: 'string', description: 'docx：段落范围，如 "12-35"；省略则读取全文' },
      section: { type: 'string', description: 'docx：按标题读取章节（模糊匹配标题，读到同级下一个标题前）' },
      parts: { type: 'string', enum: [...DOCX_PART_SELECTORS], description: 'docx：读取正文之外的部件——页眉页脚（headers/footers）、脚注尾注（footnotes/endnotes）、批注（comments）或全部（all）' },
      tables: { type: 'boolean', description: 'docx：把表格按行读取（一个 w:tr 一个单元，单元格以制表符分隔），便于把表格当数据用' },
      formulas: { type: 'boolean', description: 'xlsx：公式检视模式——逐格返回 =公式（含地址），而不是数据行' },
      max_chars: { type: 'number', description: `本次返回的最大字符数（默认与上限 ${config.maxCharsPerCall}）` },
    },
    output: {
      schema: {
        type: 'object',
        additionalProperties: false,
        properties: {
          source: { type: 'string', required: true },
          format: { type: 'string', required: true },
          mode: { type: 'string', required: true },
          requested: { type: 'string', required: true },
          unit_count: { type: 'integer', required: true },
          units: {
            type: 'array',
            required: true,
            items: {
              type: 'object',
              additionalProperties: false,
              properties: {
                unit: { type: 'string', required: true },
                chars: { type: 'integer', required: true },
                text: { type: 'string', required: true },
                empty: { type: 'boolean', description: 'docx：空段落（无可见文本；分页符/空行/分节符段落）' },
              },
            },
          },
          total_chars: { type: 'integer', required: true },
          truncated: { type: 'boolean', required: true },
          revisions: {
            type: 'object',
            additionalProperties: false,
            description: 'docx：未接受的修订痕迹计数（无修订时 present=false）',
            required: true,
            properties: {
              present: { type: 'boolean', required: true },
              insertions: { type: 'integer', required: true },
              deletions: { type: 'integer', required: true },
              format_changes: { type: 'integer', required: true },
            },
          },
          chars_per_unit: {
            type: 'array',
            required: true,
            items: {
              type: 'object',
              additionalProperties: false,
              properties: {
                unit: { type: 'string', required: true },
                chars: { type: 'integer', required: true },
              },
            },
          },
        },
      },
      render: (_args, value) => [{ type: 'text', text: renderRead(value) }],
    },
    isConcurrencySafe: () => true,
    async execute(args, exec) {
      return readOffice(cache, args, exec, caps, config)
    },
    presentCall: (args): GenericCallView => {
      const selector = args.sheet ?? args.range ?? args.slides ?? args.paragraphs ?? args.section
        ?? (args.parts !== undefined ? `parts=${args.parts}` : undefined)
        ?? (args.tables === true ? 'tables' : undefined)
        ?? (args.formulas === true ? 'formulas' : undefined)
      const suffix = selector !== undefined ? ` (${selector})` : ' (full)'
      return {
        card: 'generic',
        title: `Read ${basename(args.source)}${suffix}`,
        kind: 'read',
        ...isLocalPath(args.source) ? { locations: [{ path: args.source }] } : {},
      }
    },
  })), 'office_read')

  // ---------- office_find：关键词搜索 ----------
  ctx.effect(() => ctx.tools.register(defineTool({
    name: 'office_find',
    description: [
      '在 Office 文档（.docx/.xlsx/.pptx）中搜索关键词：大小写不敏感、全半角与空白容错，返回命中单元（docx 段落 / xlsx 数据行 Sheet!第N行 / pptx 幻灯片）、每单元次数与上下文片段。可用与 office_read 相同的选择器缩小搜索范围。',
      'regex=true 时 query 按 JavaScript 正则解释，直接在原文上匹配：不再做全半角折叠与空白剥离（需要容错空白就自己写 \\s），默认忽略大小写、可用 case_sensitive=true 区分（JavaScript 没有 (?i) 内联标志，别写进模式里）。此时每个命中还会回传捕获组 captured（按序，未参与的组为 null）；正则模式按单元封顶计数，capped 时 truncated=true。',
    ].join(''),
    parameters: {
      source: { type: 'string', required: true, description: SOURCE_DESC },
      query: { type: 'string', required: true, description: '要搜索的关键词或短语；regex=true 时为正则表达式' },
      regex: { type: 'boolean', description: '把 query 当正则表达式（默认 false：按字面量做 NFKC + 大小写折叠 + 空白不敏感匹配）' },
      case_sensitive: { type: 'boolean', description: 'regex=true 时区分大小写（默认 false：忽略大小写）' },
      sheet: { type: 'string', description: 'xlsx：只搜索该工作表（名称或序号）' },
      range: { type: 'string', description: 'xlsx：只搜索该单元格区域（需与 sheet 同用）' },
      slides: { type: 'string', description: 'pptx：只搜索幻灯片范围，如 "12-35"' },
      paragraphs: { type: 'string', description: 'docx：只搜索段落范围，如 "12-35"' },
      section: { type: 'string', description: 'docx：只搜索该章节（模糊匹配标题）' },
      parts: { type: 'string', enum: [...DOCX_PART_SELECTORS], description: 'docx：只搜索正文之外的部件（页眉页脚/脚注尾注/批注）' },
      tables: { type: 'boolean', description: 'docx：只搜索表格行单元' },
      formulas: { type: 'boolean', description: 'xlsx：只搜索公式（=公式 文本）' },
      max_matches: { type: 'number', description: `返回命中数上限（默认 ${config.maxSearchMatches}）` },
    },
    output: {
      schema: {
        type: 'object',
        additionalProperties: false,
        properties: {
          source: { type: 'string', required: true },
          format: { type: 'string', required: true },
          query: { type: 'string', required: true },
          regex: { type: 'boolean', required: true },
          case_sensitive: { type: 'boolean', required: true },
          matches: {
            type: 'array',
            required: true,
            items: {
              type: 'object',
              additionalProperties: false,
              properties: {
                unit: { type: 'string', required: true },
                count: { type: 'integer', required: true },
                snippets: { type: 'array', required: true, items: { type: 'string' } },
                captured: {
                  type: 'array',
                  description: 'regex=true 时每个命中的捕获组（按序，未参与的组为 null）',
                  items: { type: 'json' },
                },
              },
            },
          },
          total_matches: { type: 'integer', required: true },
          searched_units: { type: 'integer', required: true },
          truncated: { type: 'boolean', required: true },
        },
      },
      render: (_args, value) => [{ type: 'text', text: renderFind(value) }],
    },
    isConcurrencySafe: () => true,
    async execute(args, exec) {
      return findOffice(cache, args, exec, caps, config)
    },
    presentCall: (args): GenericCallView => ({
      card: 'generic',
      title: `Find "${args.query}" in ${basename(args.source)}`,
      kind: 'search',
      ...isLocalPath(args.source) ? { locations: [{ path: args.source }] } : {},
    }),
  })), 'office_find')

  // ---------- office_audit：良构性与内容体检 ----------
  ctx.effect(() => ctx.tools.register(defineTool({
    name: 'office_audit',
    description: [
      '体检一个 Office 文件（.docx/.xlsx/.pptx）而不修改它：zip 完整性与必需部件、逐部件 XML 良构性、',
      'content-type 覆盖与悬空 Override、关系目标可解析，再用本插件自己的解析器重读一遍。',
      '另报内容层信号：未接受的修订痕迹、未填 {{占位符}}、TODO/TBD/lorem 草稿标记、',
      'xlsx 的错误值单元格与「公式没有缓存值」、pptx 的幻灯片部件缺失。',
      '用途是在没有 Office/LibreOffice 的环境里判断「这个文件打开会不会弹修复」，以及继承来的文件能不能安全接手；',
      '它不修文件，修由 office_edit 完成。',
    ].join(''),
    parameters: {
      source: { type: 'string', required: true, description: SOURCE_DESC },
    },
    output: {
      schema: {
        type: 'object',
        additionalProperties: false,
        properties: {
          source: { type: 'string', required: true },
          format: { type: 'string', required: true },
          ok: { type: 'boolean', required: true },
          entry_count: { type: 'integer', required: true },
          counts: {
            type: 'object',
            additionalProperties: false,
            required: true,
            properties: {
              paragraphs: { type: 'integer' },
              empty_paragraphs: { type: 'integer' },
              parts: { type: 'integer' },
              tables: { type: 'integer' },
              revisions: { type: 'integer' },
              sheets: { type: 'integer' },
              formulas: { type: 'integer' },
              formulas_without_value: { type: 'integer' },
              slides: { type: 'integer' },
              comments: { type: 'integer' },
              hyperlinks: { type: 'integer' },
            },
          },
          issues: {
            type: 'array',
            required: true,
            items: {
              type: 'object',
              additionalProperties: false,
              properties: {
                severity: { type: 'string', required: true },
                message: { type: 'string', required: true },
                where: { type: 'string' },
              },
            },
          },
        },
      },
      render: (_args, value) => [{ type: 'text', text: renderAudit(value) }],
    },
    isConcurrencySafe: () => true,
    async execute(args, exec) {
      if (args.source.trim() === '') throw new Error('source 参数不能为空')
      const read = await cache.readBytes(args.source, exec, caps)
      exec.signal.throwIfAborted()
      const report = auditOffice(read.bytes, read.displayName)
      const counts: Record<string, number> = {}
      const map: Array<[string, number | undefined]> = [
        ['paragraphs', report.counts.paragraphs],
        ['empty_paragraphs', report.counts.emptyParagraphs],
        ['parts', report.counts.parts],
        ['tables', report.counts.tables],
        ['revisions', report.counts.revisions],
        ['sheets', report.counts.sheets],
        ['formulas', report.counts.formulas],
        ['formulas_without_value', report.counts.formulasWithoutValue],
        ['slides', report.counts.slides],
        ['comments', report.counts.comments],
        ['hyperlinks', report.counts.hyperlinks],
      ]
      for (const [key, value] of map) {
        if (typeof value === 'number') counts[key] = value
      }
      return {
        source: report.source,
        format: report.format,
        ok: report.ok,
        entry_count: report.entryCount,
        counts,
        issues: report.issues.map((issue) => ({ severity: issue.severity, message: issue.message, ...(issue.where !== undefined ? { where: issue.where } : {}) })),
      }
    },
    presentCall: (args): GenericCallView => ({
      card: 'generic',
      title: `Audit ${basename(args.source)}`,
      kind: 'read',
      ...isLocalPath(args.source) ? { locations: [{ path: args.source }] } : {},
    }),
  })), 'office_audit')
}

// ---------- tool implementations ----------

interface InspectArgs { source: string }

async function inspectOffice(cache: OfficeCache, args: InspectArgs, exec: ToolExecution, caps: SourceCaps, config: Config) {
  if (args.source.trim() === '') throw new Error('source 参数不能为空')
  const doc = await openFor(cache, args.source, exec, caps)
  let units: Array<{ unit: string; chars: number; preview?: string }>
  let unitsTruncated = false
  const sheets: Array<{
    name: string; range: string; rows: number; cols: number; formulas: number; visibility: string
    header: string[]; column_types: string[]; comments: number; hyperlinks: number
  }> = []
  if (doc.format === 'xlsx') {
    // One aggregate pass per sheet has already produced rows/cols/formulas/header/types/samples, so
    // the overview never materializes row text (report §3.2.4 #9).
    units = (doc.xlsx?.sheets ?? []).map((sheet) => {
      const preview = sheet.sampleRows.length > 0 ? sheet.sampleRows.join(' ⏎ ').slice(0, config.previewChars * 3) : undefined
      const chars = sheet.sampleRows.reduce((total, row) => total + row.length, 0)
      return { unit: sheet.name, chars, ...(preview !== undefined ? { preview } : {}) }
    })
    for (const sheet of doc.xlsx?.sheets ?? []) {
      sheets.push({
        name: sheet.name,
        range: sheet.range,
        rows: sheet.rowCount,
        cols: sheet.colCount,
        formulas: sheet.formulaCount,
        visibility: sheet.visibility,
        header: sheet.headerCells.slice(0, 12),
        column_types: sheet.columnTypes.slice(0, 12),
        comments: sheet.commentCount,
        hyperlinks: sheet.hyperlinkCount,
      })
    }
  } else {
    const cap = Math.max(1, config.maxUnitsPerCall)
    units = doc.units
      .slice(0, cap)
      .map((unit) => (unit.empty === true
        ? { unit: unit.label, chars: 0, preview: '（空段落）' }
        : { unit: unit.label, chars: unit.chars, preview: unit.text.slice(0, config.previewChars) }))
    unitsTruncated = doc.units.length > cap
  }
  const revisions = doc.revisions
  return {
    source: doc.displayName,
    format: doc.format,
    metadata: cleanCoreProps(doc.metadata),
    outline: doc.outline.map((entry) => ({ unit: entry.unit, level: entry.level, title: entry.title })),
    units,
    units_truncated: unitsTruncated,
    sheets,
    defined_names: (doc.xlsx?.definedNames ?? []).map((entry) => ({
      name: entry.name,
      refers_to: entry.refersTo,
      ...(entry.scope !== undefined ? { scope: entry.scope } : {}),
    })),
    docx: {
      parts: (doc.parts ?? []).map((part) => ({
        label: part.label,
        kind: part.kind,
        chars: part.chars,
        ...(part.author !== undefined ? { author: part.author } : {}),
      })),
      table_count: new Set((doc.tables ?? []).map((row) => row.table)).size,
      table_rows: (doc.tables ?? []).length,
      revisions: {
        present: revisions?.present ?? false,
        insertions: revisions?.insertions ?? 0,
        deletions: revisions?.deletions ?? 0,
        format_changes: revisions?.formatChanges ?? 0,
        ...(revisions !== undefined && revisions.parts.length > 0 ? { parts: revisions.parts.join('、') } : {}),
      },
    },
  }
}

interface ReadArgs extends SelectorArgs { source: string; max_chars?: number }

async function readOffice(cache: OfficeCache, args: ReadArgs, exec: ToolExecution, caps: SourceCaps, config: Config) {
  if (args.source.trim() === '') throw new Error('source 参数不能为空')
  const doc = await openFor(cache, args.source, exec, caps)
  const selection = selectUnits(doc, args, config.maxUnitsPerCall)
  const requestedMax = args.max_chars ?? config.maxCharsPerCall
  if (!Number.isFinite(requestedMax) || requestedMax <= 0) throw new Error('max_chars 必须为正数')
  const maxChars = Math.floor(Math.min(requestedMax, config.maxCharsPerCall))
  const units: Array<{ unit: string; chars: number; text: string; empty?: boolean }> = []
  const charsPerUnit: Array<{ unit: string; chars: number }> = []
  let total = 0
  let truncated = false
  for (const unit of selection.units) {
    exec.signal.throwIfAborted()
    const remaining = maxChars - total
    if (unit.chars <= remaining) {
      charsPerUnit.push({ unit: unit.label, chars: unit.chars })
      // `empty` marks a structural, text-free unit (blank/page-break paragraph) so the reader can
      // keep it in the listing and preserve the paragraph numbering the editor also uses.
      units.push({ unit: unit.label, chars: unit.chars, text: unit.text, ...(unit.empty === true ? { empty: true } : {}) })
      total += unit.chars
      continue
    }
    if (remaining > 0) {
      // A partially returned unit reports the *page* length in chars and the slice in text,
      // so `chars` stays comparable with the un-truncated listing.
      charsPerUnit.push({ unit: unit.label, chars: unit.chars })
      units.push({ unit: unit.label, chars: unit.chars, text: unit.text.slice(0, remaining) })
      total += remaining
    }
    truncated = true
    break
  }
  const revisions = doc.revisions
  return {
    source: doc.displayName,
    format: doc.format,
    mode: selection.mode,
    requested: selection.requested,
    unit_count: doc.units.length,
    units,
    total_chars: total,
    truncated,
    revisions: {
      present: revisions?.present ?? false,
      insertions: revisions?.insertions ?? 0,
      deletions: revisions?.deletions ?? 0,
      format_changes: revisions?.formatChanges ?? 0,
    },
    chars_per_unit: charsPerUnit,
  }
}

interface FindArgs extends SelectorArgs { source: string; query: string; regex?: boolean; case_sensitive?: boolean; max_matches?: number }

async function findOffice(cache: OfficeCache, args: FindArgs, exec: ToolExecution, caps: SourceCaps, config: Config) {
  if (args.source.trim() === '') throw new Error('source 参数不能为空')
  const query = args.query.trim()
  if (query === '') throw new Error('query 不能为空')
  const useRegex = args.regex === true
  if (query.length > config.maxRegexPatternChars) {
    throw new Error(`query 过长（${query.length} 字符，上限 ${config.maxRegexPatternChars}）：请给出更具体的模式`)
  }
  // Literal mode still rejects a query that normalizes away to nothing; regex mode validates the
  // pattern instead, so a syntax error is reported before any unit is scanned.
  if (!useRegex && normalizeForMatch(query) === '') throw new Error('query 输入无效（全部为空白字符）')
  const pattern = useRegex ? compilePattern(query, args.case_sensitive === true) : null
  const doc = await openFor(cache, args.source, exec, caps)
  // 搜索在已解析数据上进行（缓存后近零成本），所以先数完再截断：`total_matches` 是「实际命中
  // 总数」而不是「返回了几条」，做到上限就停会让后者冒充前者（300 处命中报成 50 处）。
  const selection = selectUnits(doc, args, Number.POSITIVE_INFINITY)
  const maxMatches = Math.max(1, Math.min(args.max_matches ?? config.maxSearchMatches, config.maxSearchMatches))
  const found: Array<{ unit: string; count: number; snippets: string[]; captured?: JsonLike[] }> = []
  let truncated = false
  let searched = 0
  for (const unit of selection.units) {
    exec.signal.throwIfAborted()
    searched++
    const result = pattern === null
      ? literalSearch(unit.text, normalizeForMatch(query), config.contextChars)
      : regexSearch(unit.text, pattern, config.contextChars, 5, config.maxRegexMatchesPerUnit)
    if (result.count === 0) continue
    // `captured` is omitted entirely when no group matched: an undefined-valued key would break the
    // harness result channel, which requires lossless JSON.
    found.push({
      unit: unit.label,
      count: result.count,
      snippets: result.snippets,
      ...(result.captured.length > 0 ? { captured: result.captured } : {}),
    })
    // A unit that hit the per-unit counting cap, or holds more hits than its snippet budget, makes
    // the reported detail (not the count) partial — that is the "truncated" the caller must know.
    if (result.capped || result.count > result.snippets.length) truncated = true
  }
  const matches = found.slice(0, maxMatches)
  if (found.length > matches.length) truncated = true
  return {
    source: doc.displayName,
    format: doc.format,
    query,
    regex: useRegex,
    case_sensitive: pattern === null ? false : args.case_sensitive === true,
    matches,
    total_matches: matches.reduce((sum, match) => sum + match.count, 0),
    searched_units: searched,
    truncated,
  }
}

/**
 * Run the compiled pattern over one unit and collect the first capture groups alongside the
 * snippets, so a pattern like `(\d{4})年` answers "which years" without a second pass.
 *
 * Only the snippets that were actually kept contribute captures, keeping the two lists aligned:
 * `captured[i]` belongs to `snippets[i]`. Patterns without groups yield an empty array per hit,
 * which is filtered out by the caller.
 */
function regexSearch(unitText: string, pattern: RegExp, contextChars: number, maxSnippets: number, maxMatches: number): { count: number; snippets: string[]; captured: JsonLike[]; capped: boolean } {
  const result = searchUnitRegex(unitText, pattern, contextChars, maxSnippets, maxMatches)
  const captured: JsonLike[] = []
  if (result.snippets.length > 0) {
    pattern.lastIndex = 0
    for (let m = pattern.exec(unitText); m !== null && captured.length < result.snippets.length; m = pattern.exec(unitText)) {
      if (m[0] === '') {
        pattern.lastIndex = m.index + 1
        continue
      }
      if (m.length > 1) captured.push([...m].slice(1).map((group) => group ?? null))
    }
  }
  return { count: result.count, snippets: result.snippets, captured, capped: result.capped }
}

/** Literal-mode search: the normalized pipeline, shaped like {@link regexSearch} for the loop. */
function literalSearch(unitText: string, queryNorm: string, contextChars: number): { count: number; snippets: string[]; captured: JsonLike[]; capped: boolean } {
  const result = searchUnitFor(0, unitText, queryNorm, contextChars, 5)
  return { count: result.count, snippets: result.snippets, captured: [], capped: false }
}

// ---------- shared helpers ----------

function isLocalPath(source: string): boolean {
  return !/^(?:https?|file):/i.test(source)
}

function basename(source: string): string {
  const trimmed = source.trim()
  if (/^https?:\/\//i.test(trimmed)) return trimmed
  return trimmed.split(/[\\/]/u).at(-1) ?? trimmed
}

// ---------- renderers ----------

function renderInspect(value: {
  source: string; format: string; metadata: Record<string, unknown>;
  outline: Array<{ unit: string; level: number; title: string }>;
  units: Array<{ unit: string; chars: number; preview?: string }>;
  units_truncated: boolean;
  sheets: Array<{
    name: string; range: string; rows: number; cols: number; formulas: number; visibility: string
    header: string[]; column_types: string[]; comments: number; hyperlinks: number
  }>;
  defined_names: Array<{ name: string; refers_to: string; scope?: string }>;
  docx: {
    parts: Array<{ label: string; kind: string; chars: number; author?: string }>
    table_count: number; table_rows: number
    revisions: { present: boolean; insertions: number; deletions: number; format_changes: number }
  };
}): string {
  const lines: string[] = []
  const metaField = (key: string): string | undefined => {
    const raw = value.metadata[key]
    return typeof raw === 'string' && raw !== '' ? raw : undefined
  }
  const metaBits = [
    metaField('title') !== undefined ? `标题：${metaField('title')}` : undefined,
    metaField('creator') !== undefined ? `作者：${metaField('creator')}` : undefined,
    metaField('lastModifiedBy') !== undefined ? `最后修改：${metaField('lastModifiedBy')}` : undefined,
    metaField('created') !== undefined ? `创建：${metaField('created')}` : undefined,
    metaField('modified') !== undefined ? `修改：${metaField('modified')}` : undefined,
  ].filter((b): b is string => b !== undefined)
  lines.push(`Office 概览：${value.source}`)
  lines.push(`格式：${FORMAT_LABEL[value.format as keyof typeof FORMAT_LABEL] ?? value.format}`)
  if (metaBits.length > 0) lines.push(`元数据：${metaBits.join('；')}`)
  if (value.outline.length > 0) {
    lines.push('大纲：')
    for (const entry of value.outline) lines.push(`${'  '.repeat(Math.max(0, entry.level - 1))}- ${entry.title}（${entry.unit}）`)
  }
  lines.push(value.format === 'xlsx' ? '工作表内容：' : '内容单元：')
  for (const entry of value.units) {
    lines.push(`${entry.unit}（${entry.chars} 字符）${entry.preview !== undefined ? `：${entry.preview.replace(/\n/gu, ' ')}` : ''}`)
  }
  if (value.units_truncated) lines.push(`[单元列表已截断，仅显示前 ${value.units.length} 个]`)
  if (value.sheets.length > 0) {
    lines.push('工作表属性：')
    for (const sheet of value.sheets) {
      const header = sheet.header.length > 0 ? `；表头 ${sheet.header.join(' | ')}` : ''
      const types = sheet.column_types.length > 0 ? `；列类型 ${sheet.column_types.join('/')}` : ''
      const extras = [
        sheet.formulas > 0 ? `公式 ${sheet.formulas}` : undefined,
        sheet.comments > 0 ? `批注 ${sheet.comments}` : undefined,
        sheet.hyperlinks > 0 ? `超链接 ${sheet.hyperlinks}` : undefined,
        sheet.visibility !== 'visible' ? `可见性 ${sheet.visibility}` : undefined,
      ].filter((bit): bit is string => bit !== undefined)
      lines.push(`- ${sheet.name}：${sheet.rows} 行 × ${sheet.cols} 列，区域 ${sheet.range}${header}${types}${extras.length > 0 ? `；${extras.join('，')}` : ''}`)
    }
  }
  if (value.defined_names.length > 0) {
    lines.push(`定义名（${value.defined_names.length}）：`)
    for (const name of value.defined_names) lines.push(`- ${name.name}${name.scope !== undefined ? `（${name.scope}）` : ''} = ${name.refers_to}`)
  }
  if (value.docx.parts.length > 0) {
    lines.push(`正文之外：${value.docx.parts.map((part) => `${part.label}(${part.chars}字符${part.author !== undefined ? `，${part.author}` : ''})`).join('、')}`)
  }
  if (value.docx.table_count > 0) lines.push(`表格：${value.docx.table_count} 个，共 ${value.docx.table_rows} 行（用 tables=true 按行读取）`)
  if (value.docx.revisions.present === true) {
    lines.push(`修订痕迹：插入 ${value.docx.revisions.insertions}、删除 ${value.docx.revisions.deletions}、格式变更 ${value.docx.revisions.format_changes}（用 accept_revisions 接受全部）`)
  }
  return lines.join('\n')
}

function renderRead(value: {
  source: string; format: string; mode: string; requested: string; unit_count: number;
  units: Array<{ unit: string; chars: number; text: string; empty?: boolean }>;
  total_chars: number; truncated: boolean;
  revisions: { present: boolean; insertions: number; deletions: number; format_changes: number };
  chars_per_unit: Array<{ unit: string; chars: number }>;
}): string {
  const lines: string[] = []
  lines.push(`Office 文档文本（${value.source}）`)
  lines.push(`格式：${FORMAT_LABEL[value.format as keyof typeof FORMAT_LABEL] ?? value.format}；范围：${value.requested}；共 ${value.unit_count} 个单元；本次返回 ${value.total_chars} 字符`)
  if (value.revisions.present) {
    lines.push(`⚠ 本文档含未接受的修订：插入 ${value.revisions.insertions} 处、删除 ${value.revisions.deletions} 处、格式变更 ${value.revisions.format_changes} 处。此处显示的是「插入已接受、删除已丢弃」的文本；定稿前请用 office_edit 的 accept_revisions 明确接受，或先在 Word 中处理。`)
  }
  if (value.units.length === 0) lines.push('（未提取到文本：文档可能为空或仅含图片/图形）')
  for (const unit of value.units) {
    // Empty units are kept in the listing (not skipped) so that `段落 N` in this output is always
    // the same N that office_edit accepts — a hidden unit would silently shift every later index.
    if (unit.text.length === 0) {
      lines.push(`${unit.unit}：空段落`)
      continue
    }
    lines.push(`\n—— ${unit.unit}（${unit.chars} 字符）——`)
    lines.push(unit.text)
  }
  if (value.truncated) {
    const perUnit = value.chars_per_unit.map((e) => `${e.unit} ${e.chars}字符`).join('、')
    lines.push(`\n[已截断] 内容超过本次上限，已读单元字数：${perUnit}。请用范围选择器（paragraphs/section/parts/tables/slides/sheet+range/formulas）分段继续读取。`)
  }
  return lines.join('\n')
}

function renderFind(value: {
  source: string; format: string; query: string; regex: boolean; case_sensitive: boolean;
  matches: Array<{ unit: string; count: number; snippets: string[]; captured?: JsonLike[] }>;
  total_matches: number; searched_units: number; truncated: boolean;
}): string {
  const lines: string[] = []
  lines.push(value.regex ? `正则搜索：/${value.query}/g${value.case_sensitive ? '' : 'i'}` : `关键词搜索：${value.query}`)
  lines.push(`文档：${value.source}（${FORMAT_LABEL[value.format as keyof typeof FORMAT_LABEL] ?? value.format}，搜索了 ${value.searched_units} 个单元）`)
  if (!value.regex) lines.push('匹配规则：大小写不敏感、全半角与空白容错（如需按模式匹配请用 regex=true）')
  if (value.total_matches === 0) {
    lines.push('未找到匹配。')
    return lines.join('\n')
  }
  lines.push(`命中 ${value.total_matches} 处${value.truncated ? '（已达上限，仅显示部分）' : ''}：`)
  for (const match of value.matches) {
    lines.push(`${match.unit}（${match.count} 处）`)
    for (const snippet of match.snippets) lines.push(`  ${snippet}`)
  }
  const groups = value.matches.flatMap((match) => match.captured ?? []).filter((entry) => Array.isArray(entry) && entry.length > 0)
  if (groups.length > 0) {
    lines.push(`捕获组（按命中序 ${groups.length} 组）：${groups.slice(0, 20).map((entry) => JSON.stringify(entry)).join('、')}${groups.length > 20 ? ' …' : ''}`)
  }
  return lines.join('\n')
}

function renderAudit(value: {
  source: string; format: string; ok: boolean; entry_count: number;
  counts: Record<string, unknown>;
  issues: Array<{ severity: string; message: string; where?: string }>;
}): string {
  const lines: string[] = []
  lines.push(`Office 体检：${value.source}`)
  lines.push(`格式：${FORMAT_LABEL[value.format as keyof typeof FORMAT_LABEL] ?? value.format}；部件 ${value.entry_count} 个；结论：${value.ok ? '结构完整' : '发现结构性问题'}`)
  const counts = Object.entries(value.counts)
    .filter(([, raw]) => typeof raw === 'number')
    .map(([key, raw]) => `${key}=${String(raw)}`)
  if (counts.length > 0) lines.push(`计数：${counts.join('，')}`)
  if (value.issues.length === 0) {
    lines.push('未发现问题：zip 完整、必需部件齐全、所有 XML 良构、content-type 与关系均可解析。')
    return lines.join('\n')
  }
  const order = { ERROR: 0, WARN: 1, INFO: 2 } as Record<string, number>
  const sorted = [...value.issues].sort((a, b) => (order[a.severity] ?? 9) - (order[b.severity] ?? 9))
  lines.push(`发现 ${sorted.length} 项：`)
  for (const issue of sorted) lines.push(`[${issue.severity}] ${issue.message}${issue.where !== undefined ? `（${issue.where}）` : ''}`)
  return lines.join('\n')
}
