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
import { type CoreProps } from './ooxml.js';
/** One body paragraph with its 1-based index and optional heading level. */
export interface DocxParagraph {
    /** 1-based index among all top-level paragraphs (empty ones included, so the index is stable). */
    index: number;
    /** Heading level (1-9) when the paragraph carries one. */
    level?: number;
    /** Extracted text; empty for blank and page-break-only paragraphs. */
    text: string;
    /** Text length in characters. */
    chars: number;
    /** True when the paragraph carries no visible text (blank line, page break, section break). */
    empty: boolean;
}
/** Kinds of non-body parts this reader exposes. */
export type DocxPartKind = 'header' | 'footer' | 'footnote' | 'endnote' | 'comment';
/** One unit of a non-body part (a header/footer as a whole, or one note/comment). */
export interface DocxPart {
    kind: DocxPartKind;
    /** Part file name (e.g. `word/header1.xml`). */
    part: string;
    /** User-facing label, e.g. `页眉 1` / `脚注 3` / `批注 0`. */
    label: string;
    /** Plain text (paragraphs joined by newline; notes/comments keep their own paragraphs). */
    text: string;
    chars: number;
    /** Comment/note author, when the part carries one. */
    author?: string;
    /** Comment/note timestamp (`w:date`), when present. */
    date?: string;
    /** Note/comment id inside its part (`w:id`). */
    id?: number;
}
/** One table row in the table-row unit mode. */
export interface DocxTableRow {
    /** 1-based table index in document order. */
    table: number;
    /** 1-based row index inside that table. */
    row: number;
    /** Cell texts joined by tab ("列1\t列2"). */
    text: string;
    chars: number;
    /** Cell texts, for callers that want the columns separately. */
    cells: string[];
    /** True for the first row of the table (usually the header row). */
    header: boolean;
}
/** Revision markers found in the package. */
export interface DocxRevisions {
    /** `w:ins` insertions. */
    insertions: number;
    /** `w:del` deletions. */
    deletions: number;
    /** Formatting changes (`w:rPrChange`, `w:pPrChange`, …). */
    formatChanges: number;
    /** Parts that carry at least one marker. */
    parts: string[];
    /** True when any marker exists. */
    present: boolean;
}
/** Parsed .docx document. */
export interface DocxDocument {
    format: 'docx';
    paragraphs: DocxParagraph[];
    /** Non-body parts, in a stable order (headers, footers, footnotes, endnotes, comments). */
    parts: DocxPart[];
    /** Every table row, in document order. */
    tables: DocxTableRow[];
    /** Revision markers across the package. */
    revisions: DocxRevisions;
    metadata: CoreProps;
    /** Total visible characters of all paragraphs (lets the parse cache size the doc without a re-walk). */
    charCount: number;
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
export declare function readDocx(bytes: Uint8Array): DocxDocument;
/** Number of empty (text-free) paragraphs, for reader summaries. */
export declare function countEmptyParagraphs(paragraphs: DocxParagraph[]): number;
