/**
 * 统一文档模型：格式识别、单元（unit）构建与读取/搜索范围选择。
 * 单元粒度：docx = 段落（另有 `parts`/`tables` 两个旁路模式）；xlsx = 工作表数据行
 * （另有 `formulas` 公式检视模式）；pptx = 幻灯片。
 * @module dsh-office-reader/src/office
 */
import { cleanCoreProps, type CoreProps } from './ooxml.js';
import { type DocxParagraph, type DocxPart, type DocxTableRow, type DocxRevisions } from './docx.js';
import { type PptxSlide } from './pptx.js';
import { type XlsxSheetInfo, type XlsxRow, type XlsxDocument } from './xlsx.js';
export type OfficeFormat = 'docx' | 'xlsx' | 'pptx';
/** One readable/searchable unit with a user-facing label. */
export interface OfficeUnit {
    label: string;
    text: string;
    chars: number;
    /** True for a structural but text-free unit (a blank/page-break docx paragraph). */
    empty?: boolean;
}
/** Outline entry pointing at a unit. */
export interface OutlineEntry {
    /** Unit label, e.g. "段落 3" / "Sheet1" / "幻灯片 2". */
    unit: string;
    /** 1-based index into the format's own unit list. */
    index: number;
    /** Heading level (docx headings); 1 for flat lists. */
    level: number;
    /** Title text. */
    title: string;
}
/** xlsx-specific data kept for sheet/range/formula selection. */
export interface XlsxData {
    workbook: unknown;
    sheets: XlsxSheetInfo[];
    /** Materialize every data row (cached by the parser). */
    rows: () => XlsxRow[];
    formulas: XlsxDocument['formulas'];
    formulasTruncated: boolean;
    definedNames: XlsxDocument['definedNames'];
    comments: XlsxDocument['comments'];
    hyperlinks: XlsxDocument['hyperlinks'];
}
/** Unified parsed document. */
export interface OfficeDocument {
    format: OfficeFormat;
    displayName: string;
    metadata: CoreProps;
    /**
     * Read/search granularity units, in document order.
     *
     * For xlsx this is a lazy getter: `office_inspect` never touches it, so a large workbook is not
     * materialized into row strings just to print a table of contents.
     */
    units: OfficeUnit[];
    outline: OutlineEntry[];
    /** Total unit characters, available without materializing `units` (used by the parse cache). */
    charCount: number;
    paragraphs?: DocxParagraph[];
    /** docx non-body parts (headers/footers/footnotes/endnotes/comments). */
    parts?: DocxPart[];
    /** docx table rows, for the table-row unit mode. */
    tables?: DocxTableRow[];
    /** docx revision markers. */
    revisions?: DocxRevisions;
    slides?: PptxSlide[];
    xlsx?: XlsxData;
}
/**
 * Open a document by extension, falling back to zip content sniffing.
 * .xls/.xlsm ride the SheetJS path alongside .xlsx.
 * @throws an Error with a user-facing Chinese message for unknown/legacy/encrypted files.
 */
export declare function openDocument(bytes: Uint8Array, displayName: string): OfficeDocument;
/** Kinds `parts` can select from a docx package. */
export declare const DOCX_PART_SELECTORS: readonly ["all", "headers", "footers", "footnotes", "endnotes", "comments"];
export type DocxPartSelector = typeof DOCX_PART_SELECTORS[number];
/** Model-supplied selectors (validated against the document format). */
export interface SelectorArgs {
    sheet?: string;
    range?: string;
    slides?: string;
    paragraphs?: string;
    section?: string;
    /** docx: read the non-body parts instead of the paragraphs. */
    parts?: DocxPartSelector | '';
    /** docx: read table rows instead of the paragraphs. */
    tables?: boolean;
    /** xlsx: read formula cells instead of data rows. */
    formulas?: boolean;
}
/** A validated selection of units. */
export interface UnitSelection {
    mode: 'full' | 'sheet' | 'range' | 'slides' | 'paragraphs' | 'section' | 'parts' | 'tables' | 'formulas';
    requested: string;
    units: OfficeUnit[];
}
/** Select units for read/find from format-specific selector args. */
export declare function selectUnits(doc: OfficeDocument, sel: SelectorArgs, maxUnits: number): UnitSelection;
export { cleanCoreProps };
