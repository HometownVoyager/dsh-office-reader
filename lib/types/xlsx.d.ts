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
import { type CoreProps, type OoxmlPackage } from './ooxml.js';
/** Column type tags used by the per-sheet overview. */
export type ColumnType = 'int' | 'float' | 'bool' | 'datetime' | 'str' | 'formula' | 'mixed' | 'empty';
/** Sheet summary derived from one aggregate pass over the used range. */
export interface XlsxSheetInfo {
    /** Sheet name. */
    name: string;
    /** 0-based position in the workbook. */
    index: number;
    /** Used range in A1 notation, e.g. "A1:C12". */
    range: string;
    /** Number of non-empty data rows. */
    rowCount: number;
    /** Widest row in cells. */
    colCount: number;
    /** Cells carrying a formula. */
    formulaCount: number;
    /** First non-empty row's cells (the header row, by convention). */
    headerCells: string[];
    /** Inferred type per column over the data rows below the header. */
    columnTypes: ColumnType[];
    /** Up to three row texts, for the overview preview. */
    sampleRows: string[];
    /** Visibility as stored in `workbook.xml` (`state` attribute). */
    visibility: 'visible' | 'hidden' | 'veryHidden';
    /** Convenience flag: `visibility !== 'visible'`. */
    hidden: boolean;
    /** Legacy cell comments attached to this sheet. */
    commentCount: number;
    /** Hyperlinks attached to this sheet. */
    hyperlinkCount: number;
}
/** One non-empty data row flattened to text. */
export interface XlsxRow {
    /** Sheet name. */
    sheet: string;
    /** 0-based sheet index. */
    sheetIndex: number;
    /** 1-based sheet row number (position in the used range). */
    row: number;
    /** Cell texts with trailing empties trimmed. */
    cells: string[];
    /** Cells joined by tab. */
    text: string;
    /** Text length in characters. */
    chars: number;
}
/** One formula cell, for formula-inspection reads. */
export interface XlsxFormulaCell {
    sheet: string;
    sheetIndex: number;
    /** A1 address, e.g. `D2`. */
    addr: string;
    /** 1-based row of the cell, so callers can order/group by row. */
    row: number;
    /** Formula text without the leading `=`. */
    formula: string;
}
/** One cell comment (legacy note), read from the comments part. */
export interface XlsxComment {
    sheet: string;
    sheetIndex: number;
    /** A1 address the note is attached to. */
    addr: string;
    author: string;
    text: string;
}
/** One hyperlink, read from the sheet part plus its relationships. */
export interface XlsxHyperlink {
    sheet: string;
    sheetIndex: number;
    addr: string;
    /** External URL, or `#Sheet1!A1` for an internal link. */
    url: string;
    tooltip?: string;
}
/** One defined name. */
export interface XlsxDefinedName {
    name: string;
    refersTo: string;
    /** Sheet the name is scoped to; absent means workbook scope. */
    scope?: string;
}
/** Parsed workbook. */
export interface XlsxDocument {
    format: 'xlsx';
    workbook: unknown;
    sheets: XlsxSheetInfo[];
    metadata: CoreProps;
    /** Workbook- and sheet-scoped defined names. */
    definedNames: XlsxDefinedName[];
    comments: XlsxComment[];
    hyperlinks: XlsxHyperlink[];
    /** Every formula cell in the workbook (bounded by Config.maxFormulas). */
    formulas: XlsxFormulaCell[];
    /** True when the formula list hit its cap. */
    formulasTruncated: boolean;
    /** Total characters of all row texts, computed without materializing them. */
    charCount: number;
    /** Materialize every data row (cached after the first call). */
    rows(): XlsxRow[];
}
/**
 * Parse a workbook (.xlsx/.xlsm/.xls) via SheetJS.
 * @param bytes - raw file bytes.
 * @param pkg - unpacked OOXML package for xlsx/xlsm parts (comments/hyperlinks). Omit to let the
 *   parser unpack the container itself; pass `null` for legacy `.xls`, which has no parts.
 * @throws an Error with a user-facing Chinese message when parsing fails.
 */
export declare function readXlsx(bytes: Uint8Array, pkg?: OoxmlPackage | null): XlsxDocument;
/**
 * Read a cell range from one sheet as text rows with their real row numbers.
 * @param workbook - the parsed workbook (opaque handle from {@link XlsxDocument}).
 * @param sheetName - exact sheet name.
 * @param range - A1 range like "A1:C10", or a start cell like "B2" (reads to the end).
 */
export declare function readXlsxRangeRows(workbook: unknown, sheetName: string, range: string): Array<{
    row: number;
    text: string;
}>;
