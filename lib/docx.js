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
import { attribute, openOoxml, readCoreProps, splitElements, splitWordParagraphs, wordHeadingLevel, } from './ooxml.js';
/**
 * Parse a .docx document into numbered paragraphs, non-body parts, table rows and revision counts.
 *
 * Every top-level paragraph becomes a unit, empty ones included: skipping them (as this reader
 * used to) makes `段落 N` disagree with the editor's paragraph index, so a model that reads a
 * document and then edits "paragraph 5" rewrites a different paragraph.
 * @param bytes - raw document bytes.
 * @throws an Error with a user-facing Chinese message when not a valid docx.
 */
export function readDocx(bytes) {
    const pkg = openOoxml(bytes);
    const documentXml = pkg.text('word/document.xml');
    if (documentXml === null)
        throw new Error('不是有效的 .docx 文档（缺少 word/document.xml）');
    const paragraphs = splitWordParagraphs(documentXml).map((span, i) => ({
        index: i + 1,
        level: wordHeadingLevel(span.xml),
        text: span.text,
        chars: span.text.length,
        empty: span.empty,
    }));
    return {
        format: 'docx',
        paragraphs,
        parts: readParts(pkg),
        tables: readTables(documentXml),
        revisions: countRevisions(pkg, documentXml),
        metadata: readCoreProps(pkg),
        charCount: paragraphs.reduce((total, p) => total + p.chars, 0),
    };
}
/** Number of empty (text-free) paragraphs, for reader summaries. */
export function countEmptyParagraphs(paragraphs) {
    return paragraphs.reduce((total, p) => total + (p.empty ? 1 : 0), 0);
}
/** Text of one part's paragraphs, joined by newline; empty paragraphs drop out. */
function partText(xml) {
    return splitWordParagraphs(xml)
        .map((span) => span.text)
        .filter((text) => text !== '')
        .join('\n');
}
/**
 * Read every non-body part the package carries.
 *
 * Footnotes and endnotes are split per `w:footnote`/`w:endnote`: Word's separator entries
 * (`w:id="-1"`/`"0"`) are structural, not content, so they are skipped. Comments keep their author
 * and date, and are numbered by `w:id` so the ids match what `office_edit`'s `add_comment` writes.
 */
function readParts(pkg) {
    const out = [];
    const names = Object.keys(pkg.files).sort();
    for (const kind of ['header', 'footer']) {
        const parts = names.filter((name) => new RegExp(`^word/${kind}\\d+\\.xml$`, 'u').test(name));
        parts.forEach((name, i) => {
            const text = partText(pkg.text(name));
            out.push({ kind, part: name, label: `${kind === 'header' ? '页眉' : '页脚'} ${i + 1}`, text, chars: text.length });
        });
    }
    for (const [kind, file, label] of [['footnote', 'word/footnotes.xml', '脚注'], ['endnote', 'word/endnotes.xml', '尾注']]) {
        const xml = pkg.text(file);
        if (xml === null)
            continue;
        for (const note of splitElements(xml, kind)) {
            const id = Number(attribute(note.xml, 'w:id') ?? 'NaN');
            if (id === -1 || id === 0 || !Number.isFinite(id))
                continue;
            const text = partText(note.xml);
            out.push({ kind, part: file, label: `${label} ${id}`, text, chars: text.length, id });
        }
    }
    const commentsXml = pkg.text('word/comments.xml');
    if (commentsXml !== null) {
        for (const comment of splitElements(commentsXml, 'comment')) {
            const id = Number(attribute(comment.xml, 'w:id') ?? 'NaN');
            const author = attribute(comment.xml, 'w:author');
            const date = attribute(comment.xml, 'w:date');
            const text = partText(comment.xml);
            out.push({
                kind: 'comment',
                part: 'word/comments.xml',
                label: `批注 ${Number.isFinite(id) ? id : out.length}`,
                text,
                chars: text.length,
                ...(Number.isFinite(id) ? { id } : {}),
                ...(author !== undefined && author !== '' ? { author } : {}),
                ...(date !== undefined && date !== '' ? { date } : {}),
            });
        }
    }
    return out;
}
/**
 * Flatten every top-level table into row units.
 *
 * Cells are joined by a tab (the same separator `office_read` uses for xlsx rows), so a table read
 * in this mode lines up column-for-column with the spreadsheet reading of the same data. A nested
 * table's rows belong to the inner table and are not repeated here.
 */
function readTables(documentXml) {
    const out = [];
    splitElements(documentXml, 'tbl').forEach((table, tableIndex) => {
        splitElements(table.xml, 'tr').forEach((row, rowIndex) => {
            const cells = splitElements(row.xml, 'tc').map((cell) => partText(cell.xml).replace(/\t/gu, ' ').replace(/\r?\n/gu, ' '));
            const text = cells.join('\t');
            out.push({ table: tableIndex + 1, row: rowIndex + 1, text, chars: text.length, cells, header: rowIndex === 0 });
        });
    });
    return out;
}
/** Count revision markers across every content part. */
function countRevisions(pkg, documentXml) {
    let insertions = 0;
    let deletions = 0;
    let formatChanges = 0;
    const parts = [];
    for (const name of Object.keys(pkg.files)) {
        if (!/\.xml$/u.test(name) || name.startsWith('docProps/') || name === '[Content_Types].xml')
            continue;
        const xml = name === 'word/document.xml' ? documentXml : pkg.text(name);
        if (xml === null)
            continue;
        const ins = (xml.match(/<w:ins(?=[\s/>])/gu) ?? []).length;
        const del = (xml.match(/<w:del(?=[\s/>])/gu) ?? []).length;
        const fmt = (xml.match(/<w:(?:rPr|pPr|tblPr|trPr|tcPr|sectPr|numbering)Change(?=[\s/>])/gu) ?? []).length;
        if (ins + del + fmt === 0)
            continue;
        insertions += ins;
        deletions += del;
        formatChanges += fmt;
        parts.push(name);
    }
    return { insertions, deletions, formatChanges, parts, present: insertions + deletions + formatChanges > 0 };
}
//# sourceMappingURL=docx.js.map