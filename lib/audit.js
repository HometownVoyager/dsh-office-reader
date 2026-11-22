/**
 * office_audit 的实现：把「这个文件在 Excel/Word/PowerPoint 里打开会不会弹修复」的检查搬到
 * 纯 Node 里做完——zip 完整性、必需部件、逐部件 XML 良构、content-type 覆盖、关系可解析，
 * 再用本插件自己的解析器重读一遍（报告 §4.2 第 19 条、§6 P2「audit 工具」）。
 *
 * 与写入侧校验的分工：`dsh-office-editor` 的 verify* 只保证「刚写出来的字节没有结构性损坏」；
 * 本模块面向**继承来的文件**（别人给的模板、上一步流程的产物），因此除结构外还报内容层的可疑
 * 信号：修订痕迹、未填占位符、公式缓存值缺失、错误值 token（#REF! / #NAME? …）。
 *
 * @module dsh-office-reader/src/audit
 */
import { strFromU8, unzipSync } from 'fflate';
import XLSX from 'xlsx';
import { assertWellFormed, notOoxmlMessage, openOoxml, resolveRelTarget } from './ooxml.js';
import { openDocument } from './office.js';
/** Error-value and leftover tokens worth reporting, with their severity. */
const TOKEN_CHECKS = [
    { pattern: /#REF!/u, severity: 'ERROR', label: '公式引用了已删除的单元格（#REF!）' },
    { pattern: /#NAME\?/u, severity: 'WARN', label: '公式含无法识别的名称（#NAME?）' },
    { pattern: /#VALUE!/u, severity: 'WARN', label: '公式参数类型错误（#VALUE!）' },
    { pattern: /#DIV\/0!/u, severity: 'WARN', label: '公式除以零（#DIV/0!）' },
    { pattern: /#N\/A/u, severity: 'WARN', label: '查值未命中（#N/A）' },
    { pattern: /#NULL!/u, severity: 'WARN', label: '区域交集为空（#NULL!）' },
    { pattern: /#NUM!/u, severity: 'WARN', label: '数值无效（#NUM!）' },
    { pattern: /\{\{[^}]{1,80}\}\}/u, severity: 'WARN', label: '存在未填充的 {{占位符}}' },
    { pattern: /\bTODO\b|\bTBD\b|\bXXXX\b|\blorem ipsum\b/iu, severity: 'INFO', label: '存在 TODO/TBD/lorem 之类的草稿标记' },
];
/** Highest severity present, for the summary line. */
function worst(issues) {
    if (issues.some((issue) => issue.severity === 'ERROR'))
        return 'ERROR';
    if (issues.some((issue) => issue.severity === 'WARN'))
        return 'WARN';
    if (issues.length > 0)
        return 'INFO';
    return null;
}
/**
 * Audit one Office file.
 * @param bytes - raw file bytes.
 * @param displayName - source name used in messages.
 * @throws an Error when the bytes are not an OOXML zip container at all (nothing to audit).
 */
export function auditOffice(bytes, displayName) {
    let files;
    try {
        files = unzipSync(bytes);
    }
    catch {
        throw new Error(notOoxmlMessage(bytes, '无法审计'));
    }
    const parts = Object.keys(files).sort();
    const issues = [];
    const pkg = openOoxml(bytes);
    const format = pkg.has('word/document.xml') ? 'docx' : (pkg.has('ppt/presentation.xml') ? 'pptx' : 'xlsx');
    const required = format === 'docx'
        ? ['[Content_Types].xml', '_rels/.rels', 'word/document.xml']
        : format === 'pptx'
            ? ['[Content_Types].xml', '_rels/.rels', 'ppt/presentation.xml']
            : ['[Content_Types].xml', '_rels/.rels', 'xl/workbook.xml'];
    for (const name of required) {
        if (files[name] === undefined)
            issues.push({ severity: 'ERROR', message: `缺少必需部件 ${name}——Office 打开时会报「内容有问题」`, where: name });
    }
    // 1. Well-formedness of every XML part (the damage class string surgery produces).
    let checkedParts = 0;
    for (const name of parts) {
        if (!/\.(?:xml|rels)$/u.test(name))
            continue;
        checkedParts++;
        try {
            assertWellFormed(strFromU8(files[name]), name);
        }
        catch (error) {
            issues.push({ severity: 'ERROR', message: error instanceof Error ? error.message : String(error), where: name });
        }
    }
    // 2. Content-type coverage, both directions.
    const contentTypes = files['[Content_Types].xml'] === undefined ? null : strFromU8(files['[Content_Types].xml']);
    if (contentTypes !== null) {
        const defaults = new Set([...contentTypes.matchAll(/<Default\b[^>]*Extension="([^"]*)"/gu)].map((m) => m[1].toLowerCase()));
        const overrides = new Set([...contentTypes.matchAll(/<Override\b[^>]*PartName="([^"]*)"/gu)].map((m) => m[1].replace(/^\//u, '')));
        for (const name of parts) {
            if (name === '[Content_Types].xml')
                continue;
            const ext = name.split('.').at(-1)?.toLowerCase() ?? '';
            if (overrides.has(name) || defaults.has(ext))
                continue;
            issues.push({ severity: 'WARN', message: `部件没有登记 content-type（既无 Override 也无同扩展名 Default），Office 可能拒绝打开或自动修复`, where: name });
        }
        for (const part of overrides) {
            if (files[part] === undefined)
                issues.push({ severity: 'WARN', message: 'content-type 指向的部件不存在（悬空 Override）', where: part });
        }
    }
    // 3. Relationship targets resolve.
    for (const name of parts) {
        const declared = /^(.*?)_rels\/([^/]*)\.rels$/u.exec(name);
        if (declared === null)
            continue;
        // `xl/_rels/workbook.xml.rels` declares relationships of `xl/workbook.xml`; `_rels/.rels`
        // declares the package's own, i.e. base directory = package root.
        const owner = `${declared[1]}${declared[2]}`;
        const relsXml = strFromU8(files[name]);
        for (const match of relsXml.matchAll(/<Relationship\b[^>]*\/>/gu)) {
            const target = /Target="([^"]*)"/u.exec(match[0])?.[1];
            if (target === undefined || /TargetMode="External"/u.test(match[0]))
                continue;
            const resolved = resolveRelTarget(owner, target);
            if (resolved === null || files[resolved] !== undefined)
                continue;
            issues.push({ severity: 'WARN', message: `关系指向的部件不存在：${target}`, where: name });
        }
    }
    // 4. Re-read with this plugin's own parser.
    const counts = {};
    let doc = null;
    try {
        doc = openDocument(bytes, displayName);
    }
    catch (error) {
        issues.push({ severity: 'ERROR', message: `本插件解析失败：${error instanceof Error ? error.message : String(error)}`, where: 'parse' });
    }
    if (doc !== null) {
        if (doc.format === 'docx') {
            counts.paragraphs = doc.units.length;
            counts.emptyParagraphs = (doc.paragraphs ?? []).filter((p) => p.empty).length;
            counts.parts = (doc.parts ?? []).length;
            counts.tables = new Set((doc.tables ?? []).map((row) => row.table)).size;
            counts.comments = (doc.parts ?? []).filter((part) => part.kind === 'comment').length;
            const revisions = doc.revisions;
            counts.revisions = revisions === undefined ? 0 : revisions.insertions + revisions.deletions + revisions.formatChanges;
            if (revisions?.present === true) {
                const detail = `插入 ${revisions.insertions}、删除 ${revisions.deletions}、格式变更 ${revisions.formatChanges}，涉及 ${revisions.parts.length} 个部件`;
                issues.push({ severity: 'WARN', message: `文档含未接受的修订痕迹（${detail}）：读取按「已接受插入、丢弃删除」呈现，定稿前请确认，或用 office_edit 的 accept_revisions 一次性接受`, where: 'revisions' });
            }
            if (counts.paragraphs === 0)
                issues.push({ severity: 'WARN', message: '正文没有任何段落', where: 'word/document.xml' });
            for (const unit of doc.units)
                scanTokens(unit.text, unit.label, issues);
            for (const part of doc.parts ?? [])
                scanTokens(part.text, part.label, issues);
        }
        else if (doc.format === 'pptx') {
            counts.slides = doc.units.length;
            if (counts.slides === 0)
                issues.push({ severity: 'WARN', message: '演示文稿没有任何幻灯片', where: 'ppt/presentation.xml' });
            // Every slide part must be reachable from sldIdLst and registered as a content type.
            const presentation = pkg.text('ppt/presentation.xml') ?? '';
            const listed = new Set([...presentation.matchAll(/<p:sldId\b[^>]*r:id="([^"]*)"/gu)].map((m) => m[1]));
            const rels = pkg.text('ppt/_rels/presentation.xml.rels') ?? '';
            const relTargets = new Map([...rels.matchAll(/<Relationship\b[^>]*\/>/gu)].map((m) => [/Id="([^"]*)"/u.exec(m[0])?.[1] ?? '', /Target="([^"]*)"/u.exec(m[0])?.[1] ?? '']));
            for (const id of listed) {
                const target = relTargets.get(id);
                if (target === undefined) {
                    issues.push({ severity: 'ERROR', message: `presentation.xml 的幻灯片关系 ${id} 在 rels 中不存在`, where: 'ppt/_rels/presentation.xml.rels' });
                    continue;
                }
                const part = `ppt/${target.replace(/^\.\//u, '')}`;
                if (files[part] === undefined)
                    issues.push({ severity: 'ERROR', message: `幻灯片部件缺失：${part}`, where: part });
            }
            for (const unit of doc.units)
                scanTokens(unit.text, unit.label, issues);
        }
        else {
            counts.sheets = doc.xlsx?.sheets.length ?? 0;
            counts.formulas = doc.xlsx?.formulas.length ?? 0;
            counts.comments = doc.xlsx?.comments.length ?? 0;
            counts.hyperlinks = doc.xlsx?.hyperlinks.length ?? 0;
            if (counts.formulas !== undefined && counts.formulas > 0) {
                const cached = countFormulaValues(doc.xlsx?.workbook);
                counts.formulasWithoutValue = doc.xlsx?.formulasTruncated === true ? undefined : Math.max(0, counts.formulas - cached);
                if ((counts.formulasWithoutValue ?? 0) > 0) {
                    issues.push({ severity: 'WARN', message: `${counts.formulasWithoutValue} 个公式单元格没有缓存值：Excel 打开时会重算，但按缓存值读取的工具（pandas、本插件的 office_read、CSV 导出）会看到空白或 0`, where: 'xl/worksheets' });
                }
            }
            if (!/<calcPr\b/u.test(pkg.text('xl/workbook.xml') ?? '')) {
                issues.push({ severity: 'INFO', message: 'xl/workbook.xml 没有 <calcPr fullCalcOnLoad="1"/>：Excel 不一定会重算公式缓存值', where: 'xl/workbook.xml' });
            }
            for (const formula of doc.xlsx?.formulas ?? [])
                scanTokens(formula.formula, `${formula.sheet}!${formula.addr}`, issues);
            for (const comment of doc.xlsx?.comments ?? [])
                scanTokens(comment.text, `${comment.sheet}!${comment.addr}`, issues);
            scanWorkbookValues(doc.xlsx?.workbook, issues);
        }
    }
    // 5. Placeholder leftovers in the raw parts, for parts the reader does not surface as text.
    if (format === 'pptx') {
        for (const name of parts.filter((n) => /^ppt\/notesSlides\/notesSlide\d+\.xml$/u.test(n))) {
            scanTokens(pkg.text(name) ?? '', name, issues);
        }
    }
    return {
        source: displayName,
        format,
        entryCount: parts.length,
        parts,
        issues,
        counts,
        ok: !issues.some((issue) => issue.severity === 'ERROR'),
    };
}
/** Report token hits inside one text blob. */
function scanTokens(text, where, issues) {
    if (text === '')
        return;
    for (const check of TOKEN_CHECKS) {
        if (!check.pattern.test(text))
            continue;
        issues.push({ severity: check.severity, message: check.label, where });
    }
}
/** Count formula cells that carry a cached value, for the workbook-level warning. */
function countFormulaValues(workbook) {
    const wb = workbook;
    if (wb === undefined)
        return 0;
    let withValue = 0;
    for (const name of wb.SheetNames) {
        const ws = wb.Sheets[name];
        if (ws === undefined)
            continue;
        for (const key of Object.keys(ws)) {
            if (!/^[A-Z]+[0-9]+$/u.test(key))
                continue;
            const cell = ws[key];
            if (typeof cell?.f !== 'string' || cell.f === '')
                continue;
            // `t: 'z'` is SheetJS's stub for "cell exists, no cached value" — its `v` is a placeholder 0.
            if (cell.t === 'z' || cell.v === undefined)
                continue;
            withValue++;
        }
    }
    return withValue;
}
/** Scan cell display values for error tokens without materializing row text. */
function scanWorkbookValues(workbook, issues) {
    const wb = workbook;
    if (wb === undefined)
        return;
    for (const name of wb.SheetNames) {
        const ws = wb.Sheets[name];
        if (ws === undefined)
            continue;
        const ref = typeof ws['!ref'] === 'string' ? ws['!ref'] : null;
        if (ref === null)
            continue;
        const range = XLSX.utils.decode_range(ref);
        const errors = [];
        for (let r = range.s.r; r <= range.e.r; r++) {
            for (let c = range.s.c; c <= range.e.c; c++) {
                const addr = XLSX.utils.encode_cell({ r, c });
                const cell = ws[addr];
                if (cell === undefined)
                    continue;
                if (cell.t === 'e' || (typeof cell.w === 'string' && cell.w.startsWith('#'))) {
                    errors.push(`${addr}=${String(cell.w ?? cell.v)}`);
                }
            }
        }
        if (errors.length === 0)
            continue;
        // One finding per sheet: a broken lookup column can produce hundreds of identical cells.
        issues.push({
            severity: 'WARN',
            message: `工作表有 ${errors.length} 个错误值单元格：${errors.slice(0, 5).join('、')}${errors.length > 5 ? ' 等' : ''}`,
            where: name,
        });
    }
}
/** Count of parts and the worst severity, for the tool's one-line summary. */
export function auditHeadline(report) {
    const level = worst(report.issues);
    if (level === null)
        return `未发现问题（${report.entryCount} 个部件，全部通过）`;
    return `${level}：${report.issues.length} 项发现`;
}
//# sourceMappingURL=audit.js.map