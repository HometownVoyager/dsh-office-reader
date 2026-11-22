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
export type AuditSeverity = 'ERROR' | 'WARN' | 'INFO';
/** One finding. */
export interface AuditIssue {
    severity: AuditSeverity;
    /** What was found. */
    message: string;
    /** Part or unit the finding belongs to. */
    where?: string;
}
/** Audit result for one file. */
export interface AuditReport {
    source: string;
    format: 'docx' | 'xlsx' | 'pptx';
    entryCount: number;
    /** Package parts, sorted. */
    parts: string[];
    issues: AuditIssue[];
    /** Counting fields the caller can render or assert on. */
    counts: {
        paragraphs?: number;
        emptyParagraphs?: number;
        parts?: number;
        tables?: number;
        revisions?: number;
        sheets?: number;
        formulas?: number;
        formulasWithoutValue?: number;
        slides?: number;
        comments?: number;
        hyperlinks?: number;
    };
    /** True when no ERROR-level issue was found. */
    ok: boolean;
}
/**
 * Audit one Office file.
 * @param bytes - raw file bytes.
 * @param displayName - source name used in messages.
 * @throws an Error when the bytes are not an OOXML zip container at all (nothing to audit).
 */
export declare function auditOffice(bytes: Uint8Array, displayName: string): AuditReport;
/** Count of parts and the worst severity, for the tool's one-line summary. */
export declare function auditHeadline(report: AuditReport): string;
