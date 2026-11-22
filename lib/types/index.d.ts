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
import type { Context } from 'cordis';
import z from 'schemastery';
export declare const name = "@dsh-external/dsh-office-reader";
export declare const inject: string[];
/** Deployment caps and defaults for the Office tools (all changeable from cordis.yml). */
export interface Config {
    /** Maximum characters returned by one `office_read` call; over-length output truncates. */
    maxCharsPerCall: number;
    /** Maximum units (paragraphs/rows/slides) selected by one call. */
    maxUnitsPerCall: number;
    /** Maximum hits returned by one `office_find` call across units. */
    maxSearchMatches: number;
    /** Maximum hits counted per unit in regex mode, bounding the work one pattern can cost. */
    maxRegexMatchesPerUnit: number;
    /** Maximum length of a `regex=true` pattern, bounding compile and match cost. */
    maxRegexPatternChars: number;
    /** Character length of the per-unit preview in `office_inspect`. */
    previewChars: number;
    /** Context characters kept on each side of a search match snippet. */
    contextChars: number;
    /** Maximum local document file size in bytes. */
    maxFileBytes: number;
    /** Maximum downloaded document body size in bytes. */
    maxDownloadBytes: number;
    /** Download timeout in milliseconds. */
    downloadTimeoutMs: number;
    /** Regex list of allowed URL hosts; empty allows any host. */
    allowedHosts: string[];
    /** Parsed-document cache: maximum cached documents (LRU). */
    cacheMaxEntries: number;
    /** Parsed-document cache: maximum total unit chars across cached documents. */
    cacheMaxChars: number;
}
export declare const Config: z<Schemastery.ObjectS<NoInfer<{
    maxCharsPerCall: z<number, number, "defined">;
    maxUnitsPerCall: z<number, number, "defined">;
    maxSearchMatches: z<number, number, "defined">;
    maxRegexMatchesPerUnit: z<number, number, "defined">;
    maxRegexPatternChars: z<number, number, "defined">;
    previewChars: z<number, number, "defined">;
    contextChars: z<number, number, "defined">;
    maxFileBytes: z<number, number, "defined">;
    maxDownloadBytes: z<number, number, "defined">;
    downloadTimeoutMs: z<number, number, "defined">;
    allowedHosts: z<string[], string[], "defined">;
    cacheMaxEntries: z<number, number, "defined">;
    cacheMaxChars: z<number, number, "defined">;
}>>, Schemastery.ObjectT<NoInfer<{
    maxCharsPerCall: z<number, number, "defined">;
    maxUnitsPerCall: z<number, number, "defined">;
    maxSearchMatches: z<number, number, "defined">;
    maxRegexMatchesPerUnit: z<number, number, "defined">;
    maxRegexPatternChars: z<number, number, "defined">;
    previewChars: z<number, number, "defined">;
    contextChars: z<number, number, "defined">;
    maxFileBytes: z<number, number, "defined">;
    maxDownloadBytes: z<number, number, "defined">;
    downloadTimeoutMs: z<number, number, "defined">;
    allowedHosts: z<string[], string[], "defined">;
    cacheMaxEntries: z<number, number, "defined">;
    cacheMaxChars: z<number, number, "defined">;
}>>, "plain">;
export declare function apply(ctx: Context, config: Config): void;
