/**
 * 进程内解析缓存：同文档的 read/find/inspect 复用一次解析结果。
 *
 * - 本地文件：key = 绝对路径 + size + mtimeMs，命中时不读盘、不解析，仅一次 stat；
 * - http(s)/file://：下载后按 sha1(bytes) 作 key（下载本身不可避免）；
 * - LRU 淘汰：条目数与总字符数双上限；同路径旧 mtime 条目即时清理；
 * - 并发去重：同 key 解析中时共享同一 Promise。
 * @module dsh-office-reader/src/cache
 */
import type { ToolExecution } from '@deepseek-ai/dsh-tools';
import { type OfficeDocument } from './office.js';
import { type SourceCaps } from './source.js';
/** Cache size caps. */
export interface CacheCaps {
    /** Maximum cached documents (LRU). */
    maxEntries: number;
    /** Maximum total unit chars across cached documents. */
    maxChars: number;
}
export declare class OfficeCache {
    private caps;
    private entries;
    private inflight;
    private counter;
    constructor(caps: CacheCaps);
    /**
     * Open a document through the cache: stat-based key for local files
     * (cache hit skips file IO and parsing entirely), content hash for URLs.
     */
    open(source: string, exec: ToolExecution, sourceCaps: SourceCaps): Promise<OfficeDocument>;
    private openRemote;
    /**
     * Read a source's raw bytes without parsing.
     *
     * `office_audit` works on the package itself (zip entries, parts, relationships), so it needs the
     * bytes rather than a parsed document. Parsed-document caching does not apply; local files are
     * read once per call, remote sources go through the same download path as everything else.
     */
    readBytes(source: string, exec: ToolExecution, sourceCaps: SourceCaps): Promise<{
        bytes: Uint8Array;
        displayName: string;
        absolute: string | null;
    }>;
    private through;
    private store;
    private localPath;
}
