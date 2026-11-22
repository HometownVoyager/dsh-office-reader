/**
 * 进程内解析缓存：同文档的 read/find/inspect 复用一次解析结果。
 *
 * - 本地文件：key = 绝对路径 + size + mtimeMs，命中时不读盘、不解析，仅一次 stat；
 * - http(s)/file://：下载后按 sha1(bytes) 作 key（下载本身不可避免）；
 * - LRU 淘汰：条目数与总字符数双上限；同路径旧 mtime 条目即时清理；
 * - 并发去重：同 key 解析中时共享同一 Promise。
 * @module dsh-office-reader/src/cache
 */
import { createHash } from 'node:crypto';
import { readFile, stat } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { resolve } from 'node:path';
import { openDocument } from './office.js';
import { resolveSource } from './source.js';
export class OfficeCache {
    caps;
    entries = new Map();
    inflight = new Map();
    counter = 0;
    constructor(caps) {
        this.caps = caps;
    }
    /**
     * Open a document through the cache: stat-based key for local files
     * (cache hit skips file IO and parsing entirely), content hash for URLs.
     */
    async open(source, exec, sourceCaps) {
        const trimmed = source.trim();
        if (trimmed === '')
            throw new Error('source 参数不能为空');
        if (/^https?:\/\//i.test(trimmed))
            return this.openRemote(trimmed, exec, sourceCaps);
        const absolute = this.localPath(trimmed, exec);
        exec.signal.throwIfAborted();
        const info = await stat(absolute).catch(() => null);
        if (info === null)
            throw new Error(`找不到文件：${absolute}（请确认路径正确，或改用 http(s) 链接）`);
        if (info.size > sourceCaps.maxFileBytes) {
            throw new Error(`文件过大（${Math.round(info.size / 1024 / 1024)}MB），超过上限 ${Math.round(sourceCaps.maxFileBytes / 1024 / 1024)}MB`);
        }
        const key = `local|${absolute}|${info.size}|${info.mtimeMs}`;
        return this.through(key, absolute, async () => {
            exec.signal.throwIfAborted();
            const bytes = new Uint8Array(await readFile(absolute, { signal: exec.signal }));
            return openDocument(bytes, absolute);
        });
    }
    async openRemote(url, exec, sourceCaps) {
        exec.signal.throwIfAborted();
        const resolved = await resolveSource(url, exec, sourceCaps);
        exec.signal.throwIfAborted();
        const digest = createHash('sha1').update(resolved.bytes).digest('hex');
        return this.through(`url|${resolved.displayName}|${digest}`, null, async () => openDocument(resolved.bytes, resolved.displayName));
    }
    /**
     * Read a source's raw bytes without parsing.
     *
     * `office_audit` works on the package itself (zip entries, parts, relationships), so it needs the
     * bytes rather than a parsed document. Parsed-document caching does not apply; local files are
     * read once per call, remote sources go through the same download path as everything else.
     */
    async readBytes(source, exec, sourceCaps) {
        const trimmed = source.trim();
        if (trimmed === '')
            throw new Error('source 参数不能为空');
        if (/^https?:\/\//i.test(trimmed)) {
            exec.signal.throwIfAborted();
            const resolved = await resolveSource(trimmed, exec, sourceCaps);
            return { bytes: resolved.bytes, displayName: resolved.displayName, absolute: null };
        }
        const absolute = this.localPath(trimmed, exec);
        exec.signal.throwIfAborted();
        const info = await stat(absolute).catch(() => null);
        if (info === null)
            throw new Error(`找不到文件：${absolute}（请确认路径正确，或改用 http(s) 链接）`);
        if (info.size > sourceCaps.maxFileBytes) {
            throw new Error(`文件过大（${Math.round(info.size / 1024 / 1024)}MB），超过上限 ${Math.round(sourceCaps.maxFileBytes / 1024 / 1024)}MB`);
        }
        const bytes = new Uint8Array(await readFile(absolute, { signal: exec.signal }));
        return { bytes, displayName: absolute, absolute };
    }
    async through(key, pathKey, parse) {
        const hit = this.entries.get(key);
        if (hit !== undefined) {
            hit.at = ++this.counter;
            return hit.doc;
        }
        const pending = this.inflight.get(key);
        if (pending !== undefined)
            return pending;
        const promise = parse()
            .then((doc) => {
            this.store(key, pathKey, doc);
            return doc;
        })
            .finally(() => {
            this.inflight.delete(key);
        });
        this.inflight.set(key, promise);
        return promise;
    }
    store(key, pathKey, doc) {
        // `doc.charCount` is precomputed by each parser. Iterating `doc.units` here would materialize a
        // workbook's rows on every store — exactly the cost the xlsx overview path avoids.
        const chars = doc.charCount;
        if (chars > this.caps.maxChars)
            return; // 单文档超限，不缓存
        // 同路径旧版本（mtime 已变）即时清理
        if (pathKey !== null) {
            for (const [existingKey, entry] of this.entries) {
                if (entry.pathKey === pathKey && existingKey !== key)
                    this.entries.delete(existingKey);
            }
        }
        this.entries.set(key, { key, pathKey, doc, chars, at: ++this.counter });
        let total = chars;
        while (this.entries.size > this.caps.maxEntries || total > this.caps.maxChars) {
            const oldest = this.entries.keys().next();
            if (oldest.done === true)
                break;
            const evicted = this.entries.get(oldest.value);
            this.entries.delete(oldest.value);
            if (evicted !== undefined)
                total -= evicted.chars;
            if (oldest.value === key)
                break; // 新条目自身超限时不缓存
        }
    }
    localPath(source, exec) {
        if (/^file:\/\//i.test(source))
            return fileURLToPath(source);
        const cwd = exec.agent?.session.header.cwd ?? process.cwd();
        return resolve(cwd, source);
    }
}
//# sourceMappingURL=cache.js.map