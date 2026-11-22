/**
 * Source resolution for Office tools: local files (absolute, or relative to the session
 * working directory) and remote http(s) URLs (fetched into memory, never written to disk).
 * Cloned from dsh-pdf-reader/src/source.ts.
 * @module dsh-office-reader/src/source
 */
import type { ToolExecution } from '@deepseek-ai/dsh-tools';
/** Deployment caps shared by all Office tools (resolved plugin config). */
export interface SourceCaps {
    /** Maximum local file size in bytes. */
    maxFileBytes: number;
    /** Maximum downloaded body size in bytes. */
    maxDownloadBytes: number;
    /** Download timeout in milliseconds. */
    downloadTimeoutMs: number;
    /** Regex sources allowed as URL hosts; empty means any host. */
    allowedHosts: readonly string[];
}
/** A document resolved to its bytes, whatever the original input form. */
export interface ResolvedSource {
    /** Raw document bytes. */
    bytes: Uint8Array;
    /** Stable user-facing name (file path or URL). */
    displayName: string;
}
/**
 * Resolve a model-supplied source string to document bytes.
 * @param source - a local path (or file:// URL), or an http(s) URL.
 * @param exec - the tool execution supplying the session cwd and cancellation.
 * @param caps - resolved deployment caps.
 * @returns document bytes plus a display name.
 * @throws an Error with a user-facing Chinese message on any failure.
 */
export declare function resolveSource(source: string, exec: ToolExecution, caps: SourceCaps): Promise<ResolvedSource>;
