/**
 * Keyword search over rebuilt unit text: NFKC + case folding normalization, and
 * whitespace-insensitive matching. Cloned from dsh-pdf-reader/src/search.ts.
 * @module dsh-office-reader/src/search
 */
/**
 * Normalize both sides for matching: NFKC (full/half-width unification) and
 * case folding, with all whitespace removed.
 * @param input - raw text or query.
 * @returns the normalized comparison string.
 */
export declare function normalizeForMatch(input: string): string;
/** Search outcome for one unit. */
export interface UnitSearch {
    /** 1-based unit index. */
    index: number;
    /** Number of matches found in this unit. */
    count: number;
    /** Up to the per-unit cap of snippets. */
    snippets: string[];
}
/**
 * Search one unit's text for a normalized query with whitespace-insensitive
 * matching, returning original-text snippets.
 * @param unitText - the rebuilt unit text.
 * @param queryNorm - query already normalized by {@link normalizeForMatch}.
 * @param contextChars - characters of context to keep on each side.
 * @param maxSnippets - maximum snippets returned per unit.
 * @returns match count and original-text snippets.
 */
export declare function searchUnit(unitText: string, queryNorm: string, contextChars: number, maxSnippets: number): Omit<UnitSearch, 'index'>;
/** Convenience wrapper carrying the unit index. */
export declare function searchUnitFor(index: number, unitText: string, queryNorm: string, contextChars: number, maxSnippets: number): UnitSearch;
/**
 * Compile a caller-supplied pattern, turning a syntax error into a user-facing message.
 *
 * `caseSensitive` decides the `i` flag. It is a parameter rather than something the caller writes
 * into the pattern because JavaScript has **no inline flag syntax**: `(?i)alpha` — the natural
 * Python/PCRE spelling — is a syntax error, and `new RegExp(..., 'gu')` would reject it with a
 * message about an invalid group. The flag is therefore the only way to express "ignore case" here.
 *
 * @param pattern - the caller's pattern.
 * @param caseSensitive - when false (the default) matching ignores case.
 * @returns the compiled global pattern.
 */
export declare function compilePattern(pattern: string, caseSensitive?: boolean): RegExp;
/**
 * Search one unit with a caller-supplied regular expression.
 *
 * Regex mode deliberately **skips the normalization pipeline** {@link searchUnit} applies (NFKC
 * folding, lowercasing, whitespace stripping). Those transforms rewrite the haystack, which would
 * invalidate every offset the pattern's own anchors and character classes depend on — and
 * `\s`-free patterns are exactly what a caller reaches for regex mode to express. The pattern is
 * therefore matched against the unit text as written: whitespace-sensitive always, and
 * case-sensitive only when the caller asked for that (see {@link compilePattern} for why the `i`
 * flag cannot be written into the pattern).
 *
 * Two guard rails, both because a caller pattern is untrusted input:
 *
 * - zero-width matches (`^`, `\b`, `a*`) are counted but never advance the scan past their own
 *   start, so `lastIndex` is bumped by hand — `String.prototype.matchAll` loops forever on them;
 * - counting stops at `maxMatches`, which bounds the work a single unit can cost. There is no
 *   time-based escape (JavaScript cannot interrupt a running regex), so the remaining exposure is
 *   catastrophic backtracking within one unit; unit texts here are one paragraph or one
 *   spreadsheet row, not a whole document.
 *
 * @param unitText - the rebuilt unit text.
 * @param pattern - compiled caller pattern (global).
 * @param contextChars - characters of context to keep on each side.
 * @param maxSnippets - maximum snippets returned per unit.
 * @param maxMatches - stop counting at this many hits.
 * @returns match count and original-text snippets (empty for zero-width hits), plus whether the
 *   count hit `maxMatches`.
 */
export declare function searchUnitRegex(unitText: string, pattern: RegExp, contextChars: number, maxSnippets: number, maxMatches: number): Omit<UnitSearch, 'index'> & {
    capped: boolean;
};
