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
export function normalizeForMatch(input) {
    return input.normalize('NFKC').toLowerCase().replace(/\s+/gu, '');
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
export function searchUnit(unitText, queryNorm, contextChars, maxSnippets) {
    if (queryNorm === '' || unitText === '')
        return { count: 0, snippets: [] };
    // Map every original character to its normalized form; strip whitespace into a
    // parallel index so match offsets translate back to original positions.
    const normChars = [];
    const originIndices = [];
    for (let i = 0; i < unitText.length; i++) {
        const ch = unitText[i];
        if (/\s/u.test(ch))
            continue;
        normChars.push(ch.normalize('NFKC').toLowerCase());
        originIndices.push(i);
    }
    const stripped = normChars.join('');
    let count = 0;
    const snippets = [];
    let from = 0;
    while (from + queryNorm.length <= stripped.length) {
        const at = stripped.indexOf(queryNorm, from);
        if (at === -1)
            break;
        count++;
        const start = originIndices[at];
        const endOrigin = at + queryNorm.length - 1 < originIndices.length
            ? originIndices[at + queryNorm.length - 1]
            : unitText.length;
        const windowStart = Math.max(0, start - contextChars);
        const windowEnd = Math.min(unitText.length, endOrigin + contextChars);
        let snippet = unitText.slice(windowStart, windowEnd).replace(/\s+/gu, ' ');
        if (windowStart > 0)
            snippet = `…${snippet}`;
        if (windowEnd < unitText.length)
            snippet = `${snippet}…`;
        if (snippets.length < maxSnippets)
            snippets.push(snippet);
        from = at + queryNorm.length;
    }
    return { count, snippets };
}
/** Convenience wrapper carrying the unit index. */
export function searchUnitFor(index, unitText, queryNorm, contextChars, maxSnippets) {
    return { index, ...searchUnit(unitText, queryNorm, contextChars, maxSnippets) };
}
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
export function compilePattern(pattern, caseSensitive = false) {
    try {
        return new RegExp(pattern, caseSensitive ? 'gu' : 'giu');
    }
    catch (error) {
        throw new Error(`正则表达式无效：${error instanceof Error ? error.message : String(error)}`);
    }
}
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
export function searchUnitRegex(unitText, pattern, contextChars, maxSnippets, maxMatches) {
    if (unitText === '')
        return { count: 0, snippets: [], capped: false };
    const snippets = [];
    let count = 0;
    pattern.lastIndex = 0;
    for (let m = pattern.exec(unitText); m !== null; m = pattern.exec(unitText)) {
        count++;
        const matched = m[0];
        if (matched !== '' && snippets.length < maxSnippets) {
            const start = m.index;
            const end = m.index + matched.length;
            const windowStart = Math.max(0, start - contextChars);
            const windowEnd = Math.min(unitText.length, end + contextChars);
            let snippet = unitText.slice(windowStart, windowEnd).replace(/\s+/gu, ' ');
            if (windowStart > 0)
                snippet = `…${snippet}`;
            if (windowEnd < unitText.length)
                snippet = `${snippet}…`;
            snippets.push(snippet);
        }
        if (count >= maxMatches)
            return { count, snippets, capped: true };
        // A zero-width match would leave `lastIndex` where it started: step over one character so the
        // scan always progresses, which is the same rule `String.prototype.replaceAll` uses.
        if (matched === '')
            pattern.lastIndex = m.index + 1;
    }
    return { count, snippets, capped: false };
}
//# sourceMappingURL=search.js.map