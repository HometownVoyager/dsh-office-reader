/**
 * .pptx 解析：幻灯片顺序（`p:sldIdLst`）、标题占位符识别、文本提取与演讲者备注。
 *
 * 顺序与备注都必须走 PresentationML 的登记关系，而**不是**文件名数字：
 * `ppt/slides/slideN.xml` 的编号是创建时分配的，重排幻灯片时不会重新编号。
 * 按文件名排序会把重排过的 deck 读错顺序（并让 `slides="3-5"` 返回错页），
 * 备注按 `notesSlide{index}.xml` 拼路径同样会张冠李戴。
 * @module dsh-office-reader/src/pptx
 */
import { openOoxml, readCoreProps, extractRunsText } from './ooxml.js';
const BREAKS = [
    [/^<a:br\b/u, '\n'],
    [/^<a:tab\b/u, '\t'],
];
/**
 * Parse a .pptx presentation into slides in **presentation order**.
 * @param bytes - raw document bytes.
 * @throws an Error with a user-facing Chinese message when not a valid pptx.
 */
export function readPptx(bytes) {
    const pkg = openOoxml(bytes);
    const parts = slidePartsInOrder(pkg);
    if (parts.length === 0)
        throw new Error('不是有效的 .pptx 演示文稿（缺少 ppt/slides 或 ppt/presentation.xml 的 sldIdLst）');
    const slides = [];
    for (const [i, part] of parts.entries()) {
        const xml = pkg.text(part) ?? '';
        const shapes = slideShapes(xml);
        const title = slideTitle(shapes);
        const notes = readNotes(pkg, part);
        const lines = shapes.filter((shape) => !shape.decorative).flatMap((shape) => shape.paragraphs);
        if (notes !== undefined)
            lines.push(`[备注] ${notes}`);
        const text = lines.join('\n');
        slides.push({ index: i + 1, part, title, text, notes, chars: text.length });
    }
    return { format: 'pptx', slides, metadata: readCoreProps(pkg) };
}
/** Slide part paths in presentation order, from `ppt/presentation.xml`'s `p:sldIdLst`. */
function slidePartsInOrder(pkg) {
    const presentation = pkg.text('ppt/presentation.xml');
    const rels = readRels(pkg, 'ppt/_rels/presentation.xml.rels');
    if (presentation !== null) {
        const ids = [...presentation.matchAll(/<p:sldId\b[^>]*r:id="([^"]+)"/gu)].map((m) => m[1]);
        const ordered = [];
        for (const id of ids) {
            const target = rels.get(id);
            if (target === undefined || !/^slides\/slide\d+\.xml$/u.test(target))
                continue;
            const normalized = `ppt/${target}`;
            if (pkg.has(normalized) && !ordered.includes(normalized))
                ordered.push(normalized);
        }
        if (ordered.length > 0)
            return ordered;
    }
    // Fallback for a deck whose sldIdLst is unreadable: every slide part, in part-number order.
    // Reported as a readable order rather than thrown: a partially damaged deck is still useful.
    return Object.keys(pkg.files)
        .filter((name) => /^ppt\/slides\/slide\d+\.xml$/u.test(name))
        .sort((a, b) => partNumber(a) - partNumber(b));
}
function partNumber(path) {
    return Number(/(\d+)\.xml$/u.exec(path)?.[1] ?? 0);
}
/** Parse a `.rels` part into an rId → target map. */
function readRels(pkg, relsPath) {
    const out = new Map();
    const xml = pkg.text(relsPath);
    if (xml === null)
        return out;
    for (const m of xml.matchAll(/<Relationship\b[^>]*\/>/gu)) {
        const id = /Id="([^"]+)"/u.exec(m[0])?.[1];
        const target = /Target="([^"]+)"/u.exec(m[0])?.[1];
        const mode = /TargetMode="External"/u.test(m[0]);
        if (id !== undefined && target !== undefined && !mode)
            out.set(id, target);
    }
    return out;
}
/** Placeholder types that only carry template chrome rather than authored content. */
const DECORATIVE_PLACEHOLDERS = new Set(['sldNum', 'sldImg', 'dt', 'ftr']);
/**
 * Extract every shape's text from a slide, keeping placeholder identity.
 *
 * Reading shape-by-shape (rather than flattening every `<a:p>` in the part) is what keeps table
 * cells from being read as loose slide text and lets decorative placeholders be filtered out —
 * previously the date/footer/slide-number placeholder text leaked into both the slide body and
 * the speaker notes and therefore into `office_find` results.
 */
export function slideShapes(xml) {
    const shapes = [];
    for (const shapeXml of splitElements(xml, 'p:sp')) {
        const name = /<p:cNvPr\b[^>]*name="([^"]*)"/u.exec(shapeXml)?.[1];
        const placeholder = /<p:ph\b[^>]*\btype="([^"]*)"/u.exec(shapeXml)?.[1];
        const decorative = placeholder !== undefined && DECORATIVE_PLACEHOLDERS.has(placeholder);
        const paragraphs = paragraphTexts(shapeXml);
        shapes.push({
            ...(name !== undefined ? { name } : {}),
            ...(placeholder !== undefined ? { placeholder } : {}),
            decorative,
            paragraphs,
        });
    }
    return shapes;
}
/** Non-empty `<a:p>` texts inside a fragment, in document order. */
function paragraphTexts(xml) {
    const out = [];
    for (const paragraph of splitElements(xml, 'a:p')) {
        // Fields (`a:fld`, e.g. a slide number or a date) carry a cached value; it is template chrome,
        // not authored text, so it is dropped rather than mistaken for a sentence.
        const withoutFields = paragraph.replace(/<a:fld\b[^>]*>[\s\S]*?<\/a:fld>/gu, '');
        const text = extractRunsText(withoutFields, 'a:t', BREAKS).replace(/[ \t]+/gu, ' ').trim();
        if (text !== '')
            out.push(text);
    }
    return out;
}
/** Split a fragment into its top-level `<name …>…</name>` elements (nesting aware). */
export function splitElements(xml, name) {
    const out = [];
    const escaped = name.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    const re = new RegExp(`<${escaped}(?=\\s|/?>)(?:"[^"]*"|'[^']*'|[^>"'])*?(/?)>|</${escaped}\\s*>`, 'gu');
    let depth = 0;
    let start = -1;
    for (let m = re.exec(xml); m !== null; m = re.exec(xml)) {
        const token = m[0];
        if (token.startsWith('</')) {
            depth--;
            if (depth === 0 && start >= 0) {
                out.push(xml.slice(start, m.index + token.length));
                start = -1;
            }
            continue;
        }
        if (m[1] === '/') {
            if (depth === 0)
                out.push(token);
            continue;
        }
        if (depth === 0)
            start = m.index;
        depth++;
    }
    return out;
}
/** Title from a `type="title"`/`ctrTitle` placeholder shape; fallback to the first body paragraph. */
function slideTitle(shapes) {
    for (const shape of shapes) {
        if (shape.placeholder === 'title' || shape.placeholder === 'ctrTitle') {
            const first = shape.paragraphs.find((p) => p !== '');
            if (first !== undefined)
                return first;
        }
    }
    for (const shape of shapes) {
        if (shape.decorative)
            continue;
        const first = shape.paragraphs.find((p) => p !== '');
        if (first !== undefined)
            return first;
    }
    return '';
}
/**
 * Speaker notes for a slide, resolved through the slide's own relationships.
 *
 * `ppt/notesSlides/notesSlideN.xml` is the *slide part's* notes target only by coincidence: the
 * numbering of the two families is independent, so the old `notesSlide{slideNumber}` guess showed
 * one slide's script on another after any reorder.
 */
function readNotes(pkg, slidePart) {
    const relsPath = `${slidePart.slice(0, slidePart.lastIndexOf('/'))}/_rels/${slidePart.slice(slidePart.lastIndexOf('/') + 1)}.rels`;
    const rels = readRels(pkg, relsPath);
    let notesPart;
    for (const target of rels.values()) {
        const normalized = normalizePart(slidePart, target);
        if (/^ppt\/notesSlides\/notesSlide\d+\.xml$/u.test(normalized) && pkg.has(normalized)) {
            notesPart = normalized;
            break;
        }
    }
    if (notesPart === undefined)
        return undefined;
    const xml = pkg.text(notesPart);
    if (xml === null)
        return undefined;
    const lines = slideShapes(xml)
        .filter((shape) => !shape.decorative)
        .flatMap((shape) => shape.paragraphs)
        // A slide-number placeholder that is not typed as one still renders as a bare number.
        .filter((line) => !/^\d+$/u.test(line));
    const text = lines.join('\n').trim();
    return text === '' ? undefined : text;
}
/** Resolve a relationship target against the part that owns the rels file. */
function normalizePart(ownerPart, target) {
    if (target.startsWith('/'))
        return target.slice(1);
    const base = ownerPart.slice(0, ownerPart.lastIndexOf('/'));
    const segments = `${base}/${target}`.split('/');
    const resolved = [];
    for (const segment of segments) {
        if (segment === '' || segment === '.')
            continue;
        if (segment === '..')
            resolved.pop();
        else
            resolved.push(segment);
    }
    return resolved.join('/');
}
//# sourceMappingURL=pptx.js.map