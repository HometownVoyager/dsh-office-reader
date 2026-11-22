/**
 * .pptx 解析：幻灯片顺序（`p:sldIdLst`）、标题占位符识别、文本提取与演讲者备注。
 *
 * 顺序与备注都必须走 PresentationML 的登记关系，而**不是**文件名数字：
 * `ppt/slides/slideN.xml` 的编号是创建时分配的，重排幻灯片时不会重新编号。
 * 按文件名排序会把重排过的 deck 读错顺序（并让 `slides="3-5"` 返回错页），
 * 备注按 `notesSlide{index}.xml` 拼路径同样会张冠李戴。
 * @module dsh-office-reader/src/pptx
 */
import { type CoreProps } from './ooxml.js';
/** One slide with its title, text (notes appended) and speaker notes. */
export interface PptxSlide {
    /** 1-based slide number in presentation order. */
    index: number;
    /** Package path of the slide part, e.g. `ppt/slides/slide3.xml`. */
    part: string;
    /** Slide title (title placeholder text, else first paragraph). */
    title: string;
    /** Slide body text with notes appended as `[备注]` lines. */
    text: string;
    /** Speaker notes without the appended marker. */
    notes?: string;
    /** Text length in characters. */
    chars: number;
}
/** A shape's text as read from the slide (used when a caller needs structure). */
export interface PptxShapeText {
    /** `p:cNvPr/@name` when present. */
    name?: string;
    /** Placeholder type (`title`/`ctrTitle`/`body`/`sldNum`/…), absent for plain shapes. */
    placeholder?: string;
    /** True for shapes that only carry a template decoration (slide number, date, footer). */
    decorative: boolean;
    /** Paragraph texts, one entry per non-empty `<a:p>`. */
    paragraphs: string[];
}
/** Parsed .pptx document. */
export interface PptxDocument {
    format: 'pptx';
    slides: PptxSlide[];
    metadata: CoreProps;
}
/**
 * Parse a .pptx presentation into slides in **presentation order**.
 * @param bytes - raw document bytes.
 * @throws an Error with a user-facing Chinese message when not a valid pptx.
 */
export declare function readPptx(bytes: Uint8Array): PptxDocument;
/**
 * Extract every shape's text from a slide, keeping placeholder identity.
 *
 * Reading shape-by-shape (rather than flattening every `<a:p>` in the part) is what keeps table
 * cells from being read as loose slide text and lets decorative placeholders be filtered out —
 * previously the date/footer/slide-number placeholder text leaked into both the slide body and
 * the speaker notes and therefore into `office_find` results.
 */
export declare function slideShapes(xml: string): PptxShapeText[];
/** Split a fragment into its top-level `<name …>…</name>` elements (nesting aware). */
export declare function splitElements(xml: string, name: string): string[];
