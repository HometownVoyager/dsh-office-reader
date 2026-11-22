/**
 * OOXML 基础层：zip 容器解包（fflate）与轻量 XML 文本提取工具。
 * @module dsh-office-reader/src/ooxml
 */
/** An unpacked OOXML (zip) package with convenient entry access. */
export interface OoxmlPackage {
    /** Zip entries keyed by their package path. */
    files: Record<string, Uint8Array>;
    /** Whether an entry exists. */
    has(name: string): boolean;
    /** Entry decoded as UTF-8 text, or null when absent. */
    text(name: string): string | null;
}
/**
 * Describe why a file could not be opened as an Office Open XML container.
 *
 * A zero-byte file is the most common cause of an unzip failure, and it is neither a legacy nor an
 * encrypted format — blaming those sends the reader looking in the wrong place. Every entry point
 * builds its message here so the separate copies cannot drift apart.
 *
 * @param bytes - the raw file bytes that failed to parse.
 * @param verb - the entry point's verb, e.g. `无法审计`.
 * @returns the user-facing message.
 */
export declare function notOoxmlMessage(bytes: Uint8Array, verb: string): string;
/**
 * Unpack an Office Open XML container.
 * @param bytes - raw document bytes.
 * @throws an Error with a user-facing Chinese message when not a valid zip container.
 */
export declare function openOoxml(bytes: Uint8Array): OoxmlPackage;
/** Dublin Core metadata from docProps/core.xml（全部可选）。 */
export interface CoreProps {
    title?: string;
    subject?: string;
    keywords?: string;
    creator?: string;
    lastModifiedBy?: string;
    created?: string;
    modified?: string;
}
/** Keep only defined metadata fields. */
export declare function cleanCoreProps(props: CoreProps): Record<string, string>;
/**
 * Extract Dublin Core metadata from docProps/core.xml.
 * @param pkg - unpacked package.
 */
export declare function readCoreProps(pkg: OoxmlPackage): CoreProps;
/** Decode the five XML predefined entities plus numeric character references. */
export declare function decodeXmlEntities(input: string): string;
/**
 * Concatenate the text of all `<tag>` runs inside an XML fragment, decoding entities and
 * resolving element-valued breaks (`<w:tab/>` → `\t`, `<w:br/>` → `\n`) **in document order**.
 *
 * Order matters: <w:t>制表</w:t><w:tab/><w:t>符</w:t> is `制表\t符`, which a "replace breaks
 * first, then collect text" pass would silently turn into `制表符`.
 *
 * @param xml - XML fragment.
 * @param tag - run tag local name (e.g. `w:t`, `a:t`).
 * @param breaks - optional element→character replacements.
 */
export declare function extractRunsText(xml: string, tag: string, breaks?: Array<[RegExp, string]>): string;
/** Text contributions inside a Word paragraph: tabs and explicit breaks. */
export declare const WORD_BREAKS: Array<[RegExp, string]>;
/**
 * One top-level Word paragraph: its exact XML slice plus the offsets it occupies in the
 * parent document, so callers can slice and splice without re-scanning.
 */
export interface WordParagraphSpan {
    /** Offset of `<w:p …>` (or `<w:p/>`) in the parent XML. */
    start: number;
    /** End offset (exclusive) of `</w:p>` (or of the self-closing tag). */
    end: number;
    /** Paragraph XML, i.e. `parent.slice(start, end)`. */
    xml: string;
    /** Visible text: `<w:t>` content with tabs and breaks resolved and entities decoded. */
    text: string;
    /** True when the paragraph carries no visible text (a page-break-only or blank paragraph). */
    empty: boolean;
}
/**
 * Split WordprocessingML into top-level `<w:p>` paragraphs.
 *
 * A paragraph nested inside another paragraph (a textbox drawing inside a run) is *not* a
 * separate unit: Word addresses it through its host paragraph, and a flattened list would make
 * the "paragraph N" numbering shared with dsh-office-editor ambiguous. Paragraphs inside table
 * cells are siblings of body paragraphs at this depth, so they stay addressable.
 *
 * Nesting is resolved by a real element walk: `<w:p>` also occurs inside table cells, whose
 * closers interleave with the host paragraph's in token order, so a depth counter cannot tell
 * "the paragraph ended" from "an inner box ended".
 *
 * This is the single traversal contract for every `段落 N` label the reader emits and every
 * `index` the editor accepts. `dsh-office-editor/src/docx.ts` mirrors it; the two plugins are
 * independent (no build-time link between them), so `tests/paragraph-contract.mjs` asserts both
 * implementations agree on the same bytes.
 *
 * @param xml - word/document.xml content.
 * @throws an Error when the fragment is not balanced.
 */
export declare function splitWordParagraphs(xml: string): WordParagraphSpan[];
/** One XML element slice found by {@link splitElements}. */
export interface ElementSpan {
    /** Offset of the open tag in the parent XML. */
    start: number;
    /** End offset (exclusive) of the close tag (or of the self-closing tag). */
    end: number;
    /** The element XML, i.e. `parent.slice(start, end)`. */
    xml: string;
}
/**
 * Split an XML fragment into the spans of one element name, skipping elements of that name nested
 * inside another one.
 *
 * This is the single traversal rule behind both the paragraph contract (`w:p`, where nested
 * textbox paragraphs must not become units) and the table/row/cell extraction added for table-row
 * reading (`w:tbl`/`w:tr`/`w:tc`, where a nested table's rows belong to the inner table).
 *
 * Nesting is resolved by a real element walk rather than a depth counter: `<w:p>` also occurs
 * inside table cells, whose closers interleave with the host paragraph's in token order.
 *
 * @param xml - the fragment to walk.
 * @param local - local element name to collect (namespace prefix ignored).
 * @throws an Error when the fragment is not balanced.
 */
export declare function splitElements(xml: string, local: string): ElementSpan[];
/** Local name of a possibly namespace-prefixed element name. */
export declare function localName(name: string): string;
/** One relationship of a `.rels` part. */
export interface Relationship {
    id: string;
    type: string;
    target: string;
    /** True for `TargetMode="External"` relationships (hyperlinks, external links). */
    external: boolean;
}
/**
 * Parse a relationships part into id → relationship.
 *
 * Used instead of SheetJS's own hyperlink parsing, which decodes non-ASCII relationship targets as
 * latin1 and turns `https://例子.测试/甲` into mojibake.
 */
export declare function parseRels(relsXml: string | null): Map<string, Relationship>;
/** Attribute value of an element's open tag, entity-decoded. */
export declare function attribute(openTag: string, name: string): string | undefined;
/**
 * Resolve a relationship target against the part that declares it.
 *
 * OOXML targets are relative to the declaring part's directory (`../comments1.xml` from
 * `xl/worksheets/sheet1.xml` means `xl/comments1.xml`), may be package-absolute (`/xl/...`), and may
 * carry a fragment. Naive concatenation produced false "missing part" reports for exactly the
 * commentary relationships a sheet declares.
 *
 * @param partName - the declaring part, e.g. `xl/worksheets/sheet1.xml`.
 * @param target - the raw `Target` attribute.
 * @returns the package-relative part name, or null for external targets.
 */
export declare function resolveRelTarget(partName: string, target: string): string | null;
/**
 * A tag-stack well-formedness check for OOXML parts.
 *
 * Deliberately not a general XML parser: it verifies that elements close in order and that nothing
 * sits outside the root, which is the damage class string-surgery editing produces.
 * @param xml - one XML part.
 * @param partName - part name used in the error message.
 * @returns the number of elements walked, or throws with the imbalance named.
 */
export declare function assertWellFormed(xml: string, partName: string): number;
/**
 * Heading level of a paragraph (1-9), from the style's outline level or its `w:pStyle`.
 * The numeric `styleId` fallback only fires for Word's own localized heading ids
 * (`Heading1`/`heading 1`/`标题 1`), never for an arbitrary numeric style id.
 * @param paragraphXml - one `<w:p>` slice.
 */
export declare function wordHeadingLevel(paragraphXml: string): number | undefined;
