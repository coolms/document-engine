/**
 * An XML layer built for SURGERY: read a part, change the few things you mean
 * to change, and leave every other byte exactly as it was.
 *
 * ## Why not a normal XML library
 *
 * A conventional parser builds a model and a serialiser writes that model back
 * out. Anything the model does not represent is lost — and a real `.docx` is
 * full of things this engine will never model: `mc:AlternateContent` fallbacks,
 * VML shapes, chart parts, revision tracking, custom XML, twenty-three
 * namespace declarations on the root element. Round-tripping through a model
 * silently drops them, and the damage shows up as a chart that vanished from a
 * document nobody edited.
 *
 * ## How byte-exactness is guaranteed
 *
 * Every node keeps the exact slice of source it was parsed from. Serialising an
 * untouched node emits that slice verbatim; only a node that was actually
 * modified is rebuilt from its parts, and modifying one invalidates it and its
 * ancestors — nothing else. So fidelity is not something the serialiser has to
 * be careful about, it is a property of the structure: the ONLY bytes that can
 * differ are the ones under a node someone changed.
 *
 * That is what makes it safe to open a document authored in Word, edit one
 * paragraph, and hand it back.
 */

/** Base for everything in the tree. */
export abstract class XmlNode {
    /** Set by the parent when this node is attached; the root's stays null. */
    parent: XmlElement | null = null;

    /**
     * The exact source text this node was parsed from, or null once it has been
     * modified and must be rebuilt.
     */
    private pristine: string | null;

    protected constructor(pristine: string | null) {
        this.pristine = pristine;
    }

    /** Rebuild this node's text from its parts. Only called once modified. */
    protected abstract rebuild(): string;

    toString(): string {
        return null === this.pristine ? this.rebuild() : this.pristine;
    }

    /**
     * @internal Record the source this node was parsed from, once the parser
     * knows where it ended. Only the parser calls this, and only on a node it
     * has just built.
     */
    seal(source: string): void {
        this.pristine = source;
    }

    /**
     * Mark this node and every ancestor as modified.
     *
     * The walk upwards is the whole mechanism: a parent whose child changed can
     * no longer be emitted from its own source slice, but a SIBLING still can.
     */
    invalidate(): void {
        if (null === this.pristine) {
            // Already invalidated — so are the ancestors, and stopping here
            // keeps a deep edit from walking the spine once per changed node.
            return;
        }

        this.pristine = null;
        this.parent?.invalidate();
    }
}

/** Character data. CDATA sections are text too, spelled differently. */
export class XmlText extends XmlNode {
    constructor(
        private value: string,
        private readonly cdata: boolean,
        pristine: string | null = null,
    ) {
        super(pristine);
    }

    get textContent(): string {
        return this.value;
    }

    setTextContent(value: string): void {
        this.value = value;
        this.invalidate();
    }

    protected rebuild(): string {
        return this.cdata ? `<![CDATA[${this.value}]]>` : escapeText(this.value);
    }
}

/**
 * A comment, processing instruction or doctype: preserved exactly, never
 * interpreted, and contributing nothing to text content.
 */
export class XmlVerbatim extends XmlNode {
    constructor(private readonly source: string) {
        super(source);
    }

    protected rebuild(): string {
        return this.source;
    }
}

export interface XmlAttribute {
    readonly name: string;
    /** As it appears in the file, still escaped — so it re-emits identically. */
    rawValue: string;
    /** The quote character used, because a file may use either. */
    quote: string;
}

export class XmlElement extends XmlNode {
    private readonly attributes: XmlAttribute[] = [];
    private readonly childNodes: XmlNode[] = [];

    constructor(
        readonly name: string,
        private selfClosing: boolean,
        pristine: string | null = null,
    ) {
        super(pristine);
    }

    /** @internal — used by the parser while building the tree. */
    addAttribute(attribute: XmlAttribute): void {
        this.attributes.push(attribute);
    }

    /** @internal — used by the parser while building the tree. */
    appendChild(node: XmlNode): void {
        node.parent = this;
        this.childNodes.push(node);
    }

    get children(): readonly XmlNode[] {
        return this.childNodes;
    }

    /** Direct child elements, optionally filtered by qualified name. */
    elements(name?: string): XmlElement[] {
        const out: XmlElement[] = [];
        for (const child of this.childNodes) {
            if (child instanceof XmlElement && (undefined === name || child.name === name)) {
                out.push(child);
            }
        }

        return out;
    }

    /** First direct child element with this name, or null. */
    element(name: string): XmlElement | null {
        for (const child of this.childNodes) {
            if (child instanceof XmlElement && child.name === name) {
                return child;
            }
        }

        return null;
    }

    /**
     * Every descendant element with this name, in document order.
     *
     * Depth-first and pre-order, so `descendants('w:p')` returns paragraphs in
     * reading order — which is the order layout needs them in.
     */
    descendants(name: string): XmlElement[] {
        const out: XmlElement[] = [];
        const visit = (element: XmlElement): void => {
            for (const child of element.childNodes) {
                if (child instanceof XmlElement) {
                    if (child.name === name) {
                        out.push(child);
                    }
                    visit(child);
                }
            }
        };
        visit(this);

        return out;
    }

    attribute(name: string): string | null {
        for (const attribute of this.attributes) {
            if (attribute.name === name) {
                return unescapeXml(attribute.rawValue);
            }
        }

        return null;
    }

    /**
     * Set an attribute, keeping an existing one in its original POSITION.
     *
     * Position matters for a diff, not for correctness: moving an attribute to
     * the end would make a one-value edit show up as a rewritten element in
     * every review of the resulting file.
     */
    setAttribute(name: string, value: string): void {
        const escaped = escapeAttribute(value);
        for (const attribute of this.attributes) {
            if (attribute.name === name) {
                if (attribute.rawValue !== escaped) {
                    attribute.rawValue = escaped;
                    // The escaping above protects `"` and not `'`, so the value
                    // is only safe inside double quotes. An attribute the file
                    // wrote single-quoted has to switch, or a value containing
                    // an apostrophe would end the attribute early and produce a
                    // file that no longer parses.
                    attribute.quote = '"';
                    this.invalidate();
                }

                return;
            }
        }

        this.attributes.push({ name, rawValue: escaped, quote: '"' });
        this.invalidate();
    }

    removeAttribute(name: string): void {
        const index = this.attributes.findIndex((attribute) => attribute.name === name);
        if (index >= 0) {
            this.attributes.splice(index, 1);
            this.invalidate();
        }
    }

    /**
     * All text beneath this element, concatenated — XPath's string-value.
     *
     * Descendants rather than direct children because a paragraph's text lives
     * two levels down, inside its runs.
     */
    get text(): string {
        let out = '';
        const visit = (element: XmlElement): void => {
            for (const child of element.childNodes) {
                if (child instanceof XmlText) {
                    out += child.textContent;
                } else if (child instanceof XmlElement) {
                    visit(child);
                }
            }
        };
        visit(this);

        return out;
    }

    /**
     * Replace this element's content with a single text node.
     *
     * Nothing clears `selfClosing` here on purpose: {@link rebuild} refuses to
     * self-close an element that has children, so the two cannot disagree. A
     * flag cleared at every mutation site is a rule that holds until someone
     * adds a site and forgets — losing the content silently, since a
     * self-closing tag is perfectly valid XML.
     */
    setText(value: string): void {
        this.childNodes.length = 0;
        this.appendChild(new XmlText(value, false));
        this.invalidate();
    }

    protected rebuild(): string {
        let out = `<${this.name}`;
        for (const attribute of this.attributes) {
            out += ` ${attribute.name}=${attribute.quote}${attribute.rawValue}${attribute.quote}`;
        }

        // An element with children can never self-close, whatever the flag
        // says. Enforcing it HERE rather than at each mutation site is what
        // lets setText and friends ignore the flag entirely.
        if (this.selfClosing && 0 === this.childNodes.length) {
            return `${out}/>`;
        }

        out += '>';
        for (const child of this.childNodes) {
            out += child.toString();
        }

        return `${out}</${this.name}>`;
    }
}

export class XmlDocument {
    private constructor(
        /** Everything before the root: the declaration, comments, doctype. */
        private readonly prolog: string,
        readonly root: XmlElement,
        private readonly epilog: string,
    ) {}

    static parse(source: string): XmlDocument {
        const { prolog, root, epilog } = new Parser(source).parseDocument();

        return new XmlDocument(prolog, root, epilog);
    }

    toString(): string {
        return this.prolog + this.root.toString() + this.epilog;
    }
}

const NAME_PATTERN = /[A-Za-z_:][A-Za-z0-9_:.\-]*/y;
const WHITESPACE_PATTERN = /[ \t\r\n]+/y;

class Parser {
    private index = 0;

    constructor(private readonly source: string) {}

    parseDocument(): { prolog: string; root: XmlElement; epilog: string } {
        const start = this.index;
        this.skipProlog();
        const prolog = this.source.slice(start, this.index);

        if ('<' !== this.at()) {
            throw this.fail('expected a root element');
        }

        const root = this.parseElement();

        return { prolog, root, epilog: this.source.slice(this.index) };
    }

    /** Whitespace, the XML declaration, comments and the doctype. */
    private skipProlog(): void {
        for (;;) {
            if (this.skipWhitespace()) {
                continue;
            }
            if (this.looking('<?')) {
                this.consumeUntil('?>', 'unterminated processing instruction');
                continue;
            }
            if (this.looking('<!--')) {
                this.consumeUntil('-->', 'unterminated comment');
                continue;
            }
            if (this.looking('<!')) {
                // A doctype. Ours never have internal subsets, and a document
                // that did would be rejected here rather than mis-parsed.
                this.consumeUntil('>', 'unterminated declaration');
                continue;
            }

            return;
        }
    }

    private parseElement(): XmlElement {
        const start = this.index;
        this.index++; // '<'

        const name = this.readName('element name');
        const attributes: XmlAttribute[] = [];
        let selfClosing = false;

        for (;;) {
            this.skipWhitespace();

            if (this.looking('/>')) {
                this.index += 2;
                selfClosing = true;
                break;
            }
            if ('>' === this.at()) {
                this.index++;
                break;
            }
            if ('' === this.at()) {
                throw this.fail(`unterminated start tag for <${name}>`);
            }

            attributes.push(this.parseAttribute());
        }

        const element = new XmlElement(name, selfClosing);
        for (const attribute of attributes) {
            element.addAttribute(attribute);
        }

        if (!selfClosing) {
            this.parseChildren(element, name);
        }

        // Sealed LAST, so the recorded span covers the end tag too.
        element.seal(this.source.slice(start, this.index));

        return element;
    }

    private parseChildren(element: XmlElement, name: string): void {
        for (;;) {
            if ('' === this.at()) {
                throw this.fail(`unclosed <${name}>`);
            }

            if (this.looking('</')) {
                this.index += 2;
                const closing = this.readName('end tag name');
                if (closing !== name) {
                    throw this.fail(`</${closing}> closes <${name}>`);
                }
                this.skipWhitespace();
                if ('>' !== this.at()) {
                    throw this.fail(`unterminated end tag </${closing}>`);
                }
                this.index++;

                return;
            }

            element.appendChild(this.parseChild());
        }
    }

    private parseChild(): XmlNode {
        const start = this.index;

        if (this.looking('<!--')) {
            this.consumeUntil('-->', 'unterminated comment');

            return new XmlVerbatim(this.source.slice(start, this.index));
        }
        if (this.looking('<?')) {
            this.consumeUntil('?>', 'unterminated processing instruction');

            return new XmlVerbatim(this.source.slice(start, this.index));
        }
        if (this.looking('<![CDATA[')) {
            this.index += '<![CDATA['.length;
            const contentStart = this.index;
            this.consumeUntil(']]>', 'unterminated CDATA section');
            const value = this.source.slice(contentStart, this.index - ']]>'.length);

            return new XmlText(value, true, this.source.slice(start, this.index));
        }
        if ('<' === this.at()) {
            return this.parseElement();
        }

        const next = this.source.indexOf('<', this.index);
        this.index = next < 0 ? this.source.length : next;
        const raw = this.source.slice(start, this.index);

        return new XmlText(unescapeXml(raw), false, raw);
    }

    private parseAttribute(): XmlAttribute {
        const name = this.readName('attribute name');
        this.skipWhitespace();
        if ('=' !== this.at()) {
            throw this.fail(`attribute ${name} has no value`);
        }
        this.index++;
        this.skipWhitespace();

        const quote = this.at();
        if ('"' !== quote && "'" !== quote) {
            throw this.fail(`attribute ${name} is not quoted`);
        }
        this.index++;

        const end = this.source.indexOf(quote, this.index);
        if (end < 0) {
            throw this.fail(`unterminated value for attribute ${name}`);
        }
        const rawValue = this.source.slice(this.index, end);
        this.index = end + 1;

        return { name, rawValue, quote };
    }

    private readName(what: string): string {
        NAME_PATTERN.lastIndex = this.index;
        const match = NAME_PATTERN.exec(this.source);
        if (null === match) {
            throw this.fail(`expected ${what}`);
        }
        this.index = NAME_PATTERN.lastIndex;

        return match[0];
    }

    private skipWhitespace(): boolean {
        WHITESPACE_PATTERN.lastIndex = this.index;
        const match = WHITESPACE_PATTERN.exec(this.source);
        if (null === match) {
            return false;
        }
        this.index = WHITESPACE_PATTERN.lastIndex;

        return true;
    }

    private consumeUntil(terminator: string, message: string): void {
        const end = this.source.indexOf(terminator, this.index);
        if (end < 0) {
            throw this.fail(message);
        }
        this.index = end + terminator.length;
    }

    private looking(prefix: string): boolean {
        return this.source.startsWith(prefix, this.index);
    }

    private at(offset = 0): string {
        return this.source.charAt(this.index + offset);
    }

    /**
     * Errors carry a line and column, because "invalid XML" about a 48 000
     * character single-line part is not a diagnosis.
     */
    private fail(message: string): Error {
        const before = this.source.slice(0, this.index);
        const line = before.split('\n').length;
        const column = this.index - (before.lastIndexOf('\n') + 1) + 1;

        return new Error(`Malformed XML at line ${line}, column ${column}: ${message}`);
    }
}

const NAMED_ENTITIES = new Map<string, string>([
    ['amp', '&'],
    ['lt', '<'],
    ['gt', '>'],
    ['quot', '"'],
    ['apos', "'"],
]);

/**
 * Resolve entity references.
 *
 * XML defines only the five below; anything else must be declared in a doctype,
 * which OOXML never has. An unknown reference is therefore a corrupt file or a
 * mis-encoded write, and throwing says so at the point of damage rather than
 * letting `&nbsp;` reach a renderer as literal text.
 */
export function unescapeXml(text: string): string {
    if (!text.includes('&')) {
        return text;
    }

    return text.replace(/&(#x?[0-9A-Fa-f]+|[A-Za-z]+);/g, (match, body: string) => {
        if (body.startsWith('#')) {
            const hex = 'x' === body[1] || 'X' === body[1];
            const code = Number.parseInt(hex ? body.slice(2) : body.slice(1), hex ? 16 : 10);
            if (!Number.isFinite(code) || code < 0 || code > 0x10ffff) {
                throw new Error(`Invalid character reference ${match}`);
            }

            return String.fromCodePoint(code);
        }

        const resolved = NAMED_ENTITIES.get(body);
        if (undefined === resolved) {
            throw new Error(`Undefined entity ${match}; XML predefines only amp, lt, gt, quot and apos.`);
        }

        return resolved;
    });
}

/**
 * Escape text content.
 *
 * `>` is escaped although it is only strictly required after `]]`. Escaping it
 * always is what Word does, and matching the producer keeps a document edited
 * here from showing spurious differences against the same document saved there.
 */
export function escapeText(text: string): string {
    return text.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

/**
 * Escape an attribute value for DOUBLE quotes.
 *
 * Tabs and newlines become character references because an XML processor
 * normalises literal whitespace in an attribute value to a space — a run's
 * `w:val` containing a newline would silently become a different value on the
 * next read.
 */
export function escapeAttribute(value: string): string {
    return escapeText(value)
        .replace(/"/g, '&quot;')
        .replace(/\n/g, '&#10;')
        .replace(/\r/g, '&#13;')
        .replace(/\t/g, '&#9;');
}
