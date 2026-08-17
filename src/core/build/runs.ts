import { NS } from '../ooxml/ns.js';
import { attr, child, importNode, wEl } from '../ooxml/xml.js';
import type { StyleSheet } from '../analyze/styles.js';

/**
 * Run properties that carry meaning rather than appearance. Everything else
 * (fonts, sizes, colours, spacing) is dropped so the reference styles decide
 * how the book looks. The order matches the OOXML schema's required sequence.
 */
const KEPT_RUN_PROPS = [
  'rStyle',
  'b',
  'bCs',
  'i',
  'iCs',
  'caps',
  'smallCaps',
  'strike',
  'dstrike',
  'vanish',
  'u',
  'vertAlign',
  'rtl',
  'cs',
  'em',
  'lang',
] as const;

const EMPHASIS_ONLY = new Set(['b', 'bCs', 'i', 'iCs', 'caps', 'smallCaps', 'strike', 'dstrike', 'u', 'em']);

/** Paragraph content elements that are unwrapped rather than copied. */
const UNWRAPPED = new Set(['ins', 'moveTo', 'smartTag', 'sdtContent']);
/** Content that is dropped: rejected edits, comments, proofing marks. */
const DROPPED = new Set([
  'del',
  'moveFrom',
  'proofErr',
  'commentRangeStart',
  'commentRangeEnd',
  'permStart',
  'permEnd',
  'moveFromRangeStart',
  'moveFromRangeEnd',
  'moveToRangeStart',
  'moveToRangeEnd',
]);

export interface CopyOptions {
  targetDoc: Document;
  manuscriptStyles: StyleSheet;
  referenceStyles: StyleSheet;
  /** Keep bold/italic/etc. from the manuscript. */
  keepEmphasis: boolean;
  /** Strip tabs and spaces used as a manual first-line indent. */
  removeManualIndents: boolean;
  transformer: TextTransformer;
}

/**
 * Copy a manuscript paragraph's content into the output document, keeping the
 * words and their emphasis and discarding the manuscript's visual formatting.
 * Page breaks are dropped here because the composer controls pagination.
 */
export function copyParagraphContent(sourceP: Element, opts: CopyOptions): Element[] {
  const state = { atStart: opts.removeManualIndents };
  const out: Element[] = [];
  opts.transformer.startParagraph();
  collect(sourceP, out, opts, state);
  return out;
}

interface CopyState {
  /** Still at the run of leading whitespace that opens the paragraph. */
  atStart: boolean;
}

function collect(parent: Element, out: Element[], opts: CopyOptions, state: CopyState): void {
  for (let n = parent.firstChild; n; n = n.nextSibling) {
    if (n.nodeType !== 1) continue;
    const el = n as Element;
    if (el.namespaceURI !== NS.w) {
      out.push(importNode(opts.targetDoc, el));
      continue;
    }
    const name = el.localName ?? '';
    if (name === 'pPr' || DROPPED.has(name)) continue;
    if (UNWRAPPED.has(name)) {
      collect(el, out, opts, state);
      continue;
    }
    if (name === 'sdt') {
      const content = child(el, 'sdtContent');
      if (content) collect(content, out, opts, state);
      continue;
    }
    if (name === 'r') {
      const run = copyRun(el, opts, state);
      if (run) out.push(run);
      continue;
    }
    if (name === 'hyperlink' || name === 'fldSimple') {
      const wrapper = opts.targetDoc.createElementNS(NS.w, `w:${name}`);
      copyAttributes(el, wrapper);
      const inner: Element[] = [];
      collect(el, inner, opts, state);
      if (inner.length === 0) continue;
      for (const kid of inner) wrapper.appendChild(kid);
      out.push(wrapper);
      continue;
    }
    if (name === 'bookmarkStart' || name === 'bookmarkEnd') {
      // `_GoBack` is Word's cursor memory and has no business in the output.
      if (attr(el, 'name') === '_GoBack') continue;
      out.push(importNode(opts.targetDoc, el));
      continue;
    }
    // Equations and anything else structural: carry over untouched.
    out.push(importNode(opts.targetDoc, el));
  }
}

function copyRun(run: Element, opts: CopyOptions, state: CopyState): Element | null {
  const target = opts.targetDoc.createElementNS(NS.w, 'w:r');
  const rPr = sanitizeRunProps(child(run, 'rPr'), opts);
  if (rPr) target.appendChild(rPr);

  let emitted = false;
  for (let n = run.firstChild; n; n = n.nextSibling) {
    if (n.nodeType !== 1) continue;
    const el = n as Element;
    const name = el.localName ?? '';
    if (el.namespaceURI === NS.w) {
      if (name === 'rPr') continue;
      if (name === 'lastRenderedPageBreak') continue;
      if (name === 'br' && attr(el, 'type') === 'page') continue;
      if (name === 'tab' && state.atStart) continue;
      if (name === 't') {
        const raw = el.textContent ?? '';
        const text = state.atStart ? raw.replace(/^[\t ]+/, '') : raw;
        const transformed = opts.transformer.apply(text);
        if (transformed.length === 0) continue;
        if (transformed.trim().length > 0) state.atStart = false;
        target.appendChild(makeText(opts.targetDoc, transformed));
        emitted = true;
        continue;
      }
      if (name === 'br' || name === 'cr' || name === 'noBreakHyphen' || name === 'softHyphen') {
        state.atStart = false;
      }
      if (name === 'drawing' || name === 'pict' || name === 'object') {
        state.atStart = false;
      }
    }
    target.appendChild(importNode(opts.targetDoc, el));
    emitted = true;
  }

  return emitted ? target : null;
}

function makeText(doc: Document, text: string): Element {
  const el = doc.createElementNS(NS.w, 'w:t');
  if (text !== text.trim()) el.setAttributeNS(NS.xml, 'xml:space', 'preserve');
  el.appendChild(doc.createTextNode(text));
  return el;
}

/**
 * Rebuild `w:rPr` with only the semantic toggles. A character style that the
 * reference does not define is replaced by the emphasis it implied, so italics
 * expressed through a custom style survive.
 */
function sanitizeRunProps(source: Element | null, opts: CopyOptions): Element | null {
  if (!source) return null;
  const kept: Element[] = [];
  const doc = opts.targetDoc;

  const rStyleId = child(source, 'rStyle') ? attr(child(source, 'rStyle'), 'val') : null;
  let implied: { bold: boolean; italic: boolean; caps: boolean; smallCaps: boolean } | null = null;
  if (rStyleId) {
    if (opts.referenceStyles.has(rStyleId)) {
      kept.push(wEl(doc, 'rStyle', { val: rStyleId }));
    } else if (opts.manuscriptStyles.has(rStyleId)) {
      const props = opts.manuscriptStyles.resolve(rStyleId);
      implied = {
        bold: props.bold,
        italic: props.italic,
        caps: props.allCaps,
        smallCaps: props.smallCaps,
      };
    }
  }

  for (const name of KEPT_RUN_PROPS) {
    if (name === 'rStyle') continue;
    if (!opts.keepEmphasis && EMPHASIS_ONLY.has(name)) continue;
    const el = child(source, name);
    if (el) kept.push(importNode(doc, el));
  }

  if (implied && opts.keepEmphasis) {
    const present = new Set(kept.map((el) => el.localName));
    if (implied.bold && !present.has('b')) kept.push(wEl(doc, 'b'));
    if (implied.italic && !present.has('i')) kept.push(wEl(doc, 'i'));
    if (implied.caps && !present.has('caps')) kept.push(wEl(doc, 'caps'));
    if (implied.smallCaps && !present.has('smallCaps')) kept.push(wEl(doc, 'smallCaps'));
  }

  if (kept.length === 0) return null;
  const order = new Map(KEPT_RUN_PROPS.map((name, i) => [name as string, i]));
  kept.sort((a, b) => (order.get(a.localName ?? '') ?? 99) - (order.get(b.localName ?? '') ?? 99));

  const rPr = doc.createElementNS(NS.w, 'w:rPr');
  for (const el of kept) rPr.appendChild(el);
  return rPr;
}

function copyAttributes(source: Element, target: Element): void {
  const attrs = source.attributes;
  for (let i = 0; i < attrs.length; i++) {
    const a = attrs.item(i);
    if (!a) continue;
    if (a.namespaceURI) target.setAttributeNS(a.namespaceURI, a.name, a.value);
    else target.setAttribute(a.name, a.value);
  }
}

/**
 * Applies optional typographic cleanup while walking a paragraph's text runs.
 * State carries across runs so a quotation that starts in a bold run and ends
 * in a plain one still curls the right way.
 */
export class TextTransformer {
  private prev = '';

  constructor(
    private readonly smart: boolean,
    private readonly collapseSpaces: boolean,
  ) {}

  startParagraph(): void {
    this.prev = '';
  }

  apply(text: string): string {
    let out = text;
    if (this.collapseSpaces) {
      out = out.replace(/ {2,}/g, ' ');
      if (this.prev === ' ') out = out.replace(/^ +/, '');
    }
    if (this.smart) out = this.smarten(out);
    if (out.length > 0) this.prev = out[out.length - 1];
    return out;
  }

  private smarten(text: string): string {
    let out = text
      .replace(/(\S)---(\S)/g, '$1—$2')
      .replace(/(\S)--(\S)/g, '$1—$2')
      .replace(/ -- /g, '—')
      .replace(/\.\.\./g, '…');

    let result = '';
    let prev = this.prev;
    for (const ch of out) {
      if (ch === '"') {
        result += opensQuote(prev) ? '“' : '”';
      } else if (ch === "'") {
        if (/[A-Za-z0-9]/.test(prev)) result += '’';
        else result += opensQuote(prev) ? '‘' : '’';
      } else {
        result += ch;
      }
      prev = ch;
    }
    return result;
  }
}

function opensQuote(prev: string): boolean {
  return prev === '' || /[\s([{<—–“‘ -]/.test(prev);
}
