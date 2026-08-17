import { DOMParser, XMLSerializer } from '@xmldom/xmldom';
import { NS } from './ns.js';

const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n';

/**
 * Parse an OOXML part. Errors and warnings from the underlying parser are
 * collected rather than printed, so a slightly malformed manuscript still
 * loads instead of spewing to stderr.
 */
export function parseXml(text: string): Document {
  const problems: string[] = [];
  const parser = new DOMParser({
    onError: (level, message) => {
      if (level === 'fatalError') problems.push(message);
    },
  });
  // Strip a UTF-8 BOM: xmldom treats it as content and fails the declaration.
  const doc = parser.parseFromString(text.replace(/^﻿/, ''), 'text/xml') as unknown as Document;
  if (problems.length > 0) {
    throw new Error(`Malformed XML part: ${problems[0]}`);
  }
  return doc;
}

export function serializeXml(doc: Document | Element): string {
  const body = new XMLSerializer().serializeToString(doc as never);
  return body.startsWith('<?xml') ? body : XML_DECL + body;
}

/** Direct element children of `parent` matching `localName` in namespace `ns`. */
export function children(parent: Node | null, localName: string, ns: string = NS.w): Element[] {
  if (!parent) return [];
  const out: Element[] = [];
  for (let n = parent.firstChild; n; n = n.nextSibling) {
    if (n.nodeType === 1) {
      const el = n as Element;
      if (el.localName === localName && el.namespaceURI === ns) out.push(el);
    }
  }
  return out;
}

/** First direct element child matching `localName`, or null. */
export function child(parent: Node | null, localName: string, ns: string = NS.w): Element | null {
  if (!parent) return null;
  for (let n = parent.firstChild; n; n = n.nextSibling) {
    if (n.nodeType === 1) {
      const el = n as Element;
      if (el.localName === localName && el.namespaceURI === ns) return el;
    }
  }
  return null;
}

/** All element children at any depth matching `localName`, in document order. */
export function descendants(parent: Node | null, localName: string, ns: string = NS.w): Element[] {
  if (!parent) return [];
  const out: Element[] = [];
  const walk = (node: Node): void => {
    for (let n = node.firstChild; n; n = n.nextSibling) {
      if (n.nodeType !== 1) continue;
      const el = n as Element;
      if (el.localName === localName && el.namespaceURI === ns) out.push(el);
      walk(el);
    }
  };
  walk(parent);
  return out;
}

/** First descendant matching `localName`, or null. */
export function firstDescendant(
  parent: Node | null,
  localName: string,
  ns: string = NS.w,
): Element | null {
  return descendants(parent, localName, ns)[0] ?? null;
}

/**
 * Read a namespaced attribute. Falls back to the `w:`-prefixed literal name
 * because some producers write attributes without a namespace declaration in
 * scope, which leaves xmldom's namespaceURI null.
 */
export function attr(el: Element | null, localName: string, ns: string = NS.w): string | null {
  if (!el) return null;
  const nsValue = el.getAttributeNS(ns, localName);
  if (nsValue !== null && nsValue !== '') return nsValue;
  const prefix = ns === NS.w ? 'w' : ns === NS.r ? 'r' : ns === NS.xml ? 'xml' : null;
  if (prefix) {
    const literal = el.getAttribute(`${prefix}:${localName}`);
    if (literal !== null && literal !== '') return literal;
  }
  const bare = el.getAttribute(localName);
  return bare !== null && bare !== '' ? bare : null;
}

/** `w:val` of a direct child element, e.g. `pStyle`'s style id. */
export function childVal(
  parent: Node | null,
  localName: string,
  ns: string = NS.w,
): string | null {
  return attr(child(parent, localName, ns), 'val');
}

/**
 * Whether a boolean toggle child (`w:b`, `w:i`, ...) is on. In OOXML a bare
 * element means true; `w:val="0"`/`"false"` turns it off.
 */
export function toggleOn(parent: Node | null, localName: string): boolean {
  const el = child(parent, localName);
  if (!el) return false;
  const v = attr(el, 'val');
  return v === null || v === '1' || v === 'true' || v === 'on';
}

/** Numeric attribute value, or null when absent/unparseable. */
export function numAttr(el: Element | null, localName: string, ns: string = NS.w): number | null {
  const raw = attr(el, localName, ns);
  if (raw === null) return null;
  const n = Number(raw);
  return Number.isFinite(n) ? n : null;
}

/** Create a `w:`-namespaced element with optional `w:`-namespaced attributes. */
export function wEl(
  doc: Document,
  localName: string,
  attrs: Record<string, string | number | undefined> = {},
): Element {
  const el = doc.createElementNS(NS.w, `w:${localName}`);
  for (const [key, value] of Object.entries(attrs)) {
    if (value === undefined) continue;
    el.setAttributeNS(NS.w, `w:${key}`, String(value));
  }
  return el;
}

/** Append children to an element and return it, for terse tree building. */
export function append<T extends Element>(parent: T, ...kids: (Node | null | undefined)[]): T {
  for (const kid of kids) if (kid) parent.appendChild(kid);
  return parent;
}

/**
 * Text of a paragraph or run container. Tabs become `\t`, line breaks `\n`,
 * so classifiers can reason about the shape of a line. Deleted text (tracked
 * changes) and field instructions are skipped.
 */
export function textOf(node: Node | null): string {
  if (!node) return '';
  let out = '';
  const walk = (n: Node): void => {
    for (let c = n.firstChild; c; c = c.nextSibling) {
      if (c.nodeType === 3) continue;
      if (c.nodeType !== 1) continue;
      const el = c as Element;
      if (el.namespaceURI !== NS.w) {
        walk(el);
        continue;
      }
      switch (el.localName) {
        case 't':
        case 'delText':
          if (el.localName === 't') out += el.textContent ?? '';
          break;
        case 'tab':
          out += '\t';
          break;
        case 'br':
        case 'cr':
          out += '\n';
          break;
        case 'noBreakHyphen':
          out += '-';
          break;
        case 'softHyphen':
          break;
        case 'del':
        case 'instrText':
          break;
        default:
          walk(el);
      }
    }
  };
  walk(node);
  return out;
}

/** Remove every child node from an element. */
export function clearChildren(el: Element): void {
  while (el.firstChild) el.removeChild(el.firstChild);
}

/** Import a node into `doc` (deep). Wrapper for readability at call sites. */
export function importNode<T extends Node>(doc: Document, node: T): T {
  return doc.importNode(node as never, true) as unknown as T;
}
