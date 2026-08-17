import { NS } from '../ooxml/ns.js';
import { attr, child, children, childVal, descendants, textOf, toggleOn } from '../ooxml/xml.js';
import { StyleSheet, propsFrom } from './styles.js';

/** Everything a classifier needs to know about one `w:p`, style-resolved. */
export interface ParagraphFacts {
  el: Element;
  text: string;
  trimmed: string;
  styleId: string | null;
  styleName: string | null;
  outlineLevel: number | null;
  alignment: string | null;
  leftIndentTwips: number | null;
  firstLineIndentTwips: number | null;
  spaceBeforeTwips: number | null;
  /** `w:pageBreakBefore` in the paragraph properties or resolved style. */
  pageBreakBeforeProp: boolean;
  /** A `w:br w:type="page"` ahead of any text in this paragraph. */
  leadingPageBreak: boolean;
  /** A `w:br w:type="page"` after the text of this paragraph. */
  trailingPageBreak: boolean;
  /** The paragraph carries a `w:sectPr`, i.e. a section break ends here. */
  sectionBreakType: string | null;
  hasNumbering: boolean;
  allBold: boolean;
  allItalic: boolean;
  allCaps: boolean;
  centered: boolean;
  hasImage: boolean;
  hasFootnote: boolean;
  hasHyperlink: boolean;
  hasTabIndent: boolean;
  wordCount: number;
  isEmpty: boolean;
}

/** Read and style-resolve a paragraph. */
export function readParagraph(p: Element, styles: StyleSheet): ParagraphFacts {
  const pPr = child(p, 'pPr');
  const styleId = childVal(pPr, 'pStyle');
  const resolved = styleId && styles.has(styleId) ? styles.resolve(styleId) : null;
  const direct = propsFrom(pPr, null);

  const text = textOf(p);
  const trimmed = text.trim();
  const runs = collectRuns(p);

  const alignment = direct.alignment ?? resolved?.alignment ?? null;
  const leftIndent = direct.leftIndentTwips ?? resolved?.leftIndentTwips ?? null;
  const firstLine = direct.firstLineIndentTwips ?? resolved?.firstLineIndentTwips ?? null;

  const { leading, trailing } = pageBreakPositions(p);

  const sectPr = child(pPr, 'sectPr');

  return {
    el: p,
    text,
    trimmed,
    styleId,
    styleName: styleId ? styles.nameOf(styleId) : null,
    outlineLevel: direct.outlineLevel ?? resolved?.outlineLevel ?? null,
    alignment,
    leftIndentTwips: leftIndent,
    firstLineIndentTwips: firstLine,
    spaceBeforeTwips: direct.spaceBeforeTwips ?? resolved?.spaceBeforeTwips ?? null,
    pageBreakBeforeProp: toggleOn(pPr, 'pageBreakBefore') || (resolved?.pageBreakBefore ?? false),
    leadingPageBreak: leading,
    trailingPageBreak: trailing,
    sectionBreakType: sectPr ? (childVal(sectPr, 'type') ?? 'nextPage') : null,
    hasNumbering: child(pPr, 'numPr') !== null,
    allBold: runsAll(runs, styles, resolved?.bold ?? false, 'b'),
    allItalic: runsAll(runs, styles, resolved?.italic ?? false, 'i'),
    allCaps:
      runsAll(runs, styles, resolved?.allCaps ?? false, 'caps') ||
      (trimmed.length > 0 && trimmed === trimmed.toUpperCase() && /[A-Z]/.test(trimmed)),
    centered: alignment === 'center',
    hasImage:
      descendants(p, 'drawing').length > 0 ||
      descendants(p, 'pict').length > 0 ||
      descendants(p, 'object').length > 0,
    hasFootnote:
      descendants(p, 'footnoteReference').length > 0 ||
      descendants(p, 'endnoteReference').length > 0,
    hasHyperlink: descendants(p, 'hyperlink').length > 0,
    hasTabIndent: /^[\t ]+\S/.test(text) || startsWithTab(p),
    wordCount: countWords(trimmed),
    isEmpty: trimmed.length === 0,
  };
}

/** Runs that belong to the paragraph, including those inside hyperlinks. */
function collectRuns(p: Element): Element[] {
  return descendants(p, 'r').filter((r) => {
    // Skip runs nested in deleted content or in footnote/endnote separators.
    for (let node: Node | null = r.parentNode; node && node !== p; node = node.parentNode) {
      const el = node as Element;
      if (el.namespaceURI === NS.w && (el.localName === 'del' || el.localName === 'moveFrom')) {
        return false;
      }
    }
    return true;
  });
}

/** True when every text-bearing run has the toggle on (directly or by style). */
function runsAll(
  runs: Element[],
  styles: StyleSheet,
  styleDefault: boolean,
  toggle: 'b' | 'i' | 'caps',
): boolean {
  const textRuns = runs.filter((r) => textOf(r).trim().length > 0);
  if (textRuns.length === 0) return false;
  return textRuns.every((r) => {
    const rPr = child(r, 'rPr');
    const explicit = child(rPr, toggle);
    if (explicit) return toggleOn(rPr, toggle);
    const charStyle = childVal(rPr, 'rStyle');
    if (charStyle && styles.has(charStyle)) {
      const props = styles.resolve(charStyle);
      if (toggle === 'b' && props.bold) return true;
      if (toggle === 'i' && props.italic) return true;
      if (toggle === 'caps' && props.allCaps) return true;
    }
    return styleDefault;
  });
}

/** Locate page breaks relative to the paragraph's first piece of visible text. */
function pageBreakPositions(p: Element): { leading: boolean; trailing: boolean } {
  let seenText = false;
  let leading = false;
  let trailing = false;
  const walk = (node: Node): void => {
    for (let n = node.firstChild; n; n = n.nextSibling) {
      if (n.nodeType !== 1) continue;
      const el = n as Element;
      if (el.namespaceURI === NS.w) {
        if (el.localName === 'br' && attr(el, 'type') === 'page') {
          if (seenText) trailing = true;
          else leading = true;
          continue;
        }
        if (el.localName === 't' && (el.textContent ?? '').trim().length > 0) {
          seenText = true;
          continue;
        }
      }
      walk(el);
    }
  };
  walk(p);
  return { leading, trailing };
}

function startsWithTab(p: Element): boolean {
  for (const r of children(p, 'r')) {
    for (let n = r.firstChild; n; n = n.nextSibling) {
      if (n.nodeType !== 1) continue;
      const el = n as Element;
      if (el.namespaceURI !== NS.w) continue;
      if (el.localName === 'rPr') continue;
      if (el.localName === 'tab') return true;
      if (el.localName === 't') return (el.textContent ?? '').startsWith('\t');
      return false;
    }
  }
  return false;
}

export function countWords(text: string): number {
  const t = text.trim();
  if (!t) return 0;
  return t.split(/\s+/).length;
}
