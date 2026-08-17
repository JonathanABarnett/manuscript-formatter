import type { StyleInfo } from '../types.js';
import { attr, child, children, childVal, numAttr, toggleOn } from '../ooxml/xml.js';

/** Paragraph/run properties after inheritance through the `basedOn` chain. */
export interface ResolvedStyleProps {
  fontName: string | null;
  fontSizePt: number | null;
  alignment: string | null;
  firstLineIndentTwips: number | null;
  hangingIndentTwips: number | null;
  leftIndentTwips: number | null;
  spaceBeforeTwips: number | null;
  spaceAfterTwips: number | null;
  lineSpacing: number | null;
  lineRule: string | null;
  pageBreakBefore: boolean;
  keepNext: boolean;
  bold: boolean;
  italic: boolean;
  allCaps: boolean;
  smallCaps: boolean;
  outlineLevel: number | null;
}

const EMPTY_PROPS: ResolvedStyleProps = {
  fontName: null,
  fontSizePt: null,
  alignment: null,
  firstLineIndentTwips: null,
  hangingIndentTwips: null,
  leftIndentTwips: null,
  spaceBeforeTwips: null,
  spaceAfterTwips: null,
  lineSpacing: null,
  lineRule: null,
  pageBreakBefore: false,
  keepNext: false,
  bold: false,
  italic: false,
  allCaps: false,
  smallCaps: false,
  outlineLevel: null,
};

interface RawStyle {
  id: string;
  name: string;
  type: string;
  basedOn: string | null;
  next: string | null;
  isDefault: boolean;
  own: ResolvedStyleProps;
  el: Element;
}

/** Parsed `styles.xml`, with inheritance resolution and lookup by name. */
export class StyleSheet {
  private readonly byId = new Map<string, RawStyle>();
  private readonly resolvedCache = new Map<string, ResolvedStyleProps>();
  private readonly docDefaults: ResolvedStyleProps;
  /** Style id marked `w:default="1"` for paragraphs, usually `Normal`. */
  readonly defaultParagraphStyleId: string | null;

  constructor(stylesDoc: Document | null) {
    const root = stylesDoc?.documentElement ?? null;
    this.docDefaults = root ? readDocDefaults(root) : { ...EMPTY_PROPS };

    let defaultParagraph: string | null = null;
    for (const el of children(root, 'style')) {
      const id = attr(el, 'styleId');
      if (!id) continue;
      const type = attr(el, 'type') ?? 'paragraph';
      const isDefault = isTrue(attr(el, 'default'));
      const raw: RawStyle = {
        id,
        name: childVal(el, 'name') ?? id,
        type,
        basedOn: childVal(el, 'basedOn'),
        next: childVal(el, 'next'),
        isDefault,
        own: readStyleProps(el),
        el,
      };
      this.byId.set(id, raw);
      if (isDefault && type === 'paragraph' && !defaultParagraph) defaultParagraph = id;
    }
    this.defaultParagraphStyleId = defaultParagraph ?? (this.byId.has('Normal') ? 'Normal' : null);
  }

  has(id: string): boolean {
    return this.byId.has(id);
  }

  ids(): string[] {
    return [...this.byId.keys()];
  }

  paragraphStyleIds(): string[] {
    return [...this.byId.values()].filter((s) => s.type === 'paragraph').map((s) => s.id);
  }

  nameOf(id: string): string | null {
    return this.byId.get(id)?.name ?? null;
  }

  typeOf(id: string): string | null {
    return this.byId.get(id)?.type ?? null;
  }

  nextStyleOf(id: string): string | null {
    return this.byId.get(id)?.next ?? null;
  }

  /** Look up a style id by display name, case- and space-insensitively. */
  idByName(name: string): string | null {
    const target = normalizeName(name);
    for (const style of this.byId.values()) {
      if (normalizeName(style.name) === target) return style.id;
    }
    return null;
  }

  /** All paragraph style ids whose id or display name matches `pattern`. */
  findByPattern(pattern: RegExp): string[] {
    return [...this.byId.values()]
      .filter((s) => s.type === 'paragraph' && (pattern.test(s.name) || pattern.test(s.id)))
      .map((s) => s.id);
  }

  /** Effective properties for `id`, merged down the `basedOn` chain. */
  resolve(id: string): ResolvedStyleProps {
    const cached = this.resolvedCache.get(id);
    if (cached) return cached;
    const chain: RawStyle[] = [];
    const seen = new Set<string>();
    let current: string | null = id;
    while (current && !seen.has(current)) {
      seen.add(current);
      const style: RawStyle | undefined = this.byId.get(current);
      if (!style) break;
      chain.unshift(style);
      current = style.basedOn;
    }
    let props = { ...this.docDefaults };
    for (const style of chain) props = mergeProps(props, style.own);
    this.resolvedCache.set(id, props);
    return props;
  }

  /** Serializable view of every paragraph style, with usage counts applied. */
  toStyleInfos(usage: Map<string, number> = new Map()): StyleInfo[] {
    return [...this.byId.values()]
      .filter((s) => s.type === 'paragraph')
      .map((s) => {
        const props = this.resolve(s.id);
        return {
          id: s.id,
          name: s.name,
          type: s.type,
          basedOn: s.basedOn,
          next: s.next,
          isDefault: s.isDefault,
          usageCount: usage.get(s.id) ?? 0,
          ...props,
        } satisfies StyleInfo;
      })
      .sort((a, b) => b.usageCount - a.usageCount || a.name.localeCompare(b.name));
  }
}

function isTrue(value: string | null): boolean {
  return value === '1' || value === 'true' || value === 'on';
}

export function normalizeName(name: string): string {
  return name.toLowerCase().replace(/[\s_-]+/g, '');
}

function readDocDefaults(root: Element): ResolvedStyleProps {
  const defaults = child(root, 'docDefaults');
  const pPr = child(child(defaults, 'pPrDefault'), 'pPr');
  const rPr = child(child(defaults, 'rPrDefault'), 'rPr');
  return mergeProps({ ...EMPTY_PROPS }, propsFrom(pPr, rPr));
}

function readStyleProps(styleEl: Element): ResolvedStyleProps {
  return propsFrom(child(styleEl, 'pPr'), child(styleEl, 'rPr'));
}

/** Extract the properties we care about from a `w:pPr` / `w:rPr` pair. */
export function propsFrom(pPr: Element | null, rPr: Element | null): ResolvedStyleProps {
  const ind = child(pPr, 'ind');
  const spacing = child(pPr, 'spacing');
  const outline = childVal(pPr, 'outlineLvl');
  const sz = childVal(rPr, 'sz');
  const rFonts = child(rPr, 'rFonts');
  const lineRaw = numAttr(spacing, 'line');
  const lineRule = attr(spacing, 'lineRule');

  return {
    fontName: attr(rFonts, 'ascii') ?? attr(rFonts, 'hAnsi') ?? null,
    fontSizePt: sz !== null && Number.isFinite(Number(sz)) ? Number(sz) / 2 : null,
    alignment: childVal(pPr, 'jc'),
    firstLineIndentTwips: numAttr(ind, 'firstLine') ?? numAttr(ind, 'firstLineChars'),
    hangingIndentTwips: numAttr(ind, 'hanging'),
    leftIndentTwips: numAttr(ind, 'left') ?? numAttr(ind, 'start'),
    spaceBeforeTwips: numAttr(spacing, 'before'),
    spaceAfterTwips: numAttr(spacing, 'after'),
    // `auto`/`atLeast` line values are in twips; `exact` too. Report as-is and
    // let the caller interpret with lineRule.
    lineSpacing: lineRaw,
    lineRule: lineRule ?? (lineRaw !== null ? 'auto' : null),
    pageBreakBefore: toggleOn(pPr, 'pageBreakBefore'),
    keepNext: toggleOn(pPr, 'keepNext'),
    bold: toggleOn(rPr, 'b'),
    italic: toggleOn(rPr, 'i'),
    allCaps: toggleOn(rPr, 'caps'),
    smallCaps: toggleOn(rPr, 'smallCaps'),
    outlineLevel: outline !== null && Number.isFinite(Number(outline)) ? Number(outline) : null,
  };
}

/** Child values win over inherited ones; booleans are OR-ed only when set. */
function mergeProps(base: ResolvedStyleProps, over: ResolvedStyleProps): ResolvedStyleProps {
  const pick = <K extends keyof ResolvedStyleProps>(key: K): ResolvedStyleProps[K] =>
    (over[key] === null ? base[key] : over[key]) as ResolvedStyleProps[K];
  return {
    fontName: pick('fontName'),
    fontSizePt: pick('fontSizePt'),
    alignment: pick('alignment'),
    firstLineIndentTwips: pick('firstLineIndentTwips'),
    hangingIndentTwips: pick('hangingIndentTwips'),
    leftIndentTwips: pick('leftIndentTwips'),
    spaceBeforeTwips: pick('spaceBeforeTwips'),
    spaceAfterTwips: pick('spaceAfterTwips'),
    lineSpacing: pick('lineSpacing'),
    lineRule: pick('lineRule'),
    outlineLevel: pick('outlineLevel'),
    pageBreakBefore: over.pageBreakBefore || base.pageBreakBefore,
    keepNext: over.keepNext || base.keepNext,
    bold: over.bold || base.bold,
    italic: over.italic || base.italic,
    allCaps: over.allCaps || base.allCaps,
    smallCaps: over.smallCaps || base.smallCaps,
  };
}

/** Style id applied to a `w:p`, or null when it uses the document default. */
export function paragraphStyleId(p: Element): string | null {
  return childVal(child(p, 'pPr'), 'pStyle');
}
