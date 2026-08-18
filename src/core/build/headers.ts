import { NS, RELTYPE } from '../ooxml/ns.js';
import type { DocxPackage, Relationships } from '../ooxml/package.js';
import { descendants, textOf } from '../ooxml/xml.js';
import type { BookDetails, RunningHeads } from '../types.js';

/**
 * Puts the author's own title and name into the running heads.
 *
 * A template ships its headers filled with wording like "BOOK TITLE" and
 * "AUTHOR NAME". Those are copied across with everything else, so a book whose
 * details have been typed in would otherwise still be headed by the template's
 * placeholders on every page.
 *
 * Only text that is recognisably a placeholder is touched. Real wording an
 * author put in their own template is left exactly as it is.
 */

interface Substitution {
  pattern: RegExp;
  value: (details: BookDetails) => string;
}

const SUBSTITUTIONS: Substitution[] = [
  {
    // "BOOK TITLE", "Book Title", "[Title]", "Your Title", "TITLE"
    pattern: /^\[?\s*(your\s+)?(book\s*)?title\s*(goes\s*here)?\s*\]?$/i,
    value: (d) => d.title,
  },
  {
    // "AUTHOR NAME", "Author", "[Your Name]", "By Author Name"
    pattern: /^\[?\s*(by\s+)?(your\s+|the\s+)?(author('s)?\s*)?name\s*(goes\s*here)?\s*\]?$|^\[?\s*author\s*\]?$/i,
    value: (d) => d.author,
  },
  {
    pattern: /^\[?\s*(book\s*)?subtitle\s*(goes\s*here)?\s*\]?$/i,
    value: (d) => d.subtitle,
  },
];

/**
 * Whether this wording is a placeholder that the typed-in details will
 * replace. The preflight report asks so it does not warn about headers that
 * are about to be corrected anyway.
 */
export function willReplaceRunningHead(text: string, details: BookDetails): boolean {
  const trimmed = text.trim();
  const rule = SUBSTITUTIONS.find((s) => s.pattern.test(trimmed));
  return rule !== undefined && rule.value(details).trim().length > 0;
}

/**
 * Replace every visible word in a header part with `text`, keeping the first
 * run so the design's own font and size survive. An empty string clears it.
 */
async function setHeaderText(
  pkg: DocxPackage,
  partPath: string,
  text: string,
): Promise<boolean> {
  const doc = await pkg.readXml(partPath);
  if (!doc?.documentElement) return false;
  const paragraphs = descendants(doc.documentElement, 'p').filter(
    (p) => textOf(p).trim().length > 0 || descendants(p, 't').length > 0,
  );
  if (paragraphs.length === 0) return false;

  let done = false;
  for (const paragraph of paragraphs) {
    const texts = descendants(paragraph, 't');
    if (texts.length === 0) continue;
    // A header holding a page-number field is left alone; replacing its text
    // would strip the field and freeze the number.
    if (descendants(paragraph, 'instrText').length > 0) continue;
    const [first, ...rest] = texts;
    while (first.firstChild) first.removeChild(first.firstChild);
    if (!done && text) first.appendChild(doc.createTextNode(text));
    first.setAttributeNS(NS.xml, 'xml:space', 'preserve');
    for (const extra of rest) while (extra.firstChild) extra.removeChild(extra.firstChild);
    done = true;
  }
  if (!done) return false;
  pkg.writeXml(partPath, doc);
  return true;
}

/** Keep the design's shouting: an all-caps placeholder stays all caps. */
function matchCase(original: string, replacement: string): string {
  const letters = original.replace(/[^A-Za-z]/g, '');
  if (letters.length > 1 && letters === letters.toUpperCase()) return replacement.toUpperCase();
  return replacement;
}

export interface HeaderRewriteResult {
  /** Parts whose wording was replaced, for reporting. */
  changed: string[];
}

/**
 * Rewrite placeholder wording in every header and footer of `pkg`. Returns the
 * parts that changed so the caller can say what happened.
 */
export async function applyDetailsToRunningHeads(
  pkg: DocxPackage,
  rels: Relationships,
  details: BookDetails,
  heads: RunningHeads,
  /** Relationship id -> `w:type` of the header reference that points at it. */
  headerSides: Map<string, string> = new Map(),
): Promise<HeaderRewriteResult> {
  const changed: string[] = [];
  if (heads.mode === 'leave') return { changed };

  // Set outright, rather than only correcting placeholder wording.
  if (heads.mode === 'custom') {
    for (const rel of rels.all()) {
      if (rel.type !== RELTYPE.header) continue;
      const side = headerSides.get(rel.id);
      // `even` is the left-hand page; `default` carries the right-hand one.
      const wanted = side === 'even' ? heads.verso : side === 'default' ? heads.recto : null;
      if (wanted === null) continue;
      const partPath = pkg.resolveTarget(pkg.documentPath, rel.target);
      if (await setHeaderText(pkg, partPath, wanted.trim())) changed.push(partPath);
    }
    return { changed };
  }

  const parts = rels
    .all()
    .filter((r) => r.type === RELTYPE.header || r.type === RELTYPE.footer)
    .map((r) => pkg.resolveTarget(pkg.documentPath, r.target));

  for (const partPath of new Set(parts)) {
    const doc = await pkg.readXml(partPath);
    if (!doc?.documentElement) continue;
    let touched = false;

    for (const paragraph of descendants(doc.documentElement, 'p')) {
      const whole = textOf(paragraph).trim();
      if (!whole) continue;
      const rule = SUBSTITUTIONS.find((s) => s.pattern.test(whole));
      if (!rule) continue;
      const replacement = rule.value(details).trim();
      if (!replacement) continue;

      // Put the whole replacement in the first text run and empty the rest,
      // so the run's own font and size are kept.
      const texts = descendants(paragraph, 't');
      if (texts.length === 0) continue;
      const first = texts[0];
      while (first.firstChild) first.removeChild(first.firstChild);
      first.appendChild(doc.createTextNode(matchCase(whole, replacement)));
      first.setAttributeNS(NS.xml, 'xml:space', 'preserve');
      for (const extra of texts.slice(1)) {
        while (extra.firstChild) extra.removeChild(extra.firstChild);
      }
      touched = true;
    }

    if (touched) {
      pkg.writeXml(partPath, doc);
      changed.push(partPath);
    }
  }

  return { changed };
}
