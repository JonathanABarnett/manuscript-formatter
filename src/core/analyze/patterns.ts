/** Text-shape heuristics shared by the reference and manuscript analyzers. */

/** Ornaments that mean "scene break" on their own. */
const STRONG_ORNAMENTS = '*#~•·◆❖❦§✻✽⁂✦★☆✧❉‡†';
/** Characters that only mean a break when repeated. */
const WEAK_ORNAMENTS = '-–—_=+.';

const STRONG_RE = new RegExp(`^[${escapeClass(STRONG_ORNAMENTS)}](?:\\s*[${escapeClass(STRONG_ORNAMENTS)}])*$`, 'u');
const WEAK_RE = new RegExp(`^[${escapeClass(WEAK_ORNAMENTS)}](?:\\s*[${escapeClass(WEAK_ORNAMENTS)}])*$`, 'u');

function escapeClass(chars: string): string {
  return chars.replace(/[\\\]^-]/g, '\\$&');
}

/**
 * A line that separates scenes: `***`, `* * *`, `#`, `~`, `❦`, `---`.
 * Weak ornaments need at least three repeats so an em dash used for dialogue
 * or a stray hyphen is not mistaken for a break.
 */
export function isSceneBreakText(text: string): boolean {
  const t = text.trim();
  if (t.length === 0 || t.length > 24) return false;
  const dense = t.replace(/\s+/g, '');
  if (dense.length > 9) return false;
  if (STRONG_RE.test(t)) return true;
  return WEAK_RE.test(t) && dense.length >= 3;
}

export type LabelKind = 'chapter' | 'part' | 'frontMatter' | 'backMatter';

const LABEL_PATTERNS: Array<{ kind: LabelKind; re: RegExp }> = [
  { kind: 'part', re: /^(part|volume|book)\b(?!\s*(one|two)?\s*$)?/i },
  { kind: 'chapter', re: /^(chapter|chap\.?|ch\.?|section|interlude|episode|canto|scene)\b/i },
  {
    // "Copyright" is deliberately absent: a copyright *notice* opens with it
    // and is body text of the copyright page, not a heading over one.
    kind: 'frontMatter',
    re: /^(title page|half title|contents|table of contents|dedication|epigraph|foreword|preface|introduction|prologue|author'?s note|a note (on|from)|praise for)\b/i,
  },
  {
    kind: 'backMatter',
    re: /^(epilogue|afterword|appendix|appendices|glossary|notes?|endnotes|bibliography|works cited|further reading|index|acknowledge?ments?|about the author|also by|discussion questions|reading group guide|colophon)\b/i,
  },
];

/** Which structural label a line starts with, if any. */
export function labelKind(text: string): LabelKind | null {
  const t = text.trim();
  if (!t) return null;
  for (const { kind, re } of LABEL_PATTERNS) {
    if (re.test(t)) return kind;
  }
  return null;
}

const NUMBER_WORDS = new Set([
  'zero', 'one', 'two', 'three', 'four', 'five', 'six', 'seven', 'eight', 'nine', 'ten',
  'eleven', 'twelve', 'thirteen', 'fourteen', 'fifteen', 'sixteen', 'seventeen', 'eighteen',
  'nineteen', 'twenty', 'thirty', 'forty', 'fifty', 'sixty', 'seventy', 'eighty', 'ninety',
  'hundred',
]);

const ROMAN_RE = /^[ivxlcdm]+$/i;

/**
 * A line that is nothing but a chapter number: `7`, `VII`, `Seven`,
 * `Twenty-One`. Returns a confidence so ambiguous cases (a lone `I`) can be
 * flagged for review rather than silently promoted to a chapter title.
 */
export function standaloneNumber(text: string): { match: boolean; confidence: number } {
  const t = text.trim().replace(/[.:—–-]+$/, '').trim();
  if (!t || t.length > 40) return { match: false, confidence: 0 };

  if (/^\d{1,3}$/.test(t)) return { match: true, confidence: 0.85 };

  if (ROMAN_RE.test(t)) {
    // A bare `I` is usually the pronoun in a fragment, not a chapter number.
    const ambiguous = /^i$/i.test(t);
    return { match: true, confidence: ambiguous ? 0.35 : 0.7 };
  }

  // Chapter numbers spelled out reach two parts at most ("Twenty-One", "One
  // Hundred"). Allowing three would swallow titles like "Nineteen Eighty-Four".
  const words = t.toLowerCase().split(/[\s-]+/);
  if (words.length <= 2 && words.every((w) => NUMBER_WORDS.has(w))) {
    return { match: true, confidence: 0.7 };
  }

  return { match: false, confidence: 0 };
}

/**
 * Lines that only appear on a copyright page. Matching any one of these marks
 * the whole page, since the surrounding lines (publisher, printing history)
 * carry no markers of their own.
 */
const COPYRIGHT_MARKERS = [
  /copyright\s*(©|\(c\))?\s*\d{4}/i,
  /©\s*\d{4}/,
  /\ball rights reserved\b/i,
  /\bISBN\b/i,
  /\bfirst (edition|printing|published)\b/i,
  /\bprinted in\b/i,
  /\bno part of this (book|publication)\b/i,
  /\blibrary of congress\b/i,
  /\bcataloguing[- ]in[- ]publication\b/i,
];

export function looksLikeCopyright(text: string): boolean {
  const t = text.trim();
  if (!t || t.length > 600) return false;
  return COPYRIGHT_MARKERS.some((re) => re.test(t));
}

/** Title Case or ALL CAPS — typical of headings, unusual mid-prose. */
export function looksLikeHeadingCase(text: string): boolean {
  const t = text.trim();
  if (!t) return false;
  if (t === t.toUpperCase() && /[A-Z]/.test(t)) return true;
  const words = t.split(/\s+/).filter((w) => /[a-z]/i.test(w));
  if (words.length === 0) return false;
  const minor = new Set([
    'a', 'an', 'and', 'as', 'at', 'but', 'by', 'for', 'in', 'nor', 'of', 'on', 'or',
    'the', 'to', 'up', 'with', 'from', 'over', 'into',
  ]);
  const significant = words.filter((w, i) => i === 0 || !minor.has(w.toLowerCase()));
  if (significant.length === 0) return false;
  const capitalized = significant.filter((w) => /^[^a-z]*[A-Z]/.test(w)).length;
  return capitalized / significant.length >= 0.8;
}

/** Prose usually ends in terminal punctuation; headings usually do not. */
export function endsLikeSentence(text: string): boolean {
  return /[.!?…]["'”’)\]]?\s*$/.test(text.trim());
}

