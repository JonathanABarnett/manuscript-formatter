import type { ChapterNumberStyle } from '../types.js';

/**
 * Chapter numbers as authors write them — `Chapter 7`, `CHAPTER SEVEN`,
 * `VII`, `7. The Meeting`, `Seven: The Meeting` — read into a number and put
 * back in whatever form the author asks for. Used by the report to check the
 * sequence and by the composer to make every chapter opening match.
 */

export type NumberForm = 'digits' | 'roman' | 'words';

export interface ParsedChapterTitle {
  /** The word before the number, exactly as written ("Chapter", "CHAPTER", "Chap."). */
  label: string | null;
  /** Whitespace between the label and the number. */
  labelGap: string;
  number: number;
  form: NumberForm;
  /** The number exactly as written, so `keep` can put it back untouched. */
  numberText: string;
  /** Punctuation and spacing between the number and the rest, if any. */
  separator: string;
  /** Whatever follows: a chapter title such as "The Meeting". */
  rest: string;
  /** The label and number were written in capitals. */
  caps: boolean;
}

const ONES = [
  '', 'one', 'two', 'three', 'four', 'five', 'six', 'seven', 'eight', 'nine', 'ten',
  'eleven', 'twelve', 'thirteen', 'fourteen', 'fifteen', 'sixteen', 'seventeen',
  'eighteen', 'nineteen',
];
const TENS = ['', '', 'twenty', 'thirty', 'forty', 'fifty', 'sixty', 'seventy', 'eighty', 'ninety'];

const WORD_VALUES = new Map<string, number>();
ONES.forEach((w, i) => w && WORD_VALUES.set(w, i));
TENS.forEach((w, i) => w && WORD_VALUES.set(w, i * 10));

/** Labels the app can write a spelled-out English number after. */
const ENGLISH_LABEL = /^(chapter|chap\.?|ch\.?|section|interlude|episode|canto|scene|part)$/i;
/** Any chapter label the classifier knows, English or otherwise. */
const LABEL_RE =
  /^(chapter|chap\.?|ch\.?|section|interlude|episode|canto|scene|cap[íi]tulo|cap\.?|chapitre|kapitel|kap\.?|capitolo|hoofdstuk)/i;

const ROMAN_STRICT = /^M{0,3}(CM|CD|D?C{0,3})(XC|XL|L?X{0,3})(IX|IV|V?I{0,3})$/;

/** Read a chapter title. Null when it carries no number to speak of. */
export function parseChapterTitle(text: string): ParsedChapterTitle | null {
  const t = text.replace(/\s+/g, ' ').trim();
  if (!t) return null;

  const labelMatch = LABEL_RE.exec(t);
  let label: string | null = null;
  let labelGap = '';
  let after = t;
  if (labelMatch) {
    const following = t.slice(labelMatch[0].length);
    const gap = /^\s+/.exec(following)?.[0] ?? '';
    // "Chapters" or "Chaptered" is not a label; the label must end the word.
    if (gap.length > 0 || following.length === 0) {
      label = labelMatch[0];
      labelGap = gap;
      after = following.slice(gap.length);
    }
  }
  if (!after) return null;

  const numberMatch = readNumber(after);
  if (!numberMatch) return null;
  const { number, form, numberText } = numberMatch;
  const tail = after.slice(numberText.length);

  // Without a label, only a punctuation separator (or nothing) may follow the
  // number: "One Day More" and "I Am Legend" are titles, not numbered chapters.
  const sepMatch = /^(\s*[:.—–\-]+\s*|\s+)/.exec(tail);
  const separator = sepMatch?.[0] ?? '';
  const rest = tail.slice(separator.length);
  if (rest.length > 0 && !separator) return null;
  if (rest.length > 0 && !label && !/[:.—–\-]/.test(separator)) return null;
  // A digit run glued to letters ("3rd", "1st") is not a number on its own.
  if (form === 'digits' && /^[a-z]/i.test(tail)) return null;

  // Roman numerals are always capitals, so only the label and a spelled-out
  // number say whether the author writes chapter openings in caps.
  const capsSource = `${label ?? ''}${form === 'words' ? numberText : ''}`.replace(/[^A-Za-z]/g, '');
  const caps = capsSource.length > 0 && capsSource === capsSource.toUpperCase();

  return { label, labelGap, number, form, numberText, separator, rest, caps };
}

function readNumber(
  text: string,
): { number: number; form: NumberForm; numberText: string } | null {
  const digits = /^\d{1,3}/.exec(text);
  if (digits) return { number: Number(digits[0]), form: 'digits', numberText: digits[0] };

  // Below 400 only: no book has that many chapters, and it keeps capitalised
  // words that happen to spell a numeral ("MIX", "DIM") from being read as one.
  const romanToken = /^[IVXLCDM]+(?![A-Za-z])/.exec(text);
  if (romanToken && ROMAN_STRICT.test(romanToken[0]) && romanValue(romanToken[0]) < 400) {
    return { number: romanValue(romanToken[0]), form: 'roman', numberText: romanToken[0] };
  }

  // Up to three words: "Twenty-One", "One Hundred", "Hundred and One" is more
  // than any book needs and "Nineteen Eighty-Four" is a title, not a number.
  const wordsToken = /^([A-Za-z]+(?:[\s-]+(?:and[\s-]+)?[A-Za-z]+){0,2})/.exec(text);
  if (wordsToken) {
    // Trim back to the longest prefix that reads as a number.
    const parts = wordsToken[1].split(/([\s-]+)/);
    for (let n = parts.length; n >= 1; n -= 2) {
      const candidate = parts.slice(0, n).join('');
      const value = wordsValue(candidate);
      if (value !== null) return { number: value, form: 'words', numberText: candidate };
    }
  }
  return null;
}

function romanValue(roman: string): number {
  const values: Record<string, number> = { I: 1, V: 5, X: 10, L: 50, C: 100, D: 500, M: 1000 };
  let total = 0;
  for (let i = 0; i < roman.length; i++) {
    const v = values[roman[i]];
    const next = values[roman[i + 1]] ?? 0;
    total += v < next ? -v : v;
  }
  return total;
}

/** "Twenty-One" -> 21, "One Hundred and Two" -> 102, "Seven" -> 7. Null otherwise. */
export function wordsValue(text: string): number | null {
  const words = text.toLowerCase().split(/[\s-]+/).filter((w) => w && w !== 'and');
  if (words.length === 0) return null;
  let total = 0;
  let current = 0;
  for (const w of words) {
    if (w === 'hundred') {
      if (current === 0) current = 1;
      current *= 100;
      continue;
    }
    const v = WORD_VALUES.get(w);
    if (v === undefined) return null;
    // "Twenty One" adds; "One Twenty" does not read as a number.
    if (current % 100 !== 0 && current % 10 !== 0 && v < 10) return null;
    if (current % 100 !== 0 && v >= 10) return null;
    current += v;
  }
  total += current;
  return total > 0 ? total : null;
}

/** 21 -> "Twenty-One", 100 -> "One Hundred", 7 -> "Seven". */
export function numberWords(n: number): string {
  if (n <= 0 || n >= 1000) return String(n);
  const parts: string[] = [];
  const hundreds = Math.floor(n / 100);
  const remainder = n % 100;
  if (hundreds > 0) parts.push(`${cap(ONES[hundreds])} Hundred`);
  if (remainder > 0) {
    if (remainder < 20) parts.push(cap(ONES[remainder]));
    else {
      const tens = cap(TENS[Math.floor(remainder / 10)]);
      const ones = remainder % 10;
      parts.push(ones ? `${tens}-${cap(ONES[ones])}` : tens);
    }
  }
  return parts.join(hundreds > 0 && remainder > 0 ? ' and ' : ' ');
}

export function romanNumeral(n: number): string {
  if (n <= 0 || n >= 4000) return String(n);
  const table: Array<[number, string]> = [
    [1000, 'M'], [900, 'CM'], [500, 'D'], [400, 'CD'], [100, 'C'], [90, 'XC'],
    [50, 'L'], [40, 'XL'], [10, 'X'], [9, 'IX'], [5, 'V'], [4, 'IV'], [1, 'I'],
  ];
  let out = '';
  let rest = n;
  for (const [value, glyph] of table) {
    while (rest >= value) {
      out += glyph;
      rest -= value;
    }
  }
  return out;
}

function cap(word: string): string {
  return word ? word[0].toUpperCase() + word.slice(1) : word;
}

/**
 * Write a chapter title in the requested style, with `number` in place of
 * the one written. `keep` reproduces the author's own form and label; the
 * others impose one. Spelled-out numbers are English, so a foreign label
 * ("Capítulo") keeps digits instead.
 */
export function formatChapterTitle(
  parsed: ParsedChapterTitle,
  style: ChapterNumberStyle,
  number: number,
): string {
  const wantsLabel =
    style === 'chapterWords' || style === 'chapterDigits' || (style === 'keep' && parsed.label !== null);
  const englishLabel = parsed.label === null || ENGLISH_LABEL.test(parsed.label);

  let form: NumberForm;
  if (style === 'keep') form = parsed.form;
  else if (style === 'chapterWords' || style === 'words') form = englishLabel ? 'words' : 'digits';
  else form = 'digits';

  let numberText: string;
  if (form === 'digits') numberText = String(number);
  else if (form === 'roman') numberText = romanNumeral(number);
  else numberText = parsed.caps ? numberWords(number).toUpperCase() : numberWords(number);

  let label = '';
  if (wantsLabel) {
    const word = parsed.label ?? (parsed.caps ? 'CHAPTER' : 'Chapter');
    label = `${word}${parsed.labelGap || ' '}`;
  }
  return `${label}${numberText}${parsed.separator}${parsed.rest}`;
}

/** What the sequence check found across the book's numbered chapters. */
export interface ChapterNumberReport {
  /** Chapter titles that carry a number, in reading order. */
  numbered: Array<{ index: number; number: number; form: NumberForm; labelled: boolean }>;
  /** Numbers repeated. */
  duplicates: number[];
  /** Numbers skipped between one chapter and the next. */
  gaps: number[];
  /** Chapters whose number is lower than the one before. */
  outOfOrder: number[];
  /** Digits in some chapters and words or numerals in others, or a label in some and not others. */
  mixed: boolean;
}

export function checkChapterNumbers(
  titles: Array<{ index: number; text: string }>,
): ChapterNumberReport {
  const numbered: ChapterNumberReport['numbered'] = [];
  for (const { index, text } of titles) {
    const parsed = parseChapterTitle(text);
    if (parsed) {
      numbered.push({ index, number: parsed.number, form: parsed.form, labelled: parsed.label !== null });
    }
  }
  const duplicates: number[] = [];
  const gaps: number[] = [];
  const outOfOrder: number[] = [];
  const seen = new Set<number>();
  let previous: number | null = null;
  for (const item of numbered) {
    if (seen.has(item.number)) duplicates.push(item.number);
    seen.add(item.number);
    if (previous !== null) {
      if (item.number < previous) outOfOrder.push(item.number);
      else if (item.number > previous + 1) {
        for (let missing = previous + 1; missing < item.number; missing++) {
          if (!seen.has(missing)) gaps.push(missing);
        }
      }
    }
    previous = item.number;
  }
  const forms = new Set(numbered.map((n) => n.form));
  const labelled = new Set(numbered.map((n) => n.labelled));
  return {
    numbered,
    duplicates: [...new Set(duplicates)],
    gaps: [...new Set(gaps)],
    outOfOrder,
    mixed: numbered.length > 1 && (forms.size > 1 || labelled.size > 1),
  };
}
