import type { BookDetails, ManuscriptBlock } from '../types.js';

/**
 * Reads the book's details out of the manuscript's own opening pages.
 *
 * Most authors have already typed their title, name, copyright year and ISBN
 * somewhere in the front matter. Asking them to type it all again is work the
 * app can do for them, so this pre-fills the form and they correct what it got
 * wrong rather than starting from nothing.
 *
 * Only confident readings are returned. A blank field is better than a wrong
 * one, because a wrong one gets printed.
 */

/** Lines that are labels or boilerplate, never a person's name. */
const NOT_A_NAME =
  /^(a novel|a memoir|a story|stories|copyright|all rights|first edition|isbn|printed|published|www\.|http)/i;

const YEAR = /(?:©|\(c\)|copyright)\s*(?:©\s*)?(\d{4})|\b(\d{4})\s*(?:by\b)/i;
const BY_LINE = /^(?:by|written by)\s+(.{2,80})$/i;
const COPYRIGHT_NAME = /(?:©|\(c\)|copyright)\s*(?:©\s*)?\d{4}\s*,?\s*(?:by\s+)?(.{2,80}?)\s*$/i;
const ISBN = /\bISBN(?:-1[03])?:?\s*((?:97[89][-\s]?)?[\d][\d-\s]{7,20}[\dXx])/;

function clean(text: string): string {
  return text.replace(/\s+/g, ' ').trim().replace(/[.,;]+$/, '');
}

/** An ISBN with the punctuation an author actually typed, digits verified. */
function readIsbn(text: string): string | null {
  const match = ISBN.exec(text);
  if (!match) return null;
  const value = match[1].trim().replace(/\s+/g, '');
  const digits = value.replace(/[^0-9Xx]/g, '');
  return digits.length === 10 || digits.length === 13 ? value : null;
}

export function detectBookDetails(blocks: ManuscriptBlock[]): Partial<BookDetails> {
  const found: Partial<BookDetails> = {};
  const say = (key: keyof BookDetails, value: string | null | undefined): void => {
    const trimmed = value ? clean(value) : '';
    if (trimmed && !found[key]) found[key] = trimmed;
  };

  const titled = blocks.filter((b) => b.role === 'frontMatterTitle' && !b.isEmpty);
  const opening = blocks.filter(
    (b) => (b.role === 'frontMatter' || b.role === 'frontMatterTitle') && !b.isEmpty,
  );
  const copyright = blocks.filter((b) => b.role === 'copyright' && !b.isEmpty);

  // --- title: the line the classifier already picked as the title page ----
  say('title', titled[0]?.text);

  // --- author: a "by" line first, then the copyright notice ---------------
  for (const block of opening) {
    const byLine = BY_LINE.exec(block.text.trim());
    if (byLine) {
      say('author', byLine[1]);
      break;
    }
  }
  for (const block of copyright) {
    if (!found.author) say('author', COPYRIGHT_NAME.exec(block.text.trim())?.[1]);
    const year = YEAR.exec(block.text);
    if (year) say('copyrightYear', year[1] ?? year[2]);
    const isbn = readIsbn(block.text);
    if (isbn) say('isbn', isbn);
  }

  // A title page often sets the author on the line under the title with no
  // "by", so take the next short line that is not boilerplate.
  if (!found.author && titled.length > 0) {
    const after = opening.filter((b) => b.index > titled[0].index);
    const candidate = after.find(
      (b) => b.wordCount > 0 && b.wordCount <= 8 && !NOT_A_NAME.test(b.text.trim()),
    );
    say('author', candidate?.text);
  }

  // --- subtitle: a short line between the title and the author ------------
  if (titled.length > 0 && found.author) {
    const between = opening.find(
      (b) =>
        b.index > titled[0].index &&
        clean(b.text) !== found.author &&
        b.wordCount > 0 &&
        b.wordCount <= 15 &&
        !NOT_A_NAME.test(b.text.trim()) &&
        !BY_LINE.test(b.text.trim()),
    );
    if (between && clean(between.text) !== found.title) say('subtitle', between.text);
  }

  return found;
}
