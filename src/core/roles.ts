import type { BlockRole, FormatOptions, StyleRole } from './types.js';
import { formatChapterTitle, parseChapterTitle } from './analyze/chapterNumbers.js';

export type RoleStyleMap = Record<StyleRole, string | null>;

/**
 * Which reference style a role resolves to, walking the fallbacks used when the
 * template has no dedicated style for it.
 *
 * Shared by the composer and the review preview so that what the reviewer sees
 * on screen is what actually gets written into the document.
 */
export function styleForRole(role: BlockRole, roles: RoleStyleMap): string | null {
  switch (role) {
    case 'chapterTitle':
      return roles.chapterTitle;
    case 'partTitle':
      return roles.partTitle ?? roles.chapterTitle;
    case 'chapterSubtitle':
      // A subtitle belongs to the chapter's display head, so without a
      // dedicated style it joins the title rather than dropping to a subhead
      // meant for use mid-chapter.
      return roles.chapterSubtitle ?? roles.chapterTitle ?? roles.subheading;
    case 'subheading':
      return roles.subheading ?? roles.chapterSubtitle;
    case 'frontMatterTitle':
      return roles.frontMatterTitle ?? roles.chapterTitle;
    case 'frontMatter':
      return roles.frontMatter ?? roles.bodyFirst ?? roles.body;
    case 'copyright':
      return roles.copyright ?? roles.frontMatter ?? roles.body;
    case 'bodyFirst':
      return roles.bodyFirst ?? roles.body;
    case 'blockQuote':
      return roles.blockQuote ?? roles.body;
    case 'listItem':
      return roles.listItem ?? roles.body;
    case 'sceneBreak':
      return roles.sceneBreak ?? roles.body;
    case 'body':
    default:
      return roles.body;
  }
}

/**
 * The text each chapter title will carry once the chapter-number options are
 * applied, keyed by block index. Titles that carry no number ("Prologue") are
 * left out and are not counted when renumbering. Shared by the composer and
 * the preview so the sample pages show the very titles that will be written.
 */
export function chapterTitleTexts<T extends { index: number; text: string }>(
  blocks: T[],
  roleOf: (block: T) => BlockRole,
  options: Pick<FormatOptions, 'chapterNumberStyle' | 'renumberChapters'>,
): Map<number, string> {
  const out = new Map<number, string>();
  if (options.chapterNumberStyle === 'keep' && !options.renumberChapters) return out;
  let ordinal = 0;
  for (const block of blocks) {
    if (roleOf(block) !== 'chapterTitle') continue;
    const parsed = parseChapterTitle(block.text);
    if (!parsed) continue;
    ordinal++;
    const number = options.renumberChapters ? ordinal : parsed.number;
    const text = formatChapterTitle(parsed, options.chapterNumberStyle, number);
    if (text !== block.text.replace(/\s+/g, ' ').trim()) out.set(block.index, text);
  }
  return out;
}

/**
 * How many characters of a paragraph make its lead-in: up to four words, cut
 * short at the first punctuation mark so a sentence's shape is respected, and
 * carrying a closing quotation mark along with it. Shared by the composer and the preview.
 */
export function leadInLength(text: string): number {
  const match = /^\s*(?:\S+\s+){0,3}\S+/.exec(text);
  if (!match) return 0;
  let lead = match[0];
  const punct = /[.!?;:,—–]/.exec(lead.trimStart());
  if (punct && punct.index > 0) {
    const offset = lead.length - lead.trimStart().length;
    let end = offset + punct.index + 1;
    while (end < lead.length && /[”’"')]/.test(lead[end])) end++;
    lead = lead.slice(0, end);
  }
  return lead.length;
}
