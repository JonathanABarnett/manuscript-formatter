import {
  CHAPTER_NUMBER_STYLES,
  DEFAULT_FORMAT_OPTIONS,
  EMPTY_BOOK_DETAILS,
  NO_EXTRA_SECTIONS,
  STYLE_ROLES,
  type BlockRole,
  type BookDetails,
  type ExtraSections,
  type FormatOptions,
  type RunningHeads,
  type StyleRole,
} from './types.js';

/**
 * Everything chosen on the review screen, written to a small file so a
 * revised draft can be formatted the same way without retyping. The author
 * chooses when to save it and where; the app keeps nothing on its own.
 *
 * Paragraph-level changes ("this line is a chapter title") are stored by the
 * paragraph's words rather than its position, so they survive a draft where
 * text has been added or removed above them.
 */

export const CHOICES_FORMAT = 'manuscript-formatter-choices';
export const CHOICES_VERSION = 1;

export interface SavedChoices {
  format: typeof CHOICES_FORMAT;
  version: number;
  options: Omit<FormatOptions, 'roleOverrides'>;
  /** Role changes made by hand, keyed by the paragraph's own words. */
  overrides: Array<{ text: string; role: BlockRole; index: number }>;
}

interface BlockLike {
  index: number;
  text: string;
  autoRole: BlockRole;
}

/** The whole set of choices as JSON text, ready to write wherever the author says. */
export function serializeChoices(options: FormatOptions, blocks: BlockLike[]): string {
  const { roleOverrides, ...rest } = options;
  const byIndex = new Map(blocks.map((b) => [b.index, b]));
  const overrides: SavedChoices['overrides'] = [];
  for (const [key, role] of Object.entries(roleOverrides)) {
    const index = Number(key);
    const block = byIndex.get(index);
    if (!block) continue;
    const text = fingerprint(block.text);
    if (!text) continue;
    overrides.push({ text, role, index });
  }
  const saved: SavedChoices = {
    format: CHOICES_FORMAT,
    version: CHOICES_VERSION,
    options: rest,
    overrides,
  };
  return JSON.stringify(saved, null, 2);
}

/**
 * Read a choices file. Anything malformed or unknown falls back to the app's
 * defaults rather than failing, so a file from a newer or older version still
 * yields something sensible; a file that is not a choices file at all throws.
 */
export function parseChoices(json: string): SavedChoices {
  let raw: unknown;
  try {
    raw = JSON.parse(json);
  } catch {
    throw new Error('That file is not a saved-choices file from this app.');
  }
  if (!isRecord(raw) || raw.format !== CHOICES_FORMAT) {
    throw new Error('That file is not a saved-choices file from this app.');
  }
  const options = isRecord(raw.options) ? sanitizeOptions(raw.options) : { ...defaultsWithoutOverrides() };
  const overrides: SavedChoices['overrides'] = [];
  if (Array.isArray(raw.overrides)) {
    for (const item of raw.overrides) {
      if (!isRecord(item)) continue;
      if (typeof item.text !== 'string' || !isRole(item.role)) continue;
      overrides.push({
        text: item.text,
        role: item.role,
        index: typeof item.index === 'number' ? item.index : -1,
      });
    }
  }
  return {
    format: CHOICES_FORMAT,
    version: typeof raw.version === 'number' ? raw.version : CHOICES_VERSION,
    options,
    overrides,
  };
}

export interface AppliedChoices {
  options: FormatOptions;
  /** How many saved paragraph changes found their paragraph in this draft. */
  matched: number;
  /** How many did not, because the words changed or the line is gone. */
  unmatched: number;
}

/**
 * Turn saved choices back into options for the manuscript at hand, rebinding
 * each paragraph change to the paragraph that carries the same words. Where
 * the same words occur more than once, the nearest to the saved position wins.
 */
export function applyChoices(saved: SavedChoices, blocks: BlockLike[]): AppliedChoices {
  const byText = new Map<string, BlockLike[]>();
  for (const block of blocks) {
    const key = fingerprint(block.text);
    if (!key) continue;
    const list = byText.get(key);
    if (list) list.push(block);
    else byText.set(key, [block]);
  }
  const taken = new Set<number>();
  const roleOverrides: Record<number, BlockRole> = {};
  let matched = 0;
  let unmatched = 0;
  for (const item of saved.overrides) {
    const candidates = (byText.get(fingerprint(item.text)) ?? []).filter((b) => !taken.has(b.index));
    if (candidates.length === 0) {
      unmatched++;
      continue;
    }
    const best = candidates.reduce((a, b) =>
      Math.abs(b.index - item.index) < Math.abs(a.index - item.index) ? b : a,
    );
    taken.add(best.index);
    matched++;
    // A change back to what the classifier said anyway is not a change.
    if (item.role !== best.autoRole) roleOverrides[best.index] = item.role;
  }
  return { options: { ...saved.options, roleOverrides }, matched, unmatched };
}

/** The words of a paragraph, spacing-insensitive, for matching across drafts. */
export function fingerprint(text: string): string {
  return text.replace(/\s+/g, ' ').trim();
}

// --- validation ---------------------------------------------------------------

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value);
}

const ROLES = new Set<string>([
  'frontMatterTitle',
  'frontMatter',
  'copyright',
  'partTitle',
  'chapterTitle',
  'chapterSubtitle',
  'subheading',
  'bodyFirst',
  'body',
  'blockQuote',
  'listItem',
  'sceneBreak',
  'table',
  'pageBreak',
  'empty',
]);
const isRole = (value: unknown): value is BlockRole => typeof value === 'string' && ROLES.has(value);

function defaultsWithoutOverrides(): Omit<FormatOptions, 'roleOverrides'> {
  const { roleOverrides, ...rest } = DEFAULT_FORMAT_OPTIONS;
  void roleOverrides;
  return {
    ...rest,
    roleStyles: {},
    bookDetails: { ...EMPTY_BOOK_DETAILS },
    extraSections: { ...NO_EXTRA_SECTIONS },
    runningHeads: { ...DEFAULT_FORMAT_OPTIONS.runningHeads },
  };
}

/** Keep only known keys with values of the right shape; default the rest. */
function sanitizeOptions(raw: Record<string, unknown>): Omit<FormatOptions, 'roleOverrides'> {
  const out = defaultsWithoutOverrides();

  const bool = (key: keyof typeof out): void => {
    if (typeof raw[key] === 'boolean') (out as Record<string, unknown>)[key] = raw[key];
  };
  for (const key of [
    'chapterOpenerNoHeader',
    'firstParagraphNoIndent',
    'leadInSmallCaps',
    'removeEmptyParagraphs',
    'removeManualIndents',
    'collapseMultipleSpaces',
    'smartTypography',
    'keepEmphasis',
    'underlineToItalic',
    'includeFrontMatter',
    'hyphenate',
    'renumberChapters',
    'replaceFrontMatter',
  ] as const) {
    bool(key);
  }

  if (raw.chapterStart === 'newPage' || raw.chapterStart === 'oddPage' || raw.chapterStart === 'continuous') {
    out.chapterStart = raw.chapterStart;
  }
  if (typeof raw.chapterNumberStyle === 'string' && (CHAPTER_NUMBER_STYLES as string[]).includes(raw.chapterNumberStyle)) {
    out.chapterNumberStyle = raw.chapterNumberStyle as FormatOptions['chapterNumberStyle'];
  }
  if (raw.sceneBreakText === null || typeof raw.sceneBreakText === 'string') {
    out.sceneBreakText = raw.sceneBreakText;
  }
  for (const key of ['chapterSpaceBefore', 'chapterSpaceAfter'] as const) {
    const value = raw[key];
    if (value === null || (typeof value === 'number' && Number.isInteger(value) && value >= 0)) {
      out[key] = value;
    }
  }

  if (isRecord(raw.roleStyles)) {
    const roleStyles: Partial<Record<StyleRole, string | null>> = {};
    for (const role of STYLE_ROLES) {
      const value = raw.roleStyles[role];
      if (value === null || typeof value === 'string') roleStyles[role] = value;
    }
    out.roleStyles = roleStyles;
  }

  if (isRecord(raw.bookDetails)) {
    const details: BookDetails = { ...EMPTY_BOOK_DETAILS };
    for (const key of Object.keys(details) as Array<keyof BookDetails>) {
      const value = raw.bookDetails[key];
      if (typeof value === 'string') details[key] = value;
    }
    out.bookDetails = details;
  }

  if (isRecord(raw.extraSections)) {
    const sections: ExtraSections = { ...NO_EXTRA_SECTIONS };
    for (const key of Object.keys(sections) as Array<keyof ExtraSections>) {
      const value = raw.extraSections[key];
      if (typeof value === 'boolean') sections[key] = value;
    }
    out.extraSections = sections;
  }

  if (isRecord(raw.runningHeads)) {
    const heads: RunningHeads = { ...DEFAULT_FORMAT_OPTIONS.runningHeads };
    const mode = raw.runningHeads.mode;
    if (mode === 'auto' || mode === 'custom' || mode === 'leave') heads.mode = mode;
    if (typeof raw.runningHeads.verso === 'string') heads.verso = raw.runningHeads.verso;
    if (typeof raw.runningHeads.recto === 'string') heads.recto = raw.runningHeads.recto;
    out.runningHeads = heads;
  }

  return out;
}
