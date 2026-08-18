import { describe, expect, it } from 'vitest';
import {
  CHOICES_FORMAT,
  applyChoices,
  parseChoices,
  serializeChoices,
} from '../src/core/choices.js';
import { DEFAULT_FORMAT_OPTIONS, EMPTY_BOOK_DETAILS, type BlockRole, type FormatOptions } from '../src/core/types.js';

const block = (index: number, text: string, autoRole: BlockRole = 'body') => ({ index, text, autoRole });

const OPTIONS: FormatOptions = {
  ...DEFAULT_FORMAT_OPTIONS,
  chapterStart: 'oddPage',
  chapterNumberStyle: 'chapterWords',
  renumberChapters: true,
  smartTypography: true,
  hyphenate: false,
  sceneBreakText: '❦',
  chapterSpaceBefore: 5,
  roleStyles: { chapterTitle: 'Heading1', sceneBreak: null },
  bookDetails: { ...EMPTY_BOOK_DETAILS, title: 'The Sample Book', author: 'A. N. Author', isbn: '978-0-306-40615-7' },
  extraSections: { ...DEFAULT_FORMAT_OPTIONS.extraSections, titlePage: true, contents: true },
  runningHeads: { mode: 'custom', verso: 'A. N. Author', recto: 'The Sample Book' },
  roleOverrides: { 3: 'chapterTitle', 7: 'sceneBreak', 9: 'body' },
};

const BLOCKS = [
  block(0, 'The Sample Book'),
  block(1, 'by A. N. Author'),
  block(2, ''),
  block(3, 'The Long Night', 'subheading'),
  block(4, 'It was a dark and stormy night.'),
  block(5, ''),
  block(6, 'Morning came.'),
  block(7, '~', 'body'),
  block(8, 'The next day was brighter.'),
  block(9, 'The End', 'chapterTitle'),
];

describe('saving and loading the review-screen choices', () => {
  it('round-trips every option, and stores paragraph changes by their words', () => {
    const json = serializeChoices(OPTIONS, BLOCKS);
    const raw = JSON.parse(json);
    expect(raw.format).toBe(CHOICES_FORMAT);
    expect(raw.options.roleOverrides).toBeUndefined();
    expect(raw.overrides).toEqual([
      { text: 'The Long Night', role: 'chapterTitle', index: 3 },
      { text: '~', role: 'sceneBreak', index: 7 },
      { text: 'The End', role: 'body', index: 9 },
    ]);

    const applied = applyChoices(parseChoices(json), BLOCKS);
    expect(applied.matched).toBe(3);
    expect(applied.unmatched).toBe(0);
    expect(applied.options).toEqual(OPTIONS);
  });

  it('follows a paragraph change to its new position in a revised draft', () => {
    const json = serializeChoices(OPTIONS, BLOCKS);
    // A new opening paragraph pushes everything down; one line is reworded.
    const revised = [
      block(0, 'The Sample Book'),
      block(1, 'by A. N. Author'),
      block(2, 'A dedication, added later.'),
      block(3, ''),
      block(4, 'The Long Night', 'subheading'),
      block(5, 'It was a dark and stormy night, and then some.'),
      block(6, 'Morning came.'),
      block(7, '~', 'body'),
      block(8, 'The next day was brighter still.'),
      block(9, 'Fin', 'chapterTitle'),
    ];
    const applied = applyChoices(parseChoices(json), revised);
    expect(applied.matched).toBe(2);
    expect(applied.unmatched).toBe(1);
    expect(applied.options.roleOverrides).toEqual({ 4: 'chapterTitle', 7: 'sceneBreak' });
  });

  it('prefers the nearest of several identical lines, and drops changes back to the automatic role', () => {
    const options: FormatOptions = { ...DEFAULT_FORMAT_OPTIONS, roleOverrides: { 5: 'sceneBreak', 2: 'body' } };
    const blocks = [block(0, 'a'), block(1, '* * *'), block(2, 'b', 'body'), block(3, 'c'), block(4, 'd'), block(5, '* * *')];
    const json = serializeChoices(options, blocks);
    // Same lines, shifted by one; the override on a `body` line that was
    // already body is not saved as a change at all.
    const shifted = [block(0, 'x'), ...blocks.map((b) => block(b.index + 1, b.text, b.autoRole))];
    const applied = applyChoices(parseChoices(json), shifted);
    expect(applied.options.roleOverrides).toEqual({ 6: 'sceneBreak' });
  });

  it('is forgiving of a file from another version, and firm about a file that is something else', () => {
    const odd = JSON.stringify({
      format: CHOICES_FORMAT,
      version: 99,
      options: {
        chapterStart: 'sideways',
        hyphenate: 'yes',
        smartTypography: true,
        futureOption: 42,
        bookDetails: { title: 'Kept', dedication: 12 },
        extraSections: { titlePage: true, futureSection: true },
        runningHeads: { mode: 'custom', recto: 'Recto' },
        chapterSpaceBefore: -3,
      },
      overrides: [{ text: 'x', role: 'notARole' }, { text: 'y', role: 'sceneBreak' }, 'junk'],
    });
    const saved = parseChoices(odd);
    expect(saved.options.chapterStart).toBe(DEFAULT_FORMAT_OPTIONS.chapterStart);
    expect(saved.options.hyphenate).toBe(DEFAULT_FORMAT_OPTIONS.hyphenate);
    expect(saved.options.smartTypography).toBe(true);
    expect((saved.options as Record<string, unknown>).futureOption).toBeUndefined();
    expect(saved.options.bookDetails.title).toBe('Kept');
    expect(saved.options.bookDetails.dedication).toBe('');
    expect(saved.options.extraSections.titlePage).toBe(true);
    expect(saved.options.runningHeads).toEqual({ mode: 'custom', verso: '', recto: 'Recto' });
    expect(saved.options.chapterSpaceBefore).toBeNull();
    expect(saved.overrides).toEqual([{ text: 'y', role: 'sceneBreak', index: -1 }]);

    expect(() => parseChoices('not json')).toThrow(/not a saved-choices file/);
    expect(() => parseChoices('{"format":"something-else"}')).toThrow(/not a saved-choices file/);
  });
});
