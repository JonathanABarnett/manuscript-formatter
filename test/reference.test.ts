import { mkdtemp, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { afterAll, beforeAll, describe, expect, it } from 'vitest';
import { analyzeReference } from '../src/core/analyze/reference.js';
import { BOOK_TEMPLATE, DIGEST_SECTPR, openDocx, writeDocx } from './helpers/makeDocx.js';

const loadReference = async (path: string) => analyzeReference(await openDocx(path));

let dir = '';

beforeAll(async () => {
  dir = await mkdtemp(join(tmpdir(), 'mf-ref-'));
});
afterAll(async () => {
  await rm(dir, { recursive: true, force: true });
});

describe('reference analysis', () => {
  it('reads page geometry from the body section', async () => {
    const path = await writeDocx(join(dir, 'template.docx'), BOOK_TEMPLATE);
    const { profile } = await loadReference(path);

    expect(profile.pageSetup.widthTwips).toBe(7920);
    expect(profile.pageSetup.heightTwips).toBe(12240);
    expect(profile.pageSetup.margins.gutter).toBe(360);
    expect(profile.pageSetup.margins.top).toBe(1080);
    expect(profile.pageSetup.differentFirstPage).toBe(true);
    expect(profile.pageSizeLabel).toContain('5.5" x 8.5"');
    expect(profile.pageSizeLabel).toContain('Digest');
  });

  it('maps roles onto the template styles', async () => {
    const path = await writeDocx(join(dir, 'roles.docx'), BOOK_TEMPLATE);
    const { profile } = await loadReference(path);

    expect(profile.roleStyles.body).toBe('BodyText');
    expect(profile.roleStyles.chapterTitle).toBe('ChapterTitle');
    expect(profile.roleStyles.bodyFirst).toBe('FirstParagraph');
    expect(profile.roleStyles.sceneBreak).toBe('SceneBreak');
    expect(profile.roleStyles.blockQuote).toBe('BlockQuote');
    expect(profile.roleStyles.frontMatterTitle).toBe('BookTitle');
    expect(profile.roleStyles.subheading).toBe('Subhead');
    expect(profile.roleEvidence.body).toMatch(/regular paragraphs/);
  });

  it('detects chapter pagination and the no-indent opening paragraph', async () => {
    const path = await writeDocx(join(dir, 'layout.docx'), BOOK_TEMPLATE);
    const { profile } = await loadReference(path);

    expect(profile.chapterStartsOnNewPage).toBe(true);
    expect(profile.chapterStartsOnOddPage).toBe(false);
    expect(profile.usesFirstParagraphNoIndent).toBe(true);
    expect(profile.bodyFirstLineIndentTwips).toBe(288);
    expect(profile.bodyFontName).toBe('Garamond');
    expect(profile.bodyFontSizePt).toBe(11);
  });

  it('shows words inherited from template headers and footers', async () => {
    const path = await writeDocx(join(dir, 'running-heads.docx'), {
      ...BOOK_TEMPLATE,
      headerText: 'OLD BOOK TITLE',
      footerText: 'Old Author Name',
    });
    const { profile } = await loadReference(path);

    expect(profile.hasHeaders).toBe(true);
    expect(profile.hasFooters).toBe(true);
    expect(profile.headerFooterText).toEqual(['OLD BOOK TITLE', 'Old Author Name']);
    expect(profile.warnings.some((warning) => /headers or footers/.test(warning))).toBe(true);
  });

  it('detects recto chapter openings from an odd-page section break', async () => {
    const path = await writeDocx(join(dir, 'recto.docx'), {
      ...BOOK_TEMPLATE,
      paragraphs: [
        ...BOOK_TEMPLATE.paragraphs.slice(0, 1),
        { text: '', sectPr: `<w:type w:val="oddPage"/>${DIGEST_SECTPR}` },
        ...BOOK_TEMPLATE.paragraphs.slice(1),
      ],
    });
    const { profile } = await loadReference(path);

    expect(profile.chapterStartsOnOddPage).toBe(true);
    expect(profile.sectionCount).toBe(2);
    expect(profile.warnings.some((w) => /2 different page-layout sections/.test(w))).toBe(true);
  });

  it('uses the next-style chain to break a tie between prose styles', async () => {
    // "Opening" would otherwise tie with "Running Text"; its w:next resolves it.
    const path = await writeDocx(join(dir, 'nextchain.docx'), {
      sectPr: DIGEST_SECTPR,
      styles: [
        { id: 'Normal', name: 'Normal', isDefault: true },
        { id: 'Running', name: 'Running Text', basedOn: 'Normal', firstLineIndent: 288 },
        { id: 'Opener', name: 'Opener', basedOn: 'Normal', next: 'Running', firstLineIndent: 0 },
      ],
      paragraphs: [
        { text: 'x '.repeat(40), style: 'Opener' },
        { text: 'y '.repeat(40), style: 'Running' },
      ],
    });
    const { profile } = await loadReference(path);

    expect(profile.roleStyles.body).toBe('Running');
  });

  it('falls back to style names when the template has no prose', async () => {
    const path = await writeDocx(join(dir, 'empty.docx'), {
      ...BOOK_TEMPLATE,
      paragraphs: [{ text: '', style: 'BodyText' }],
    });
    const { profile } = await loadReference(path);

    expect(profile.roleStyles.body).toBe('BodyText');
    expect(profile.roleStyles.chapterTitle).toBe('ChapterTitle');
    expect(profile.warnings.some((w) => /very little sample text/.test(w))).toBe(true);
  });

  it('rejects files that are not Word documents', async () => {
    const notDocx = join(dir, 'notes.docx');
    await writeDocx(notDocx, BOOK_TEMPLATE);
    const { writeFile } = await import('node:fs/promises');
    await writeFile(notDocx, 'this is plain text, not a zip');

    await expect(loadReference(notDocx)).rejects.toThrow(/not a valid \.docx/);
  });
});
