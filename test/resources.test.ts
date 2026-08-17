import { mkdtemp, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { afterAll, beforeAll, describe, expect, it } from 'vitest';
import { DocxPackage } from '../src/core/ooxml/package.js';
import { NS } from '../src/core/ooxml/ns.js';
import { attr, child, children, descendants, textOf } from '../src/core/ooxml/xml.js';
import { formatManuscript } from '../src/core/platform/node.js';
import { BOOK_TEMPLATE, loadDocx, TEST_PNG, writeDocx } from './helpers/makeDocx.js';

let dir = '';
let counter = 0;

beforeAll(async () => {
  dir = await mkdtemp(join(tmpdir(), 'mf-res-'));
});
afterAll(async () => {
  await rm(dir, { recursive: true, force: true });
});

const PROSE =
  'The morning came in grey and unhurried, the way mornings did there. Hollis ' +
  'pushed the shop door open with his shoulder and waited for the bell.';

async function run(manuscriptSpec: Parameters<typeof writeDocx>[1]): Promise<DocxPackage> {
  const id = counter++;
  const reference = await writeDocx(join(dir, `ref${id}.docx`), BOOK_TEMPLATE);
  const manuscript = await writeDocx(join(dir, `ms${id}.docx`), manuscriptSpec);
  const output = join(dir, `out${id}.docx`);
  await formatManuscript({ referencePath: reference, manuscriptPath: manuscript, outputPath: output });
  return loadDocx(output);
}

describe('carrying resources across from the manuscript', () => {
  it('moves footnotes into a template that has none', async () => {
    const pkg = await run({
      footnotes: [
        { id: 1, text: 'Orwell was writing in 1948.' },
        { id: 2, text: 'The clocks are the giveaway.' },
      ],
      paragraphs: [
        { text: 'Chapter One', alignment: 'center' },
        { text: PROSE, footnoteRef: 1 },
        { text: 'A second paragraph carries the other note.', footnoteRef: 2 },
      ],
    });

    // The footnotes part must exist, be declared, and be related to the document.
    expect(pkg.has('word/footnotes.xml')).toBe(true);
    const types = await pkg.readText('[Content_Types].xml');
    expect(types).toContain('footnotes+xml');
    const rels = await pkg.relsFor(pkg.documentPath);
    expect(rels.all().some((r) => r.type.endsWith('/footnotes'))).toBe(true);

    const notes = await pkg.readXml('word/footnotes.xml');
    const real = children(notes?.documentElement ?? null, 'footnote').filter(
      (el) => Number(attr(el, 'id') ?? '0') > 0,
    );
    expect(real).toHaveLength(2);
    expect(real.map((el) => textOf(el).trim())).toEqual([
      'Orwell was writing in 1948.',
      'The clocks are the giveaway.',
    ]);

    // Every reference in the body must point at a footnote that exists.
    const doc = await pkg.readXml(pkg.documentPath);
    const referenced = descendants(child(doc?.documentElement ?? null, 'body'), 'footnoteReference')
      .map((el) => attr(el, 'id'));
    expect(referenced).toHaveLength(2);
    const defined = new Set(real.map((el) => attr(el, 'id')));
    for (const id of referenced) expect(defined.has(id!)).toBe(true);
  });

  it('remaps footnote styles the reference does not define', async () => {
    const pkg = await run({
      footnotes: [{ id: 1, text: 'A note.' }],
      paragraphs: [
        { text: 'Chapter One', alignment: 'center' },
        { text: PROSE, footnoteRef: 1 },
      ],
    });

    const notes = await pkg.readXml('word/footnotes.xml');
    const styles = await pkg.readXml('word/styles.xml');
    const known = new Set(
      children(styles?.documentElement ?? null, 'style').map((el) => attr(el, 'styleId')),
    );
    // A dangling style id would make Word fall back silently; there should be none.
    for (const el of [
      ...descendants(notes?.documentElement ?? null, 'pStyle'),
      ...descendants(notes?.documentElement ?? null, 'rStyle'),
    ]) {
      expect(known.has(attr(el, 'val'))).toBe(true);
    }
  });

  it('copies embedded images and repoints the relationship', async () => {
    const pkg = await run({
      image: true,
      paragraphs: [
        { text: 'Chapter One', alignment: 'center' },
        { text: PROSE },
        { text: '', image: true, alignment: 'center' },
      ],
    });

    const doc = await pkg.readXml(pkg.documentPath);
    const blips = descendants(doc?.documentElement ?? null, 'blip', NS.a);
    expect(blips).toHaveLength(1);

    const embedId = attr(blips[0], 'embed', NS.r);
    expect(embedId).not.toBeNull();

    const rels = await pkg.relsFor(pkg.documentPath);
    const rel = rels.byId(embedId!);
    expect(rel).toBeDefined();

    const mediaPath = pkg.resolveTarget(pkg.documentPath, rel!.target);
    expect(pkg.has(mediaPath)).toBe(true);
    const bytes = await pkg.readBinary(mediaPath);
    expect(bytes).not.toBeNull();
    expect(Buffer.from(bytes!)).toEqual(TEST_PNG);

    const types = await pkg.readText('[Content_Types].xml');
    expect(types).toContain('image/png');
  });

  it('keeps an image paragraph centred rather than indenting it', async () => {
    const pkg = await run({
      image: true,
      paragraphs: [
        { text: 'Chapter One', alignment: 'center' },
        { text: PROSE },
        { text: '', image: true },
      ],
    });

    const doc = await pkg.readXml(pkg.documentPath);
    const body = child(doc?.documentElement ?? null, 'body');
    const withImage = children(body, 'p').find((p) => descendants(p, 'drawing').length > 0);
    expect(withImage).toBeDefined();
    const pPr = child(withImage!, 'pPr');
    expect(attr(child(pPr, 'jc'), 'val')).toBe('center');
    expect(attr(child(pPr, 'ind'), 'firstLine')).toBe('0');
  });

  it('drops comments and rejected edits but keeps accepted insertions', async () => {
    const id = counter++;
    const reference = await writeDocx(join(dir, `tref${id}.docx`), BOOK_TEMPLATE);
    const manuscript = join(dir, `tms${id}.docx`);
    // Hand-built so the paragraph can carry tracked-change markup.
    const { default: JSZip } = await import('jszip');
    const base = await JSZip.loadAsync(
      await (await import('node:fs/promises')).readFile(
        await writeDocx(join(dir, `tbase${id}.docx`), {
          paragraphs: [{ text: 'Chapter One', alignment: 'center' }, { text: PROSE }],
        }),
      ),
    );
    const docXml = await base.file('word/document.xml')!.async('string');
    base.file(
      'word/document.xml',
      docXml.replace(
        '</w:body>',
        '<w:p><w:commentRangeStart w:id="1"/>' +
          '<w:ins w:id="2" w:author="Editor"><w:r><w:t xml:space="preserve">Kept insertion. </w:t></w:r></w:ins>' +
          '<w:del w:id="3" w:author="Editor"><w:r><w:delText>Deleted text.</w:delText></w:r></w:del>' +
          '<w:commentRangeEnd w:id="1"/><w:r><w:t xml:space="preserve">Plain tail.</w:t></w:r>' +
          '</w:p></w:body>',
      ),
    );
    await (await import('node:fs/promises')).writeFile(
      manuscript,
      await base.generateAsync({ type: 'nodebuffer' }),
    );

    const output = join(dir, `tout${id}.docx`);
    await formatManuscript({ referencePath: reference, manuscriptPath: manuscript, outputPath: output });

    const pkg = await loadDocx(output);
    const doc = await pkg.readXml(pkg.documentPath);
    const body = child(doc?.documentElement ?? null, 'body');
    const last = children(body, 'p').at(-1)!;

    expect(textOf(last)).toBe('Kept insertion. Plain tail.');
    expect(descendants(body, 'ins')).toHaveLength(0);
    expect(descendants(body, 'del')).toHaveLength(0);
    expect(descendants(body, 'commentRangeStart')).toHaveLength(0);
  });
});
