import { sectionHasContent } from '../core/build/matter.js';
import type { BookDetails, ExtraSections, FormatOptions } from '../core/types.js';

/**
 * The book-details form. Every section is off until the author asks for it,
 * and a section that would come out empty says so instead of quietly producing
 * a headed blank page.
 */

export interface DetailsContext {
  options: FormatOptions;
  /**
   * `structural` means the form's shape changed — a section was switched on or
   * off — and needs redrawing. Typing must never trigger it, or the field
   * being typed into would be replaced and lose focus mid-word.
   */
  onChange: (structural: boolean) => void;
}

interface FieldSpec {
  key: keyof BookDetails;
  label: string;
  hint?: string;
  placeholder?: string;
  multiline?: boolean;
  width?: 'full' | 'half';
}

const IDENTITY_FIELDS: FieldSpec[] = [
  { key: 'title', label: 'Book title', placeholder: 'The title as it should appear' },
  { key: 'subtitle', label: 'Subtitle', hint: 'Leave blank if your book has none.' },
  { key: 'author', label: 'Author name', placeholder: 'As you want it printed' },
  { key: 'copyrightYear', label: 'Copyright year', placeholder: '2026', width: 'half' },
  { key: 'publisher', label: 'Publisher', hint: 'Your own imprint is fine.', width: 'half' },
  {
    key: 'isbn',
    label: 'ISBN',
    hint: 'Only if you bought your own. KDP gives you a free one otherwise.',
    width: 'half',
  },
];

interface SectionSpec {
  key: keyof ExtraSections;
  label: string;
  hint: string;
  /** The details field holding this section's words, if it needs any. */
  textKey?: keyof BookDetails;
  textLabel?: string;
  placeholder?: string;
}

const SECTIONS: SectionSpec[] = [
  {
    key: 'alsoBy',
    label: 'Also by this author',
    hint: 'A list of your earlier books, on the page before the title page.',
    textKey: 'alsoBy',
    textLabel: 'Earlier books',
    placeholder: 'One title per line.',
  },
  {
    key: 'titlePage',
    label: 'Title page',
    hint: 'Your title, subtitle and name, on a page of their own.',
  },
  {
    key: 'copyrightPage',
    label: 'Copyright page',
    hint: 'Built from the year, author, publisher and ISBN above.',
  },
  {
    key: 'dedication',
    label: 'Dedication',
    hint: 'A short page near the front.',
    textKey: 'dedication',
    textLabel: 'Dedication',
    placeholder: 'For everyone who waited.',
  },
  {
    key: 'epigraph',
    label: 'Epigraph',
    hint: 'A short quotation to open the book, after the dedication.',
    textKey: 'epigraph',
    textLabel: 'Epigraph',
    placeholder: 'The quotation, then who said it on the last line.',
  },
  {
    key: 'contents',
    label: 'Table of contents',
    hint: 'Word fills in the chapter names and page numbers when you open the file.',
  },
  {
    key: 'acknowledgments',
    label: 'Acknowledgments',
    hint: 'Goes at the back of the book.',
    textKey: 'acknowledgments',
    textLabel: 'Acknowledgments',
    placeholder: 'Thank the people who helped. Leave a blank line between paragraphs.',
  },
  {
    key: 'aboutTheAuthor',
    label: 'About the author',
    hint: 'A short biography at the back.',
    textKey: 'aboutTheAuthor',
    textLabel: 'About the author',
    placeholder: 'A sentence or two about yourself.',
  },
  {
    key: 'bibliography',
    label: 'Bibliography',
    hint: 'One entry per line.',
    textKey: 'bibliography',
    textLabel: 'Bibliography',
    placeholder: 'One source per line.',
  },
];

export function renderDetailsForm(container: HTMLElement, ctx: DetailsContext): void {
  const details = ctx.options.bookDetails;
  const sections = ctx.options.extraSections;
  container.replaceChildren();

  const identity = document.createElement('div');
  identity.className = 'details-fields';
  for (const field of IDENTITY_FIELDS) identity.appendChild(textField(field, ctx));
  container.appendChild(identity);

  container.appendChild(runningHeadsControl(ctx));

  const list = document.createElement('div');
  list.className = 'details-sections';
  const heading = document.createElement('h4');
  heading.textContent = 'Pages to add';
  list.appendChild(heading);

  for (const spec of SECTIONS) {
    list.appendChild(sectionRow(spec, ctx, details, sections));
  }
  container.appendChild(list);

  // Replacing is only meaningful once a generated page could stand in.
  const anyGenerated = SECTIONS.some(
    (s) => sections[s.key] && sectionHasContent(s.key, details),
  );
  if (anyGenerated) {
    const replace = document.createElement('div');
    replace.className = 'option-row details-replace';
    const input = document.createElement('input');
    input.type = 'checkbox';
    input.checked = ctx.options.replaceFrontMatter;
    input.addEventListener('change', () => {
      ctx.options.replaceFrontMatter = input.checked;
      ctx.onChange(false);
    });
    const label = document.createElement('label');
    label.append(input, ' ', 'Use these instead of the front pages already in my manuscript');
    const note = document.createElement('span');
    note.className = 'option-note';
    note.textContent =
      'Switch this on if your manuscript already has a title or copyright page, so you do not end up with two.';
    label.appendChild(note);
    replace.appendChild(label);
    container.appendChild(replace);
  }
}

function textField(spec: FieldSpec, ctx: DetailsContext): HTMLElement {
  const wrap = document.createElement('label');
  wrap.className = `details-field${spec.width === 'half' ? ' half' : ''}`;

  const label = document.createElement('span');
  label.className = 'details-label';
  label.textContent = spec.label;
  wrap.appendChild(label);

  const input = spec.multiline
    ? document.createElement('textarea')
    : document.createElement('input');
  if (input instanceof HTMLInputElement) input.type = 'text';
  if (input instanceof HTMLTextAreaElement) input.rows = 3;
  input.value = ctx.options.bookDetails[spec.key];
  if (spec.placeholder) input.placeholder = spec.placeholder;
  input.addEventListener('input', () => {
    ctx.options.bookDetails = { ...ctx.options.bookDetails, [spec.key]: input.value };
    ctx.onChange(false);
  });
  wrap.appendChild(input);

  if (spec.hint) {
    const hint = document.createElement('span');
    hint.className = 'details-hint';
    hint.textContent = spec.hint;
    wrap.appendChild(hint);
  }
  return wrap;
}

function sectionRow(
  spec: SectionSpec,
  ctx: DetailsContext,
  details: BookDetails,
  sections: ExtraSections,
): HTMLElement {
  const row = document.createElement('div');
  row.className = 'details-section';

  const label = document.createElement('label');
  label.className = 'details-toggle';
  const input = document.createElement('input');
  input.type = 'checkbox';
  input.checked = sections[spec.key];
  input.addEventListener('change', () => {
    ctx.options.extraSections = { ...ctx.options.extraSections, [spec.key]: input.checked };
    ctx.onChange(true);
  });
  const title = document.createElement('span');
  title.className = 'details-toggle-title';
  title.textContent = spec.label;
  const hint = document.createElement('span');
  hint.className = 'option-note';
  hint.textContent = spec.hint;
  label.append(input, ' ', title, hint);
  row.appendChild(label);

  if (!sections[spec.key]) return row;

  if (spec.textKey && spec.textLabel) {
    const area = document.createElement('textarea');
    area.className = 'details-textarea';
    area.rows = 3;
    area.value = details[spec.textKey];
    area.placeholder = spec.placeholder ?? '';
    area.setAttribute('aria-label', spec.textLabel);
    area.addEventListener('input', () => {
      ctx.options.bookDetails = { ...ctx.options.bookDetails, [spec.textKey!]: area.value };
      ctx.onChange(false);
    });
    row.appendChild(area);
  }

  // Say plainly when a switched-on section would produce nothing.
  if (!sectionHasContent(spec.key, details)) {
    const warn = document.createElement('p');
    warn.className = 'details-empty';
    warn.textContent =
      spec.key === 'titlePage'
        ? 'Add a book title above and this page will be built.'
        : spec.key === 'copyrightPage'
          ? 'Add a copyright year or author above and this page will be built.'
          : 'Nothing typed yet, so this page will be left out.';
    row.appendChild(warn);
  }
  return row;
}

/**
 * What the small line at the top of each page says. Most designs ship with
 * placeholder wording, so the default corrects it from the details above;
 * an author who wants something else can set both sides outright.
 */
function runningHeadsControl(ctx: DetailsContext): HTMLElement {
  const heads = ctx.options.runningHeads;
  const wrap = document.createElement('div');
  wrap.className = 'details-heads';

  const heading = document.createElement('h4');
  heading.textContent = 'Page headers';
  wrap.appendChild(heading);

  const hint = document.createElement('p');
  hint.className = 'panel-hint';
  hint.textContent =
    'The small line printed at the top of each page. Printed books put the author on ' +
    'left-hand pages and the book title on right-hand ones.';
  wrap.appendChild(hint);

  const select = document.createElement('select');
  for (const [value, label] of [
    ['auto', 'Use my title and author from above'],
    ['custom', 'Say something else'],
    ['leave', 'Leave whatever the design says'],
  ] as const) {
    const option = document.createElement('option');
    option.value = value;
    option.textContent = label;
    option.selected = heads.mode === value;
    select.appendChild(option);
  }
  select.addEventListener('change', () => {
    ctx.options.runningHeads = {
      ...ctx.options.runningHeads,
      mode: select.value as typeof heads.mode,
    };
    ctx.onChange(false);
  });

  const row = document.createElement('div');
  row.className = 'option-row';
  const label = document.createElement('label');
  label.textContent = 'Headers';
  row.append(label, select);
  wrap.appendChild(row);

  if (heads.mode === 'custom') {
    const fields = document.createElement('div');
    fields.className = 'details-fields';
    fields.append(
      headField('Left-hand pages', heads.verso, ctx.options.bookDetails.author, (value) => {
        ctx.options.runningHeads = { ...ctx.options.runningHeads, verso: value };
        ctx.onChange(false);
      }),
      headField('Right-hand pages', heads.recto, ctx.options.bookDetails.title, (value) => {
        ctx.options.runningHeads = { ...ctx.options.runningHeads, recto: value };
        ctx.onChange(false);
      }),
    );
    wrap.appendChild(fields);
  }
  return wrap;
}

function headField(
  label: string,
  value: string,
  placeholder: string,
  onInput: (value: string) => void,
): HTMLElement {
  const wrap = document.createElement('label');
  wrap.className = 'details-field';
  const name = document.createElement('span');
  name.textContent = label;
  const input = document.createElement('input');
  input.type = 'text';
  input.value = value;
  input.placeholder = placeholder || 'leave empty for nothing';
  input.addEventListener('input', () => onInput(input.value));
  wrap.append(name, input);
  return wrap;
}
