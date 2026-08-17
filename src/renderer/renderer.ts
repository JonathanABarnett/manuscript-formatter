import { twipsToInches } from '../core/ooxml/ns.js';
import {
  DEFAULT_FORMAT_OPTIONS,
  ROLE_HINTS,
  ROLE_LABELS,
  STYLE_ROLES,
  type BlockRole,
  type FormatOptions,
  type FormatResult,
  type ManuscriptBlock,
  type ReferenceProfile,
  type StyleRole,
} from '../core/types.js';
import type { AnalysisResult } from '../shared/ipc.js';
import { renderPreviews } from './previews.js';

interface State {
  referencePath: string | null;
  manuscriptPath: string | null;
  analysis: AnalysisResult | null;
  options: FormatOptions;
  outputPath: string | null;
  showAllBlocks: boolean;
  /** Rendered width of a preview page, in CSS pixels. */
  previewWidth: number;
  busy: boolean;
  result: FormatResult | null;
}

const state: State = {
  referencePath: null,
  manuscriptPath: null,
  analysis: null,
  options: { ...DEFAULT_FORMAT_OPTIONS },
  outputPath: null,
  showAllBlocks: false,
  previewWidth: 380,
  busy: false,
  result: null,
};

/** Beyond this the list stops being a review aid and starts being a scroll. */
const ALL_BLOCKS_LIMIT = 2500;

/** Roles a reviewer can assign by hand, in the order they read sensibly. */
const ASSIGNABLE_ROLES: BlockRole[] = [
  'chapterTitle',
  'partTitle',
  'chapterSubtitle',
  'subheading',
  'bodyFirst',
  'body',
  'blockQuote',
  'listItem',
  'sceneBreak',
  'frontMatterTitle',
  'frontMatter',
  'empty',
];

// --- tiny DOM helpers -------------------------------------------------------

function $<T extends HTMLElement>(selector: string): T {
  const found = document.querySelector<T>(selector);
  if (!found) throw new Error(`missing element: ${selector}`);
  return found;
}

type Child = Node | string | null | undefined | false;

function el<K extends keyof HTMLElementTagNameMap>(
  tag: K,
  props: Partial<HTMLElementTagNameMap[K]> & { class?: string; dataset?: Record<string, string> } = {},
  ...kids: Child[]
): HTMLElementTagNameMap[K] {
  const node = document.createElement(tag);
  const { class: className, dataset, ...rest } = props as Record<string, unknown> & {
    class?: string;
    dataset?: Record<string, string>;
  };
  if (className) node.className = className;
  if (dataset) for (const [k, v] of Object.entries(dataset)) node.dataset[k] = v;
  Object.assign(node, rest);
  for (const kid of kids) {
    if (kid === null || kid === undefined || kid === false) continue;
    node.appendChild(typeof kid === 'string' ? document.createTextNode(kid) : kid);
  }
  return node;
}

function replace(container: HTMLElement, ...kids: Child[]): void {
  container.replaceChildren();
  for (const kid of kids) {
    if (kid === null || kid === undefined || kid === false) continue;
    container.appendChild(typeof kid === 'string' ? document.createTextNode(kid) : kid);
  }
}

function baseName(path: string): string {
  const parts = path.split(/[\\/]/);
  return parts[parts.length - 1] || path;
}

function inches(twips: number): string {
  const value = Math.round(twipsToInches(twips) * 100) / 100;
  return `${value}"`;
}

/**
 * A term with a plain-English explanation on hover or keyboard focus. Used for
 * the handful of printing words an author may not have met before.
 */
function hinted(text: string, hint: string): HTMLElement {
  return el('span', {
    class: 'hinted',
    tabIndex: 0,
    textContent: text,
    dataset: { hint },
    // Announced by screen readers, which cannot see the hover bubble.
    ariaLabel: `${text}: ${hint}`,
  });
}

function showError(message: string | null): void {
  const banner = $('#error-banner');
  if (!message) {
    banner.hidden = true;
    return;
  }
  banner.textContent = message;
  banner.hidden = false;
  window.setTimeout(() => {
    if (banner.textContent === message) banner.hidden = true;
  }, 9000);
}

// --- step 1: choosing files -------------------------------------------------

function setupPicker(kind: 'reference' | 'manuscript'): void {
  const picker = $(`#picker-${kind}`);
  const zone = picker.querySelector<HTMLElement>('.dropzone');
  if (!zone) return;

  const choose = async (): Promise<void> => {
    const path = await window.formatter.pickDocx(kind);
    if (path) void setPath(kind, path);
  };

  zone.addEventListener('click', () => void choose());
  zone.addEventListener('keydown', (event) => {
    if (event.key === 'Enter' || event.key === ' ') {
      event.preventDefault();
      void choose();
    }
  });

  for (const type of ['dragenter', 'dragover'] as const) {
    zone.addEventListener(type, (event) => {
      event.preventDefault();
      zone.classList.add('dragging');
    });
  }
  for (const type of ['dragleave', 'drop'] as const) {
    zone.addEventListener(type, () => zone.classList.remove('dragging'));
  }
  zone.addEventListener('drop', (event) => {
    event.preventDefault();
    const file = event.dataTransfer?.files?.[0];
    if (!file) return;
    const path = window.formatter.pathForFile(file);
    if (!path) {
      showError('That file could not be read. Try the browse button instead.');
      return;
    }
    if (!path.toLowerCase().endsWith('.docx')) {
      showError('Only .docx files are supported. Open the file in Word and save it as .docx.');
      return;
    }
    void setPath(kind, path);
  });
}

async function setPath(kind: 'reference' | 'manuscript', path: string): Promise<void> {
  if (kind === 'reference') state.referencePath = path;
  else state.manuscriptPath = path;

  // Any previous analysis and output belong to the old pair of files.
  state.analysis = null;
  state.result = null;
  if (kind === 'manuscript') state.outputPath = null;
  $('#step-review').hidden = true;
  $('#step-output').hidden = true;
  $('#result').hidden = true;

  const zone = $(`#picker-${kind}`).querySelector<HTMLElement>('.dropzone');
  const label = zone?.querySelector<HTMLElement>('.dropzone-label');
  const file = zone?.querySelector<HTMLElement>('.dropzone-file');
  zone?.classList.add('chosen');
  if (label) label.textContent = 'Selected';
  if (file) {
    file.textContent = baseName(path);
    file.hidden = false;
    file.title = path;
  }

  if (state.referencePath && state.manuscriptPath) await runAnalysis();
}

async function runAnalysis(): Promise<void> {
  const status = $('#analyze-status');
  status.className = 'status spinner';
  status.textContent = 'Reading your two files';
  showError(null);

  const outcome = await window.formatter.analyze({
    referencePath: state.referencePath!,
    manuscriptPath: state.manuscriptPath!,
  });

  status.className = 'status';
  if (!outcome.ok) {
    status.textContent = '';
    showError(outcome.error);
    return;
  }

  state.analysis = outcome.value;
  state.options = { ...outcome.value.suggestedOptions, roleStyles: {}, roleOverrides: {} };
  const found = outcome.value.analysis;
  status.textContent =
    `Read ${found.wordCount.toLocaleString()} words and found ${found.chapterCount} ` +
    `chapter${found.chapterCount === 1 ? '' : 's'}. Have a look below.`;

  if (!state.outputPath) {
    state.outputPath = await window.formatter.suggestOutput(state.manuscriptPath!);
  }

  renderReview();
  renderOutput();
  $('#step-review').hidden = false;
  $('#step-output').hidden = false;
  $('#step-review').scrollIntoView({ behavior: 'smooth', block: 'start' });
}

// --- step 2: review ---------------------------------------------------------

function renderReview(): void {
  if (!state.analysis) return;
  renderTemplateFacts(state.analysis.profile);
  renderOptions();
  refreshPreviews();
  renderStyleMap(state.analysis.profile);
  renderStructure();
}

/**
 * Redraw the page proofs. Called whenever anything they depend on changes — a
 * style mapping, a paragraph's role, or an option that affects layout — so the
 * proof always matches what pressing Format would produce.
 */
function refreshPreviews(): void {
  if (!state.analysis) return;
  renderPreviews($('#previews'), {
    profile: state.analysis.profile,
    blocks: state.analysis.analysis.blocks,
    options: state.options,
    pageWidthPx: state.previewWidth,
    onRoleStyleChange: (role, styleId) => {
      state.options.roleStyles[role] = styleId;
      refreshPreviews();
      renderStyleMap(state.analysis!.profile);
    },
  });
}

function renderTemplateFacts(profile: ReferenceProfile): void {
  const m = profile.pageSetup.margins;
  /** label, plain-English explanation, value. */
  const rows: Array<[string, string, string]> = [
    ['Book size', 'The finished page size of your printed book.', profile.pageSizeLabel],
    [
      'Margins',
      'The white space around your text on every page.',
      `${inches(m.top)} top and bottom, ${inches(m.left)} and ${inches(m.right)} at the sides`,
    ],
    [
      'Binding margin',
      'Extra space on the inside edge so your words are not swallowed by the spine.',
      m.gutter > 0 ? inches(m.gutter) : 'none — built into the side margins',
    ],
    [
      'Facing pages',
      'Left and right pages mirror each other, the way an open book does.',
      profile.pageSetup.mirrorMargins ? 'yes' : 'no',
    ],
    [
      'Main text',
      'The typeface and size used for ordinary paragraphs.',
      [profile.bodyFontName, profile.bodyFontSizePt ? `${profile.bodyFontSizePt} pt` : null]
        .filter(Boolean)
        .join(', ') || "the design's default",
    ],
    [
      'Paragraph indent',
      'How far the first line of each paragraph is pushed in from the margin.',
      profile.bodyFirstLineIndentTwips
        ? inches(profile.bodyFirstLineIndentTwips)
        : 'none — paragraphs start at the margin',
    ],
    [
      'Space above chapters',
      'Blank lines the design puts above a chapter title to push it down the page.',
      profile.chapterTitleBlanksBefore === 0
        ? 'none'
        : `${profile.chapterTitleBlanksBefore} blank line${profile.chapterTitleBlanksBefore === 1 ? '' : 's'}`,
    ],
    [
      'Page headers',
      'The small line of text at the top of each page — often the book or chapter name.',
      profile.hasHeaders ? 'yes, kept as they are' : 'none',
    ],
    [
      'Page numbers',
      'Numbers printed at the top or bottom of each page.',
      profile.hasPageNumbers
        ? 'yes, kept as they are'
        : profile.hasFooters
          ? 'there is a footer, but no automatic number'
          : 'none',
    ],
    [
      'Named designs',
      'Ready-made looks stored in the file that parts of your book can use.',
      `${profile.styles.length} to choose from`,
    ],
  ];

  replace(
    $('#template-facts'),
    ...rows.flatMap(([term, hint, value]) => [
      el('dt', {}, hinted(term, hint)),
      el('dd', {}, value),
    ]),
  );
}

function renderOptions(): void {
  const o = state.options;
  const profile = state.analysis?.profile;

  const chapterStart = el(
    'select',
    { id: 'opt-chapter-start' },
    el('option', { value: 'newPage', selected: o.chapterStart === 'newPage' }, 'On a new page'),
    el(
      'option',
      { value: 'oddPage', selected: o.chapterStart === 'oddPage' },
      'Always on a right-hand page',
    ),
    el(
      'option',
      { value: 'continuous', selected: o.chapterStart === 'continuous' },
      'Straight after the last chapter',
    ),
  );
  chapterStart.addEventListener('change', () => {
    o.chapterStart = chapterStart.value as FormatOptions['chapterStart'];
    refreshPreviews();
  });

  const sceneBreak = el('input', {
    type: 'text',
    value: o.sceneBreakText ?? '',
    placeholder: 'leave as you wrote it',
    size: 18,
  });
  sceneBreak.addEventListener('input', () => {
    o.sceneBreakText = sceneBreak.value.trim() === '' ? null : sceneBreak.value;
    refreshPreviews();
  });

  /** Blank lines above or below a chapter title, or "follow the design". */
  const spacing = (
    key: 'chapterSpaceBefore' | 'chapterSpaceAfter',
    fromDesign: number,
  ): HTMLSelectElement => {
    const select = el(
      'select',
      {},
      el(
        'option',
        { value: '', selected: o[key] === null },
        `Follow the design (${fromDesign})`,
      ),
    );
    for (let n = 0; n <= 12; n++) {
      select.appendChild(
        el(
          'option',
          { value: String(n), selected: o[key] === n },
          n === 0 ? 'None' : `${n} blank line${n === 1 ? '' : 's'}`,
        ),
      );
    }
    select.addEventListener('change', () => {
      o[key] = select.value === '' ? null : Number(select.value);
      refreshPreviews();
    });
    return select;
  };

  const checkbox = (key: keyof FormatOptions, label: string, note: string): HTMLElement => {
    const input = el('input', { type: 'checkbox', checked: Boolean(o[key]) });
    input.addEventListener('change', () => {
      (o[key] as unknown as boolean) = input.checked;
      refreshPreviews();
    });
    return el(
      'div',
      { class: 'option-row' },
      el('label', {}, input, ' ', label, el('span', { class: 'option-note' }, note)),
    );
  };

  const row = (label: Node | string, control: Node, note?: string): HTMLElement =>
    el(
      'div',
      { class: 'option-row' },
      el('label', {}, label, note ? el('span', { class: 'option-note' }, note) : null),
      control,
    );

  replace(
    $('#options'),
    row(
      el('span', {}, 'Where chapters begin'),
      chapterStart,
      o.chapterStart === 'oddPage'
        ? 'Like most printed books. A blank page is added when one is needed to land on the right.'
        : undefined,
    ),
    row(
      hinted('Space above chapter titles', 'Blank lines that push a chapter title down the page.'),
      spacing('chapterSpaceBefore', profile?.chapterTitleBlanksBefore ?? 0),
    ),
    row(
      el('span', {}, 'Space below chapter titles'),
      spacing('chapterSpaceAfter', profile?.chapterTitleBlanksAfter ?? 0),
    ),
    row(
      el('span', {}, 'Mark between scenes'),
      sceneBreak,
      'Replaces whatever you used — # or *** — everywhere at once.',
    ),
    checkbox(
      'firstParagraphNoIndent',
      "Don't indent the first paragraph",
      'Printed books normally start a chapter, and each new scene, hard against the margin.',
    ),
    checkbox(
      'removeEmptyParagraphs',
      'Remove empty lines',
      'Spacing comes from the design instead, which keeps it even throughout. Recommended.',
    ),
    checkbox(
      'removeManualIndents',
      'Remove tabbed indents',
      'Takes out tabs and spaces typed at the start of paragraphs, so indents stay identical.',
    ),
    checkbox('keepEmphasis', 'Keep italics and bold', 'Any emphasis in your writing is preserved.'),
    checkbox(
      'includeFrontMatter',
      'Include your title and copyright pages',
      'Turn off to start the book at chapter one.',
    ),
    checkbox(
      'smartTypography',
      'Use curly quotes and proper dashes',
      'Turns " into “ ”, -- into a dash, ... into an ellipsis. This changes characters in your text, so leave it off if you are unsure.',
    ),
    checkbox(
      'collapseMultipleSpaces',
      'Reduce double spaces to single',
      'The two spaces after a full stop become one, as printed books use.',
    ),
  );
}

function renderStyleMap(profile: ReferenceProfile): void {
  const rows = STYLE_ROLES.map((role) => {
    const current = state.options.roleStyles[role] ?? profile.roleStyles[role] ?? '';
    const select = el('select', {}, el('option', { value: '' }, 'Choose for me'));
    for (const style of profile.styles) {
      select.appendChild(
        el(
          'option',
          { value: style.id, selected: style.id === current },
          style.usageCount > 0 ? `${style.name} (used ${style.usageCount}×)` : style.name,
        ),
      );
    }
    select.addEventListener('change', () => {
      state.options.roleStyles[role] = select.value === '' ? null : select.value;
      refreshPreviews();
      updateResetButton();
    });

    return el(
      'div',
      { class: 'style-row' },
      el('span', { class: 'style-label' }, ROLE_LABELS[role]),
      el('span', { class: 'style-hint' }, ROLE_HINTS[role]),
      select,
      el('span', { class: 'style-evidence' }, whyChosen(profile, role)),
    );
  });

  replace($('#style-map'), ...rows);
  updateResetButton();
}

/** Plain-English reason a design was picked for a part of the book. */
function whyChosen(profile: ReferenceProfile, role: StyleRole): string {
  if (state.options.roleStyles[role] !== undefined && state.options.roleStyles[role] !== null) {
    return 'You chose this one.';
  }
  const evidence = profile.roleEvidence[role];
  if (evidence) return `Picked because the ${evidence}.`;
  return profile.roleStyles[role]
    ? 'Picked from your design file.'
    : 'Your design has nothing for this, so the closest match is used.';
}

/** The undo button only makes sense once something has been changed. */
function updateResetButton(): void {
  const changed =
    Object.values(state.options.roleStyles).some((v) => v !== undefined && v !== null) ||
    Object.keys(state.options.roleOverrides).length > 0;
  $('#reset-roles').hidden = !changed;
}

function resetRoles(): void {
  state.options.roleStyles = {};
  state.options.roleOverrides = {};
  if (state.analysis) renderStyleMap(state.analysis.profile);
  renderStructure();
  refreshPreviews();
}

/** Rows worth a reviewer's attention: structure, plus anything uncertain. */
function isInteresting(block: ManuscriptBlock): boolean {
  if (block.role === 'empty' || block.role === 'body' || block.role === 'bodyFirst') {
    return block.confidence < 0.6;
  }
  return block.role !== 'pageBreak';
}

function renderStructure(): void {
  const analysis = state.analysis?.analysis;
  if (!analysis) return;

  const counts = new Map<BlockRole, number>();
  for (const block of analysis.blocks) {
    const role = state.options.roleOverrides[block.index] ?? block.role;
    counts.set(role, (counts.get(role) ?? 0) + 1);
  }

  const stat = (label: string, value: number | string): HTMLElement =>
    el('span', { class: 'stat' }, el('strong', {}, String(value)), ` ${label}`);

  replace(
    $('#structure-summary'),
    stat('words', analysis.wordCount.toLocaleString()),
    stat('paragraphs', analysis.paragraphCount.toLocaleString()),
    stat('chapters', (counts.get('chapterTitle') ?? 0) + (counts.get('partTitle') ?? 0)),
    stat('scene breaks', counts.get('sceneBreak') ?? 0),
    analysis.tableCount > 0 ? stat('tables', analysis.tableCount) : null,
    analysis.imageCount > 0 ? stat('images', analysis.imageCount) : null,
    analysis.footnoteCount > 0 ? stat('footnotes', analysis.footnoteCount) : null,
  );

  const needsReview = analysis.blocks.filter(
    (b) => b.confidence < 0.6 && b.role !== 'body' && b.role !== 'empty',
  ).length;
  $('#structure-hint').textContent = state.showAllBlocks
    ? 'Every paragraph in your book. Use the menu on the right to change what any line is.'
    : needsReview > 0
      ? `Your chapters, headings and scene breaks. ${needsReview} line${
          needsReview === 1 ? ' is' : 's are'
        } marked "not sure" — worth a glance. Tick the box above to see everything.`
      : 'Your chapters, headings and scene breaks. Use the menu on the right to change what any line is.';

  const candidates = state.showAllBlocks
    ? analysis.blocks.filter((b) => b.role !== 'empty')
    : analysis.blocks.filter(isInteresting);
  const shown = candidates.slice(0, ALL_BLOCKS_LIMIT);

  const list = $('#structure-list');
  if (shown.length === 0) {
    replace(
      list,
      el(
        'p',
        { class: 'block-empty' },
        'No chapters or headings were recognised. Tick "Show every paragraph" above, then mark ' +
          'your chapter titles by hand using the menus.',
      ),
    );
    return;
  }

  replace(
    list,
    ...shown.map((block) => renderBlockRow(block)),
    candidates.length > shown.length
      ? el(
          'p',
          { class: 'list-note' },
          `Showing the first ${shown.length.toLocaleString()} of ${candidates.length.toLocaleString()} paragraphs.`,
        )
      : null,
  );
}

function renderBlockRow(block: ManuscriptBlock): HTMLElement {
  const override = state.options.roleOverrides[block.index];
  const role = override ?? block.role;
  const needsReview = block.confidence < 0.6 && block.role !== 'body';

  const select = el('select', {});
  for (const option of ASSIGNABLE_ROLES) {
    select.appendChild(
      el('option', { value: option, selected: option === role }, ROLE_LABELS[option]),
    );
  }

  const row = el(
    'div',
    {
      class: `block-row${needsReview ? ' needs-review' : ''}${override ? ' overridden' : ''}`,
      dataset: { index: String(block.index) },
    },
    el('span', { class: 'block-index' }, `#${block.index}`),
    el(
      'span',
      {
        class: 'block-text',
        title: block.reasons.length > 0 ? `Recognised because it ${block.reasons.join('; ')}` : undefined,
      },
      block.preview || '(empty line)',
      needsReview ? el('span', { class: 'flag' }, 'not sure') : null,
      override ? el('span', { class: 'flag changed' }, 'changed by you') : null,
    ),
    select,
  );

  select.addEventListener('change', () => {
    const chosen = select.value as BlockRole;
    if (chosen === block.autoRole) delete state.options.roleOverrides[block.index];
    else state.options.roleOverrides[block.index] = chosen;
    row.classList.toggle('overridden', chosen !== block.autoRole);
    row.classList.remove('needs-review');
    renderStructure();
    refreshPreviews();
    updateResetButton();
  });

  return row;
}

// --- step 3: output ---------------------------------------------------------

function renderOutput(): void {
  const target = $('#output-path');
  target.textContent = state.outputPath ?? '';
  target.title = state.outputPath ?? '';
}

async function chooseOutput(): Promise<void> {
  const chosen = await window.formatter.pickOutput(
    state.outputPath ?? (await window.formatter.suggestOutput(state.manuscriptPath!)),
  );
  if (!chosen) return;
  state.outputPath = chosen;
  renderOutput();
}

async function runFormat(): Promise<void> {
  if (state.busy || !state.referencePath || !state.manuscriptPath || !state.outputPath) return;

  const button = $<HTMLButtonElement>('#run-format');
  state.busy = true;
  button.disabled = true;
  button.classList.add('spinner');
  button.textContent = 'Making your book';
  showError(null);

  const outcome = await window.formatter.format({
    referencePath: state.referencePath,
    manuscriptPath: state.manuscriptPath,
    outputPath: state.outputPath,
    options: state.options,
  });

  state.busy = false;
  button.disabled = false;
  button.classList.remove('spinner');
  button.textContent = 'Make my formatted book';

  if (!outcome.ok) {
    showError(outcome.error);
    return;
  }
  state.result = outcome.value;
  renderResult(outcome.value);
}

function renderResult(result: FormatResult): void {
  const s = result.stats;
  const line = [
    `${s.wordCount.toLocaleString()} words`,
    `${s.chapters} chapter${s.chapters === 1 ? '' : 's'}`,
    s.sceneBreaks > 0 ? `${s.sceneBreaks} scene break${s.sceneBreaks === 1 ? '' : 's'}` : null,
    s.tables > 0 ? `${s.tables} table${s.tables === 1 ? '' : 's'}` : null,
    s.imagesCopied > 0 ? `${s.imagesCopied} picture${s.imagesCopied === 1 ? '' : 's'}` : null,
    s.footnotesCopied > 0 ? `${s.footnotesCopied} footnotes` : null,
  ]
    .filter(Boolean)
    .join(' · ');

  const onWeb = window.formatter.platform === 'web';
  const open = el(
    'button',
    { class: 'primary', type: 'button' },
    onWeb ? 'Download document' : 'Open in Word',
  );
  open.addEventListener('click', () => void window.formatter.open(result.outputPath));
  // A browser has no folder to reveal, so that action is simply not offered.
  const reveal = onWeb ? null : el('button', { class: 'secondary', type: 'button' }, 'Show in folder');
  reveal?.addEventListener('click', () => void window.formatter.reveal(result.outputPath));

  const panel = $('#result');
  replace(
    panel,
    el('h3', {}, 'Your book is ready'),
    el('p', { class: 'path' }, result.outputPath),
    el('p', {}, line),
    el(
      'p',
      { class: 'result-next' },
      'Open it in Word and flick through before you upload it. Your two original files are ' +
        'untouched, so you can change anything above and make it again.',
    ),
    result.warnings.length > 0
      ? el(
          'div',
          { class: 'warnings' },
          el('h4', {}, 'A few things worth a look'),
          el('ul', {}, ...result.warnings.map((w) => el('li', {}, w))),
        )
      : null,
    el('div', { class: 'result-actions' }, open, reveal),
  );
  if (onWeb) {
    const label = $('#output-path');
    label.textContent = result.outputPath;
  }
  panel.hidden = false;
  panel.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
}

// --- wiring -----------------------------------------------------------------

export function init(): void {
  const onWeb = window.formatter.platform === 'web';
  if (onWeb) {
    // In a browser there is no save dialog and nothing leaves the machine;
    // say so plainly rather than offering a folder picker that cannot work.
    $('#change-output').hidden = true;
    $('#output-path-label').textContent = 'Downloads as';
    $('#reassurance').textContent =
      'Everything runs inside this page. Your documents are never uploaded, and neither ' +
      'source file is modified — a new document is generated for you to download.';
  }

  setupPicker('reference');
  setupPicker('manuscript');

  $('#change-output').addEventListener('click', () => void chooseOutput());
  $('#run-format').addEventListener('click', () => void runFormat());

  const showAll = $<HTMLInputElement>('#show-all-blocks');
  showAll.addEventListener('change', () => {
    state.showAllBlocks = showAll.checked;
    renderStructure();
  });

  $('#reset-roles').addEventListener('click', resetRoles);

  const zoom = $<HTMLInputElement>('#preview-zoom');
  zoom.value = String(state.previewWidth);
  zoom.addEventListener('input', () => {
    state.previewWidth = Number(zoom.value);
    refreshPreviews();
  });

  // The window itself must never navigate when a file lands outside a dropzone.
  for (const type of ['dragover', 'drop'] as const) {
    window.addEventListener(type, (event) => event.preventDefault());
  }
}
