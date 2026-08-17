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
import { BOOK_LOOKS, TRIM_SIZES } from '../core/templates/design.js';
import { renderPreviews } from './previews.js';
import { renderDetailsForm } from './details.js';
import { preflight, type CheckLevel, type PreflightReport } from '../core/preflight.js';
import { hasSeenTour, startTour } from './tour.js';

interface State {
  /** Quick Start builds a design; "own" uses a file the author supplies. */
  mode: 'quick' | 'own';
  trimId: string;
  lookId: string;
  referencePath: string | null;
  manuscriptPath: string | null;
  analysis: AnalysisResult | null;
  options: FormatOptions;
  outputPath: string | null;
  showAllBlocks: boolean;
  preflight: PreflightReport | null;
  /** Set when the author stops a run; the finished file is then discarded. */
  cancelled: boolean;
  /** Rendered width of a preview page, in CSS pixels. */
  previewWidth: number;
  busy: boolean;
  result: FormatResult | null;
}

const state: State = {
  mode: 'quick',
  trimId: TRIM_SIZES.find((t) => t.recommended)?.id ?? TRIM_SIZES[0].id,
  lookId: BOOK_LOOKS[0].id,
  referencePath: null,
  manuscriptPath: null,
  analysis: null,
  options: { ...DEFAULT_FORMAT_OPTIONS },
  outputPath: null,
  showAllBlocks: false,
  preflight: null,
  cancelled: false,
  previewWidth: 380,
  busy: false,
  result: null,
};

/** Prevent a slower, older file analysis from replacing a newer selection. */
let analysisVersion = 0;

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

// --- step 1: quick start ----------------------------------------------------

/** Quick Start builds a design; "use my own" takes one the author supplies. */
function setMode(mode: 'quick' | 'own'): void {
  state.mode = mode;
  const quick = mode === 'quick';
  $('#quick-start').hidden = !quick;
  $('#own-design').hidden = quick;
  $('#mode-quick').setAttribute('aria-selected', String(quick));
  $('#mode-own').setAttribute('aria-selected', String(!quick));
  $('#mode-quick').classList.toggle('active', quick);
  $('#mode-own').classList.toggle('active', !quick);
  analysisVersion++;

  // The two modes source the design differently, so a half-made choice from
  // one must not leak into the other.
  state.referencePath = null;
  state.analysis = null;
  state.result = null;
  $('#step-review').hidden = true;
  $('#step-output').hidden = true;
  $('#analyze-status').textContent = '';

  if (quick) void applyQuickTemplate();
}

function renderQuickChoices(): void {
  const trims = $('#trim-choices');
  replace(
    trims,
    ...TRIM_SIZES.map((trim) =>
      choiceCard({
        name: 'trim',
        value: trim.id,
        title: trim.label,
        note: trim.note,
        badge: trim.recommended ? 'Recommended' : undefined,
        checked: state.trimId === trim.id,
        onPick: () => {
          state.trimId = trim.id;
          void applyQuickTemplate();
        },
      }),
    ),
  );

  const looks = $('#look-choices');
  replace(
    looks,
    ...BOOK_LOOKS.map((look) =>
      choiceCard({
        name: 'look',
        value: look.id,
        title: look.label,
        note: look.note,
        checked: state.lookId === look.id,
        onPick: () => {
          state.lookId = look.id;
          void applyQuickTemplate();
        },
      }),
    ),
  );
}

interface ChoiceOptions {
  name: string;
  value: string;
  title: string;
  note: string;
  badge?: string;
  checked: boolean;
  onPick: () => void;
}

function choiceCard(o: ChoiceOptions): HTMLElement {
  const input = el('input', { type: 'radio', name: o.name, value: o.value, checked: o.checked });
  input.addEventListener('change', () => {
    if (input.checked) o.onPick();
  });
  return el(
    'label',
    { class: `choice${o.checked ? ' chosen' : ''}` },
    input,
    el(
      'span',
      { class: 'choice-body' },
      el(
        'span',
        { class: 'choice-title' },
        o.title,
        o.badge ? el('span', { class: 'choice-badge' }, o.badge) : null,
      ),
      el('span', { class: 'choice-note' }, o.note),
    ),
  );
}

/**
 * Build the chosen design and use it as the reference. Runs whenever a size or
 * look changes, so the preview always reflects the current pick.
 */
async function applyQuickTemplate(): Promise<void> {
  if (state.mode !== 'quick') return;
  renderQuickChoices();
  try {
    state.referencePath = await window.formatter.useBuiltIn({
      kind: 'template',
      trimId: state.trimId,
      lookId: state.lookId,
    });
  } catch (err) {
    showError(err instanceof Error ? err.message : String(err));
    return;
  }
  if (state.manuscriptPath) await runAnalysis();
}

/** Load the bundled sample book so the whole flow can be tried risk-free. */
async function useSampleBook(): Promise<void> {
  try {
    state.manuscriptPath = await window.formatter.useBuiltIn({ kind: 'sample' });
  } catch (err) {
    showError(err instanceof Error ? err.message : String(err));
    return;
  }
  markChosen('manuscript', state.manuscriptPath);
  state.outputPath = null;
  if (!state.referencePath) await applyQuickTemplate();
  else await runAnalysis();
}

function startOver(): void {
  state.referencePath = null;
  state.manuscriptPath = null;
  state.analysis = null;
  state.result = null;
  state.outputPath = null;
  state.options = { ...DEFAULT_FORMAT_OPTIONS };
  for (const zone of document.querySelectorAll<HTMLElement>('.dropzone')) {
    zone.classList.remove('chosen');
    const label = zone.querySelector<HTMLElement>('.dropzone-label');
    const file = zone.querySelector<HTMLElement>('.dropzone-file');
    if (label) label.textContent = 'Drag a Word file here, or click to browse';
    if (file) {
      file.textContent = '';
      file.hidden = true;
    }
  }
  $('#step-review').hidden = true;
  $('#step-output').hidden = true;
  $('#result').hidden = true;
  $('#analyze-status').textContent = '';
  $('#start-over').hidden = true;
  setMode('quick');
  $('#step-files').scrollIntoView({ behavior: 'smooth', block: 'start' });
}

// --- step 1: choosing files -------------------------------------------------

function setupPicker(kind: 'reference' | 'manuscript', zone: HTMLElement | null): void {
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

/** Show a chosen file on every dropzone that stands for that slot. */
function markChosen(kind: 'reference' | 'manuscript', path: string): void {
  for (const zone of document.querySelectorAll<HTMLElement>(
    `.dropzone[data-kind="${kind}"], #picker-${kind} .dropzone`,
  )) {
    zone.classList.add('chosen');
    const label = zone.querySelector<HTMLElement>('.dropzone-label');
    const file = zone.querySelector<HTMLElement>('.dropzone-file');
    if (label) label.textContent = 'Selected';
    if (file) {
      file.textContent = baseName(path);
      file.hidden = false;
      file.title = path;
    }
  }
  $('#start-over').hidden = false;
}

async function setPath(kind: 'reference' | 'manuscript', path: string): Promise<void> {
  analysisVersion++;
  if (kind === 'reference') state.referencePath = path;
  else state.manuscriptPath = path;

  // Any previous analysis and output belong to the old pair of files.
  state.analysis = null;
  state.result = null;
  if (kind === 'manuscript') state.outputPath = null;
  $('#step-review').hidden = true;
  $('#step-output').hidden = true;
  $('#result').hidden = true;

  markChosen(kind, path);

  if (state.referencePath && state.referencePath === state.manuscriptPath) {
    $('#analyze-status').textContent = 'Choose a different file for the design and the manuscript.';
    showError('The design file and manuscript cannot be the same file.');
    return;
  }

  if (state.referencePath && state.manuscriptPath) await runAnalysis();
}

async function runAnalysis(): Promise<void> {
  const version = analysisVersion;
  const status = $('#analyze-status');
  status.className = 'status spinner';
  status.textContent = 'Reading your two files';
  showError(null);

  const outcome = await window.formatter.analyze({
    referencePath: state.referencePath!,
    manuscriptPath: state.manuscriptPath!,
  });

  if (version !== analysisVersion) return;

  status.className = 'status';
  if (!outcome.ok) {
    status.textContent = '';
    showError(outcome.error);
    return;
  }

  state.analysis = outcome.value;
  state.options = { ...outcome.value.suggestedOptions, roleStyles: {}, roleOverrides: {} };
  const found = outcome.value.analysis;
  const structureFound = [
    `${found.chapterCount} chapter${found.chapterCount === 1 ? '' : 's'}`,
    found.partCount > 0 ? `${found.partCount} part${found.partCount === 1 ? '' : 's'}` : null,
  ]
    .filter(Boolean)
    .join(' and ');
  status.textContent = `Read ${found.wordCount.toLocaleString()} words and found ${structureFound}. Check the notes below.`;

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
  renderReviewWarnings(state.analysis.profile, state.analysis.analysis);
  renderTemplateFacts(state.analysis.profile);
  renderOptions();
  renderDetails();
  refreshPreviews();
  renderPreflight();
  renderStyleMap(state.analysis.profile);
  renderStructure();
}

/** Important findings belong before the Format button, not after it. */
function renderReviewWarnings(
  profile: ReferenceProfile,
  analysis: AnalysisResult['analysis'],
): void {
  const panel = $('#review-warnings');
  const inherited = profile.headerFooterText;
  const needsReview = analysis.blocks.filter(
    (block) => block.confidence < 0.6 && block.role !== 'body' && block.role !== 'empty',
  ).length;
  const warnings = [...new Set([...profile.warnings, ...analysis.warnings])].filter(
    (warning) => inherited.length === 0 || !warning.startsWith('Words in the design file'),
  );

  if (warnings.length === 0 && inherited.length === 0) {
    panel.hidden = true;
    panel.replaceChildren();
    return;
  }

  const items: HTMLElement[] = warnings.map((warning) => el('li', {}, warning));
  if (inherited.length > 0) {
    items.unshift(
      el(
        'li',
        {},
        el('strong', {}, 'Header or footer words being reused: '),
        inherited.map((text) => `“${text}”`).join(', '),
      ),
    );
  }

  const jumpToReview =
    needsReview > 0
      ? el(
          'button',
          { class: 'secondary review-jump', type: 'button' },
          'Go to items marked “check this”',
        )
      : null;
  jumpToReview?.addEventListener('click', () => {
    const first = document.querySelector<HTMLElement>('.block-row.needs-review');
    first?.scrollIntoView({ behavior: 'smooth', block: 'center' });
    first?.querySelector<HTMLSelectElement>('select')?.focus({ preventScroll: true });
  });

  replace(
    panel,
    el('h3', {}, 'Check these before you continue'),
    el(
      'p',
      { class: 'panel-hint' },
      'The app can finish the formatting, but these items deserve a quick look first.',
    ),
    el('ul', {}, ...items),
    jumpToReview,
  );
  panel.hidden = false;
}

function renderDetails(): void {
  renderDetailsForm($('#details-form'), {
    options: state.options,
    onChange: (structural) => {
      // Only a switched section changes the form's shape; typing must not
      // redraw the field being typed into.
      if (structural) renderDetails();
      renderPreflight();
    },
  });
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
      `${inches(m.top)} top, ${inches(m.bottom)} bottom, ${inches(m.left)} left, ${inches(m.right)} right`,
    ],
    [
      'Extra binding space',
      'Extra space on the inside edge so your words are not swallowed by the spine.',
      m.gutter > 0 ? inches(m.gutter) : 'none — built into the side margins',
    ],
    [
      'Mirrored left and right pages',
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
      'Formatting styles',
      'Ready-made looks stored in the design file for titles, paragraphs, quotations, and more.',
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
      'Continue on the same page',
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
        `Use the design file (${fromDesign === 0 ? 'none' : `${fromDesign} blank line${fromDesign === 1 ? '' : 's'}`})`,
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
      'Choosing a right-hand page can add a blank page between chapters.',
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
      el('span', {}, 'Mark used between scenes'),
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
      'Remove empty paragraphs',
      'Recommended for most novels. Turn this off for poetry or layouts that use blank lines on purpose.',
    ),
    checkbox(
      'removeManualIndents',
      'Remove manually typed indents',
      'Removes tabs and spaces typed at the start of paragraphs so the design controls every indent.',
    ),
    checkbox(
      'keepEmphasis',
      'Keep italics, bold, and other emphasis',
      'Keeps emphasis such as italics, bold, underlining, and small capitals.',
    ),
    checkbox(
      'includeFrontMatter',
      'Include front matter',
      'Keeps material before the first chapter, such as the title, copyright, dedication, or epigraph.',
    ),
    checkbox(
      'smartTypography',
      'Fix straight quotes, dashes, and ellipses',
      'Changes straight quotes to curly quotes, -- to an em dash, and ... to an ellipsis. Leave this off if your manuscript is already edited.',
    ),
    checkbox(
      'collapseMultipleSpaces',
      'Reduce multiple spaces to one',
      'Changes every group of two or more spaces to a single space.',
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
  if (evidence) return `Reason: ${evidence}.`;
  return profile.roleStyles[role]
    ? 'Picked from your design file.'
      : 'No exact match was found, so the closest available formatting is used.';
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
    stat('chapters', counts.get('chapterTitle') ?? 0),
    (counts.get('partTitle') ?? 0) > 0 ? stat('parts', counts.get('partTitle') ?? 0) : null,
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
        } marked "check this". Turn on "Show every paragraph" to see the full manuscript.`
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
        'No chapters or headings were found. Turn on "Show every paragraph" above, then mark ' +
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
        title: block.reasons.length > 0 ? `Reason: ${block.reasons.join('; ')}` : undefined,
      },
      block.preview || '(empty line)',
      needsReview ? el('span', { class: 'flag' }, 'check this') : null,
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

/** The plain-English report, shown before making the book and again after. */
function renderPreflight(): void {
  const panel = $('#preflight');
  if (!state.analysis) {
    panel.hidden = true;
    return;
  }
  const report = preflight({
    profile: state.analysis.profile,
    analysis: state.analysis.analysis,
    options: state.options,
  });
  state.preflight = report;

  const HEADINGS: Record<CheckLevel, string> = {
    attention: 'Needs attention',
    check: 'Worth checking',
    ready: 'Ready',
  };
  const groups = (['attention', 'check', 'ready'] as CheckLevel[])
    .map((level) => [level, report.checks.filter((c) => c.level === level)] as const)
    .filter(([, list]) => list.length > 0);

  replace(
    panel,
    el(
      'div',
      { class: `preflight-head level-${report.level}` },
      el(
        'h3',
        {},
        report.level === 'attention'
          ? 'A few things need your attention'
          : report.level === 'check'
            ? 'Worth a quick look before you upload'
            : 'Everything looks ready',
      ),
      report.estimatedPages !== null
        ? el(
            'span',
            { class: 'preflight-pages' },
            `roughly ${report.estimatedPages} page${report.estimatedPages === 1 ? '' : 's'}`,
          )
        : null,
    ),
    ...groups.map(([level, list]) =>
      el(
        'div',
        { class: `preflight-group level-${level}` },
        el('h4', {}, HEADINGS[level], el('span', { class: 'preflight-count' }, String(list.length))),
        el(
          'ul',
          {},
          ...list.map((check) =>
            el(
              'li',
              {},
              el('span', { class: 'preflight-title' }, check.title),
              el('span', { class: 'preflight-detail' }, check.detail),
              check.examples && check.examples.length > 0
                ? el(
                    'ul',
                    { class: 'preflight-examples' },
                    ...check.examples.map((example) => el('li', {}, exampleLink(example))),
                  )
                : null,
            ),
          ),
        ),
      ),
    ),
    el(
      'p',
      { class: 'preflight-footnote' },
      'These are checks this app can make. Amazon’s own Print Previewer is the one that decides ' +
        'what gets printed — always look through it there before publishing.',
    ),
  );
  panel.hidden = false;
}

/**
 * A quoted line from the manuscript that jumps to it in the paragraph list,
 * so a finding can be looked at rather than merely read about.
 */
function exampleLink(example: { index: number; preview: string }): HTMLElement {
  const button = el(
    'button',
    { type: 'button', class: 'preflight-example' },
    `“${example.preview}”`,
  );
  button.title = 'Show me this line';
  button.addEventListener('click', () => revealBlock(example.index));
  return button;
}

/** Scroll the paragraph list to a block and mark it, expanding if hidden. */
function revealBlock(index: number): void {
  const showRow = (): HTMLElement | null =>
    document.querySelector<HTMLElement>(`.block-row[data-index="${index}"]`);

  if (!showRow()) {
    // Structural rows only are listed by default; the line may be prose.
    const toggle = $<HTMLInputElement>('#show-all-blocks');
    if (!toggle.checked) {
      toggle.checked = true;
      state.showAllBlocks = true;
      renderStructure();
    }
  }
  const row = showRow();
  if (!row) return;
  row.scrollIntoView({ behavior: 'smooth', block: 'center' });
  row.classList.add('flash');
  window.setTimeout(() => row.classList.remove('flash'), 1600);
}

/** Named stages, so waiting shows what is happening rather than a spinner. */
const FORMAT_STEPS = [
  'Reading your manuscript',
  'Finding chapters',
  'Applying the design',
  'Creating your file',
] as const;

function renderProgress(activeIndex: number): void {
  replace(
    $('#progress-steps'),
    ...FORMAT_STEPS.map((label, i) =>
      el(
        'li',
        { class: i < activeIndex ? 'done' : i === activeIndex ? 'active' : '' },
        el('span', { class: 'progress-mark' }, i < activeIndex ? '✓' : String(i + 1)),
        label,
      ),
    ),
  );
}

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
  state.cancelled = false;
  button.disabled = true;
  button.textContent = 'Making your book…';
  showError(null);
  $('#result').hidden = true;
  $('#progress').hidden = false;

  // The work itself is one call, so the steps are paced to it rather than
  // reported by it. They stop at the last stage until the file is really done.
  let step = 0;
  renderProgress(0);
  const ticker = window.setInterval(() => {
    if (step < FORMAT_STEPS.length - 1) renderProgress(++step);
  }, 320);

  const outcome = await window.formatter.format({
    referencePath: state.referencePath,
    manuscriptPath: state.manuscriptPath,
    outputPath: state.outputPath,
    options: state.options,
  });

  window.clearInterval(ticker);
  state.busy = false;
  button.disabled = false;
  button.textContent = 'Make my formatted book';
  $('#progress').hidden = true;

  if (state.cancelled) {
    // Be exact about what stopping actually achieved. In a browser nothing is
    // saved until Download is pressed, so cancelling really does discard it.
    // On the desktop the file has already been written by the time we return.
    $('#analyze-status').textContent =
      window.formatter.platform === 'web'
        ? 'Stopped. Nothing was saved to your computer.'
        : outcome.ok
          ? `Stopped, but the file had already been written to ${outcome.value.outputPath}.`
          : 'Stopped. Nothing was saved.';
    return;
  }

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
    s.parts > 0 ? `${s.parts} part${s.parts === 1 ? '' : 's'}` : null,
    s.sceneBreaks > 0 ? `${s.sceneBreaks} scene break${s.sceneBreaks === 1 ? '' : 's'}` : null,
    s.tables > 0 ? `${s.tables} table${s.tables === 1 ? '' : 's'}` : null,
    s.imagesCopied > 0 ? `${s.imagesCopied} picture${s.imagesCopied === 1 ? '' : 's'}` : null,
    s.footnotesCopied > 0
      ? `${s.footnotesCopied} footnote${s.footnotesCopied === 1 ? '' : 's'}`
      : null,
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
    el('h3', {}, 'Your formatted Word file is ready'),
    el('p', { class: 'path' }, result.outputPath),
    el('p', {}, line),
    el(
      'p',
      { class: 'result-next' },
      onWeb
        ? 'Download it, open it in Word, and review every page before uploading to KDP. Your two original files are untouched, so you can change any choice above and make it again.'
        : 'Open it in Word and review every page before uploading to KDP. Your two original files are untouched, so you can change any choice above and make it again.',
    ),
    el(
      'div',
      { class: 'final-checklist' },
      el('h4', {}, 'Final check in Word'),
      el(
        'ul',
        {},
        el('li', {}, 'Chapter and part openings begin on the pages you expect.'),
        el('li', {}, 'Page numbers run in order and any blank pages are intentional.'),
        el('li', {}, 'The correct book title and author appear in headers, footers, and opening pages.'),
        el('li', {}, 'Pictures and tables fit inside the page margins.'),
        el('li', {}, 'The copyright page and table of contents, if used, are up to date.'),
      ),
    ),
    result.warnings.length > 0
      ? el(
          'div',
          { class: 'warnings' },
          el('h4', {}, 'Please check'),
          el('ul', {}, ...result.warnings.map((w) => el('li', {}, w))),
        )
      : null,
    el(
      'div',
      { class: 'result-actions' },
      open,
      reveal,
      // Opens in the reader's own browser; the desktop shell hands http(s)
      // links to the system rather than navigating the app window.
      el(
        'a',
        {
          class: 'button-link',
          href: 'https://kdp.amazon.com/en_US/help/topic/G201834260',
          target: '_blank',
          rel: 'noopener noreferrer',
        },
        'What to do next on KDP',
      ),
    ),
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
      'Your files stay only in this tab’s memory while it is open. They are never uploaded or stored ' +
      'by the app. The finished Word file is saved only when you choose Download; nothing is sent to KDP.';
  }

  setupPicker('reference', document.querySelector('#picker-reference .dropzone'));
  setupPicker('manuscript', document.querySelector('#picker-manuscript .dropzone'));
  // Quick Start has its own manuscript dropzone feeding the same slot.
  setupPicker('manuscript', document.querySelector('#quick-manuscript .dropzone'));

  $('#mode-quick').addEventListener('click', () => setMode('quick'));
  $('#mode-own').addEventListener('click', () => setMode('own'));
  $('#try-sample').addEventListener('click', () => void useSampleBook());
  $('#start-over').addEventListener('click', startOver);
  $('#show-tour').addEventListener('click', () => startTour());
  $('#cancel-format').addEventListener('click', () => {
    // The engine runs to completion either way; cancelling discards the
    // result rather than pretending the work can be interrupted mid-file.
    state.cancelled = true;
    $('#progress').hidden = true;
  });
  // Also builds the default design and marks the tab, so the page opens ready.
  setMode('quick');

  // First visit gets the walkthrough; afterwards it waits behind the "?".
  if (!hasSeenTour()) window.setTimeout(startTour, 400);

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
