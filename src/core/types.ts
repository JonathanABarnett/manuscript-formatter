/**
 * Domain model shared by the analyzer, the composer and the UI.
 * Everything here must be structured-clone safe: it crosses the Electron IPC
 * boundary, so no DOM nodes and no class instances.
 */

/**
 * A .docx as bytes plus a display name. The engine never touches the
 * filesystem, so the same code path serves the desktop app and the browser.
 */
export interface DocxInput {
  data: Uint8Array;
  /** File name for display and for naming the output. No path required. */
  name: string;
}

/** Semantic role assigned to a manuscript block. */
export type BlockRole =
  | 'frontMatterTitle'
  | 'frontMatter'
  | 'copyright'
  | 'partTitle'
  | 'chapterTitle'
  | 'chapterSubtitle'
  | 'subheading'
  | 'bodyFirst'
  | 'body'
  | 'blockQuote'
  | 'listItem'
  | 'sceneBreak'
  | 'table'
  | 'pageBreak'
  | 'empty';

/** Roles that are mapped onto a paragraph style taken from the reference. */
export type StyleRole = Exclude<BlockRole, 'table' | 'pageBreak' | 'empty'>;

export const STYLE_ROLES: StyleRole[] = [
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
];

/** What each kind of paragraph is called on screen, in an author's words. */
export const ROLE_LABELS: Record<BlockRole, string> = {
  frontMatterTitle: 'Title page',
  frontMatter: 'Other opening-page text',
  copyright: 'Copyright page',
  partTitle: 'Part title',
  chapterTitle: 'Chapter title',
  chapterSubtitle: 'Chapter subtitle',
  subheading: 'Heading inside a chapter',
  bodyFirst: 'First paragraph',
  body: 'Normal paragraph',
  blockQuote: 'Indented quotation',
  listItem: 'List item',
  sceneBreak: 'Scene break',
  table: 'Table',
  pageBreak: 'Page break',
  empty: 'Blank line',
};

/** One plain sentence per role, shown on hover and under the pickers. */
export const ROLE_HINTS: Record<BlockRole, string> = {
  frontMatterTitle: 'Your book title, on the opening page.',
  frontMatter: 'Your author name, dedication, epigraph, or other text before the first chapter.',
  copyright: 'The © line, ISBN, edition and rights wording. Usually set smaller than the story.',
  partTitle: 'A divider above chapter level, such as "Part One".',
  chapterTitle: 'The line that opens each chapter.',
  chapterSubtitle: 'A second line under a chapter title, if your chapters have one.',
  subheading: 'A smaller heading used partway through a chapter.',
  bodyFirst:
    'The paragraph that starts a chapter or follows a scene break. Printed books normally leave this one without an indent.',
  body: 'Ordinary paragraphs — the bulk of your book.',
  blockQuote: 'A longer quotation set in from both margins, not ordinary dialogue.',
  listItem: 'A bulleted or numbered item.',
  sceneBreak: 'The small mark between scenes, like * * * or #.',
  table: 'Copied across exactly as it is.',
  pageBreak: 'Starts a new page here.',
  empty: 'An empty line with nothing on it.',
};

export interface PageMargins {
  top: number;
  right: number;
  bottom: number;
  left: number;
  header: number;
  footer: number;
  gutter: number;
}

/** Page geometry, in twips (1/1440 inch), taken from the reference sectPr. */
export interface PageSetup {
  widthTwips: number;
  heightTwips: number;
  orientation: 'portrait' | 'landscape';
  margins: PageMargins;
  mirrorMargins: boolean;
  differentFirstPage: boolean;
  differentOddEven: boolean;
  sectionBreakType: string | null;
  columns: number;
}

/** A paragraph style from the reference, with the properties we surface. */
export interface StyleInfo {
  id: string;
  name: string;
  type: string;
  basedOn: string | null;
  next: string | null;
  isDefault: boolean;
  outlineLevel: number | null;
  /** How many paragraphs in the reference body use this style. */
  usageCount: number;
  fontName: string | null;
  fontSizePt: number | null;
  alignment: string | null;
  firstLineIndentTwips: number | null;
  hangingIndentTwips: number | null;
  leftIndentTwips: number | null;
  spaceBeforeTwips: number | null;
  spaceAfterTwips: number | null;
  lineSpacing: number | null;
  lineRule: string | null;
  pageBreakBefore: boolean;
  keepNext: boolean;
  bold: boolean;
  italic: boolean;
  allCaps: boolean;
  smallCaps: boolean;
}

/** Everything learned from the reference document. */
export interface ReferenceProfile {
  fileName: string;
  pageSetup: PageSetup;
  pageSizeLabel: string;
  sectionCount: number;
  styles: StyleInfo[];
  /** Role -> reference style id, as detected. Null when nothing suitable. */
  roleStyles: Record<StyleRole, string | null>;
  /** Evidence strings explaining each role choice, for the review screen. */
  roleEvidence: Partial<Record<StyleRole, string>>;
  chapterStartsOnNewPage: boolean;
  chapterStartsOnOddPage: boolean;
  /**
   * Blank paragraphs the reference sets in the chapter-title style either side
   * of the title itself. Book templates use these to sink a chapter opening
   * partway down the page, so the output has to reproduce them.
   */
  chapterTitleBlanksBefore: number;
  chapterTitleBlanksAfter: number;
  /**
   * Where a page's text sits vertically, per kind of page. Book templates
   * park a copyright notice at the foot of its page and centre a title page,
   * and they do it with section properties rather than blank lines.
   */
  roleVAlign: Partial<Record<StyleRole, string>>;
  /**
   * Index into the reference's sections of the one the body matter uses. It is
   * not always the last: templates put each chapter in its own section, and it
   * is the first body section that restarts page numbering at 1.
   */
  bodySectionIndex: number;
  usesFirstParagraphNoIndent: boolean;
  bodyFirstLineIndentTwips: number | null;
  hasHeaders: boolean;
  hasFooters: boolean;
  hasPageNumbers: boolean;
  /** Visible words inherited from referenced header/footer parts, excluding bare page numbers. */
  headerFooterText: string[];
  hasFootnotes: boolean;
  defaultParagraphStyleId: string | null;
  bodyFontName: string | null;
  bodyFontSizePt: number | null;
  bodyLineSpacing: number | null;
  /**
   * How `bodyLineSpacing` is meant: `auto` counts in 240ths of a line, while
   * `exact` and `atLeast` are an absolute height in twips. Reading it as the
   * wrong one throws the page estimate out by a third.
   */
  bodyLineRule: string | null;
  warnings: string[];
}

/** One block of the manuscript, as classified. */
export interface ManuscriptBlock {
  index: number;
  kind: 'paragraph' | 'table';
  /** Classifier's decision. */
  autoRole: BlockRole;
  /** Effective role (user override wins). Equal to autoRole until overridden. */
  role: BlockRole;
  text: string;
  /** Trimmed preview, capped for display. */
  preview: string;
  styleId: string | null;
  styleName: string | null;
  outlineLevel: number | null;
  alignment: string | null;
  hasPageBreakBefore: boolean;
  isEmpty: boolean;
  wordCount: number;
  charCount: number;
  hasNumbering: boolean;
  leftIndentTwips: number | null;
  firstLineIndentTwips: number | null;
  allBold: boolean;
  allItalic: boolean;
  hasImage: boolean;
  hasFootnote: boolean;
  hasHyperlink: boolean;
  /** The paragraph names a book division outright ("Chapter 4", "Prologue"). */
  structuralMarker: boolean;
  /** Widest picture in this block, in twips. Null when it holds none. */
  imageWidthTwips: number | null;
  /** Declared width of a table block, in twips. Null when it does not say. */
  tableWidthTwips: number | null;
  /** 0..1 — how sure the classifier is. Low values are surfaced for review. */
  confidence: number;
  reasons: string[];
}

export interface ManuscriptAnalysis {
  fileName: string;
  blocks: ManuscriptBlock[];
  wordCount: number;
  paragraphCount: number;
  chapterCount: number;
  partCount: number;
  sceneBreakCount: number;
  tableCount: number;
  imageCount: number;
  footnoteCount: number;
  /** Index of the first block classified as body-matter content. */
  bodyStartIndex: number;
  /** The manuscript already carries a contents list of its own. */
  hasContentsPage: boolean;
  /** Title, author and so on read out of the manuscript's own opening pages. */
  detectedDetails: Partial<BookDetails>;
  warnings: string[];
}

export type ChapterStartMode = 'newPage' | 'oddPage' | 'continuous';

export interface FormatOptions {
  /** Role -> style id. Merged over the detected reference mapping. */
  roleStyles: Partial<Record<StyleRole, string | null>>;
  /** Per-block role overrides, keyed by block index. */
  roleOverrides: Record<number, BlockRole>;
  chapterStart: ChapterStartMode;
  /** Apply the no-indent style to the first paragraph after a break. */
  firstParagraphNoIndent: boolean;
  /** Drop blank paragraphs; book spacing comes from the styles instead. */
  removeEmptyParagraphs: boolean;
  /** Strip leading tabs/spaces used as manual first-line indents. */
  removeManualIndents: boolean;
  /** Collapse runs of two or more spaces to one. */
  collapseMultipleSpaces: boolean;
  /** Convert straight quotes and hyphens to typographic equivalents. */
  smartTypography: boolean;
  /** Replace scene-break text with this, e.g. "* * *". Null keeps the original. */
  sceneBreakText: string | null;
  /** Keep bold/italic/etc. from the manuscript runs. */
  keepEmphasis: boolean;
  /** Include the manuscript's front matter in the output. */
  includeFrontMatter: boolean;
  /** What the small line at the top of each page says. */
  runningHeads: RunningHeads;
  /**
   * Blank lines above and below a chapter title, which is how a template sinks
   * a chapter opening down the page. Null follows whatever the template does.
   */
  chapterSpaceBefore: number | null;
  chapterSpaceAfter: number | null;
  /** Details used to build the title and copyright pages. */
  bookDetails: BookDetails;
  /** Which extra sections to add around the book. */
  extraSections: ExtraSections;
  /**
   * Drop the front matter already in the manuscript and use the generated
   * pages instead, rather than ending up with two title pages.
   */
  replaceFrontMatter: boolean;
}

/** What the author tells us about the book, for its opening pages. */
export interface BookDetails {
  title: string;
  subtitle: string;
  author: string;
  copyrightYear: string;
  publisher: string;
  isbn: string;
  dedication: string;
  acknowledgments: string;
  aboutTheAuthor: string;
  bibliography: string;
}

export interface ExtraSections {
  titlePage: boolean;
  copyrightPage: boolean;
  dedication: boolean;
  contents: boolean;
  acknowledgments: boolean;
  aboutTheAuthor: boolean;
  bibliography: boolean;
}

/**
 * The running heads. `auto` puts the book's title and author in place of a
 * template's placeholder wording; `custom` sets them outright; `leave` does
 * not touch what the design already says.
 */
export interface RunningHeads {
  mode: 'auto' | 'custom' | 'leave';
  /** Left-hand pages. Printed books usually carry the author here. */
  verso: string;
  /** Right-hand pages. Usually the book's title. */
  recto: string;
}

export const DEFAULT_RUNNING_HEADS: RunningHeads = { mode: 'auto', verso: '', recto: '' };

export const EMPTY_BOOK_DETAILS: BookDetails = {
  title: '',
  subtitle: '',
  author: '',
  copyrightYear: '',
  publisher: '',
  isbn: '',
  dedication: '',
  acknowledgments: '',
  aboutTheAuthor: '',
  bibliography: '',
};

export const NO_EXTRA_SECTIONS: ExtraSections = {
  titlePage: false,
  copyrightPage: false,
  dedication: false,
  contents: false,
  acknowledgments: false,
  aboutTheAuthor: false,
  bibliography: false,
};

export const DEFAULT_FORMAT_OPTIONS: FormatOptions = {
  roleStyles: {},
  roleOverrides: {},
  chapterStart: 'newPage',
  firstParagraphNoIndent: true,
  removeEmptyParagraphs: true,
  removeManualIndents: true,
  collapseMultipleSpaces: false,
  smartTypography: false,
  sceneBreakText: null,
  keepEmphasis: true,
  includeFrontMatter: true,
  chapterSpaceBefore: null,
  chapterSpaceAfter: null,
  bookDetails: EMPTY_BOOK_DETAILS,
  runningHeads: DEFAULT_RUNNING_HEADS,
  extraSections: NO_EXTRA_SECTIONS,
  replaceFrontMatter: false,
};

export interface FormatStats {
  paragraphsWritten: number;
  chapters: number;
  parts: number;
  sceneBreaks: number;
  tables: number;
  imagesCopied: number;
  footnotesCopied: number;
  blanksRemoved: number;
  wordCount: number;
  /** Header or footer parts whose placeholder wording was replaced. */
  runningHeadsUpdated: number;
}

export interface FormatResult {
  outputPath: string;
  stats: FormatStats;
  warnings: string[];
}
