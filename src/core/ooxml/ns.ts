/** OOXML namespace URIs used throughout the app. */
export const NS = {
  /** WordprocessingML main */
  w: 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
  /** Relationship references inside part XML (r:id, r:embed) */
  r: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
  /** The .rels part format itself */
  pkgRel: 'http://schemas.openxmlformats.org/package/2006/relationships',
  /** [Content_Types].xml */
  contentTypes: 'http://schemas.openxmlformats.org/package/2006/content-types',
  /** DrawingML main */
  a: 'http://schemas.openxmlformats.org/drawingml/2006/main',
  /** DrawingML wordprocessing drawing */
  wp: 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
  /** VML (legacy shapes, still emitted by Word for some objects) */
  v: 'urn:schemas-microsoft-com:vml',
  xml: 'http://www.w3.org/XML/1998/namespace',
} as const;

/** Relationship type URIs. */
export const RELTYPE = {
  officeDocument:
    'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument',
  styles: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles',
  numbering: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering',
  settings: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings',
  theme: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme',
  fontTable: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable',
  header: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/header',
  footer: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer',
  footnotes: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes',
  endnotes: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes',
  image: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
  hyperlink: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink',
} as const;

/** Content types for parts we may add to the output package. */
export const CONTENT_TYPE = {
  footnotes:
    'application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml',
  endnotes:
    'application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml',
} as const;

/** Twips (1/1440 inch) helpers — OOXML's unit for page and margin dimensions. */
export const twipsToInches = (twips: number): number => twips / 1440;
export const inchesToTwips = (inches: number): number => Math.round(inches * 1440);
/** Half-points are the unit for font size (w:sz). */
export const halfPointsToPoints = (hp: number): number => hp / 2;
