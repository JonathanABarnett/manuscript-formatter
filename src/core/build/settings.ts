import { NS, RELTYPE } from '../ooxml/ns.js';
import type { DocxPackage, Relationships } from '../ooxml/package.js';
import { attr, child, children, wEl } from '../ooxml/xml.js';

/**
 * Document-level settings that make a Word file print well. None of them
 * change what the book says; they change how Word breaks lines and what it
 * keeps when the author saves.
 */

/**
 * The order `w:settings` children must appear in. Word is fussier about
 * this part than most, so a new element is slotted in ahead of the first
 * existing one that the schema places after it.
 */
const SETTINGS_ORDER = [
  'writeProtection',
  'view',
  'zoom',
  'removePersonalInformation',
  'removeDateAndTime',
  'doNotDisplayPageBoundaries',
  'displayBackgroundShape',
  'printPostScriptOverText',
  'printFractionalCharacterWidth',
  'printFormsData',
  'embedTrueTypeFonts',
  'embedSystemFonts',
  'saveSubsetFonts',
  'saveFormsData',
  'mirrorMargins',
  'alignBordersAndEdges',
  'bordersDoNotSurroundHeader',
  'bordersDoNotSurroundFooter',
  'gutterAtTop',
  'hideSpellingErrors',
  'hideGrammaticalErrors',
  'activeWritingStyle',
  'proofState',
  'formsDesign',
  'attachedTemplate',
  'linkStyles',
  'stylePaneFormatFilter',
  'stylePaneSortMethod',
  'documentType',
  'mailMerge',
  'revisionView',
  'trackRevisions',
  'doNotTrackMoves',
  'doNotTrackFormatting',
  'documentProtection',
  'autoFormatOverride',
  'styleLockTheme',
  'styleLockQFSet',
  'defaultTabStop',
  'autoHyphenation',
  'consecutiveHyphenLimit',
  'hyphenationZone',
  'doNotHyphenateCaps',
  'showEnvelope',
  'summaryLength',
  'clickAndTypeStyle',
  'defaultTableStyle',
  'evenAndOddHeaders',
  'bookFoldRevPrinting',
  'bookFoldPrinting',
  'bookFoldPrintingSheets',
  'drawingGridHorizontalSpacing',
  'drawingGridVerticalSpacing',
  'displayHorizontalDrawingGridEvery',
  'displayVerticalDrawingGridEvery',
  'doNotUseMarginsForDrawingGridOrigin',
  'drawingGridHorizontalOrigin',
  'drawingGridVerticalOrigin',
  'doNotShadeFormData',
  'noPunctuationKerning',
  'characterSpacingControl',
  'printTwoOnOne',
  'strictFirstAndLastChars',
  'noLineBreaksAfter',
  'noLineBreaksBefore',
  'savePreviewPicture',
  'doNotValidateAgainstSchema',
  'saveInvalidXml',
  'ignoreMixedContent',
  'alwaysShowPlaceholderText',
  'doNotDemarcateInvalidXml',
  'saveXmlDataOnly',
  'useXSLTWhenSaving',
  'saveThroughXslt',
  'showXMLTags',
  'alwaysMergeEmptyNamespace',
  'updateFields',
  'hdrShapeDefaults',
  'footnotePr',
  'endnotePr',
  'compat',
  'docVars',
  'rsids',
  'mathPr',
  'attachedSchema',
  'themeFontLang',
  'clrSchemeMapping',
  'doNotIncludeSubdocsInStats',
  'doNotAutoCompressPictures',
  'forceUpgrade',
  'captions',
  'readModeInkLockDown',
  'smartTagType',
  'schemaLibrary',
  'shapeDefaults',
  'doNotEmbedSmartTags',
  'decimalSymbol',
  'listSeparator',
];
const SETTINGS_RANK = new Map(SETTINGS_ORDER.map((name, i) => [name, i]));

export interface PrintSettings {
  /** Ask Word to hyphenate at line ends. */
  hyphenate: boolean;
}

/**
 * Apply the settings that help a `.docx` print well:
 *
 * - `embedTrueTypeFonts` + `saveSubsetFonts`, so that when the author saves
 *   the finished file in Word its fonts travel inside it, which is what KDP
 *   asks for from a Word upload.
 * - `autoHyphenation`, on or off as chosen, with capitals left unbroken and
 *   no more than two hyphenated line ends in a row when it is on.
 *
 * The settings part is created if the design has none.
 */
export async function applyPrintSettings(
  out: DocxPackage,
  rels: Relationships,
  settings: PrintSettings,
): Promise<void> {
  const { doc, root, path } = await openSettings(out, rels);
  ensureToggle(doc, root, 'embedTrueTypeFonts', true);
  ensureToggle(doc, root, 'saveSubsetFonts', true);

  if (settings.hyphenate) {
    ensureToggle(doc, root, 'autoHyphenation', true);
    if (!child(root, 'consecutiveHyphenLimit')) {
      insertSetting(root, wEl(doc, 'consecutiveHyphenLimit', { val: 2 }));
    }
    ensureToggle(doc, root, 'doNotHyphenateCaps', true);
  } else {
    // A design that hyphenated is asked not to; one that never did stays quiet.
    const existing = child(root, 'autoHyphenation');
    if (existing) root.removeChild(existing);
  }
  out.writeXml(path, doc);
}

/**
 * Set `w:updateFields` so Word offers to fill in the contents table on open.
 */
export async function requestFieldUpdate(out: DocxPackage, rels: Relationships): Promise<void> {
  const { doc, root, path } = await openSettings(out, rels);
  if (child(root, 'updateFields')) return;
  insertSetting(root, wEl(doc, 'updateFields', { val: 'true' }));
  out.writeXml(path, doc);
}

/**
 * Tag the output's default text with the manuscript's language, so Word
 * hyphenates and spell-checks in the language the book is actually written
 * in rather than whichever the design happened to be saved with. Only the
 * Latin-script slot is set; East Asian and complex-script tags are left as
 * the design has them.
 */
export async function applyLanguage(
  out: DocxPackage,
  rels: Relationships,
  language: string,
): Promise<void> {
  const target = rels.firstTargetOfType(RELTYPE.styles);
  if (!target) return;
  const path = out.resolveTarget(out.documentPath, target);
  const doc = await out.readXml(path);
  const root = doc?.documentElement;
  if (!doc || !root) return;

  let changed = false;
  const setLang = (rPr: Element | null): void => {
    if (!rPr) return;
    const lang = child(rPr, 'lang');
    if (lang) {
      if (attr(lang, 'val') === language) return;
      lang.setAttributeNS(NS.w, 'w:val', language);
    } else {
      // `w:lang` comes late in `w:rPr`, after the effects and before
      // `w:eastAsianLayout`/`w:specVanish`/`w:oMath`; appending is safe for
      // the elements a docDefaults or style rPr carries.
      rPr.appendChild(wEl(doc, 'lang', { val: language }));
    }
    changed = true;
  };

  const defaults = child(child(child(root, 'docDefaults'), 'rPrDefault'), 'rPr');
  if (defaults) setLang(defaults);
  else {
    const docDefaults = child(root, 'docDefaults');
    if (docDefaults) {
      let rPrDefault = child(docDefaults, 'rPrDefault');
      if (!rPrDefault) {
        rPrDefault = wEl(doc, 'rPrDefault');
        docDefaults.insertBefore(rPrDefault, docDefaults.firstChild);
      }
      const rPr = wEl(doc, 'rPr');
      rPrDefault.appendChild(rPr);
      setLang(rPr);
    }
  }

  // A style that names its own language overrides the defaults, so the
  // paragraph styles that carry the book's text must agree too.
  for (const style of children(root, 'style')) {
    if (attr(style, 'type') !== 'paragraph') continue;
    const rPr = child(style, 'rPr');
    if (rPr && child(rPr, 'lang')) setLang(rPr);
  }

  if (changed) out.writeXml(path, doc);
}

async function openSettings(
  out: DocxPackage,
  rels: Relationships,
): Promise<{ doc: Document; root: Element; path: string }> {
  let target = rels.firstTargetOfType(RELTYPE.settings);
  if (!target) {
    target = 'settings.xml';
    rels.add(RELTYPE.settings, target);
    out.writeText(
      `${out.documentDir}settings.xml`,
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<w:settings xmlns:w="${NS.w}"/>`,
    );
    await out.ensureContentType({
      partName: `${out.documentDir}settings.xml`,
      partType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml',
    });
  }
  const path = out.resolveTarget(out.documentPath, target);
  const doc = await out.readXml(path);
  const root = doc?.documentElement;
  if (!doc || !root) throw new Error('The document settings could not be read.');
  return { doc, root, path };
}

/** Make a boolean setting present and on (or off), keeping the schema order. */
function ensureToggle(doc: Document, root: Element, name: string, on: boolean): void {
  const existing = child(root, name);
  if (existing) {
    existing.removeAttributeNS(NS.w, 'val');
    existing.removeAttribute('w:val');
    if (!on) existing.setAttributeNS(NS.w, 'w:val', 'false');
    return;
  }
  insertSetting(root, on ? wEl(doc, name) : wEl(doc, name, { val: 'false' }));
}

/** Insert ahead of the first existing child the schema orders after it. */
function insertSetting(root: Element, el: Element): void {
  const rank = SETTINGS_RANK.get(el.localName ?? '') ?? Number.POSITIVE_INFINITY;
  for (let n = root.firstChild; n; n = n.nextSibling) {
    if (n.nodeType !== 1) continue;
    const other = SETTINGS_RANK.get((n as Element).localName ?? '');
    if (other !== undefined && other > rank) {
      root.insertBefore(el, n);
      return;
    }
  }
  root.appendChild(el);
}
