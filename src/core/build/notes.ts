import { CONTENT_TYPE, NS, RELTYPE } from '../ooxml/ns.js';
import { DocxPackage, Relationships } from '../ooxml/package.js';
import { attr, children, descendants, importNode, parseXml } from '../ooxml/xml.js';
import type { StyleSheet } from '../analyze/styles.js';
import { ResourceMigrator } from './resources.js';

export type NoteKind = 'footnote' | 'endnote';

interface KindSpec {
  root: string;
  item: string;
  relType: string;
  contentType: string;
  partName: string;
  textStyleName: RegExp;
  refStyleName: RegExp;
  separator: string;
  continuation: string;
}

const SPECS: Record<NoteKind, KindSpec> = {
  footnote: {
    root: 'footnotes',
    item: 'footnote',
    relType: RELTYPE.footnotes,
    contentType: CONTENT_TYPE.footnotes,
    partName: 'footnotes.xml',
    textStyleName: /^footnote\s*text$/i,
    refStyleName: /^footnote\s*reference$/i,
    separator: 'separator',
    continuation: 'continuationSeparator',
  },
  endnote: {
    root: 'endnotes',
    item: 'endnote',
    relType: RELTYPE.endnotes,
    contentType: CONTENT_TYPE.endnotes,
    partName: 'endnotes.xml',
    textStyleName: /^endnote\s*text$/i,
    refStyleName: /^endnote\s*reference$/i,
    separator: 'separator',
    continuation: 'continuationSeparator',
  },
};

function emptyPart(spec: KindSpec): string {
  const p = (marker: string): string =>
    `<w:${spec.item} w:type="${marker === 'separator' ? spec.separator : spec.continuation}" ` +
    `w:id="${marker === 'separator' ? -1 : 0}">` +
    '<w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>' +
    `<w:r><w:${marker}/></w:r></w:p></w:${spec.item}>`;
  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' +
    `<w:${spec.root} xmlns:w="${NS.w}">${p('separator')}${p('continuationSeparator')}</w:${spec.root}>`
  );
}

/**
 * Copies the manuscript's footnotes (or endnotes) into the output package,
 * renumbering them past whatever the reference already contains and remapping
 * the styles they use onto the reference's equivalents.
 */
export class NoteMerger {
  private readonly idMap = new Map<string, string>();
  private targetDoc: Document | null = null;
  private targetRoot: Element | null = null;
  private targetPartPath = '';
  private targetRels: Relationships | null = null;
  private migrator: ResourceMigrator | null = null;
  private sourceDoc: Document | null = null;
  private sourcePartPath = '';
  private nextId = 1;
  private dirty = false;
  private prepared = false;

  copied = 0;
  readonly warnings: string[] = [];

  constructor(
    private readonly kind: NoteKind,
    private readonly source: DocxPackage,
    private readonly sourceRels: Relationships,
    private readonly target: DocxPackage,
    private readonly targetDocRels: Relationships,
    private readonly referenceStyles: StyleSheet,
  ) {}

  private get spec(): KindSpec {
    return SPECS[this.kind];
  }

  private async prepare(): Promise<boolean> {
    if (this.prepared) return this.targetRoot !== null;
    this.prepared = true;

    const sourceTarget = this.sourceRels.firstTargetOfType(this.spec.relType);
    if (!sourceTarget) return false;
    this.sourcePartPath = this.source.resolveTarget(this.source.documentPath, sourceTarget);
    this.sourceDoc = await this.source.readXml(this.sourcePartPath);
    if (!this.sourceDoc) return false;

    const docDir = this.target.documentDir;
    let targetTarget = this.targetDocRels.firstTargetOfType(this.spec.relType);
    if (!targetTarget) {
      targetTarget = this.spec.partName;
      this.targetDocRels.add(this.spec.relType, targetTarget);
      this.target.writeText(`${docDir}${this.spec.partName}`, emptyPart(this.spec));
      await this.target.ensureContentType({
        partName: `${docDir}${this.spec.partName}`,
        partType: this.spec.contentType,
      });
      this.dirty = true;
    }
    this.targetPartPath = this.target.resolveTarget(this.target.documentPath, targetTarget);
    const doc = (await this.target.readXml(this.targetPartPath)) ?? parseXml(emptyPart(this.spec));
    this.targetDoc = doc;
    this.targetRoot = doc.documentElement;
    if (!this.targetRoot) return false;

    for (const el of children(this.targetRoot, this.spec.item)) {
      const id = Number(attr(el, 'id') ?? '0');
      if (Number.isFinite(id) && id >= this.nextId) this.nextId = id + 1;
    }

    this.targetRels = await this.target.relsFor(this.targetPartPath);
    this.migrator = new ResourceMigrator(
      this.source,
      this.sourcePartPath,
      await this.source.relsFor(this.sourcePartPath),
      this.target,
      this.targetPartPath,
      this.targetRels,
    );
    return true;
  }

  /** Map a manuscript note id onto one valid in the output. */
  async mapId(sourceId: string): Promise<string | null> {
    const cached = this.idMap.get(sourceId);
    if (cached) return cached;
    const numeric = Number(sourceId);
    // Ids 0 and -1 are the separator notes, which the output already has.
    if (!Number.isFinite(numeric) || numeric <= 0) return null;
    if (!(await this.prepare())) return null;
    if (!this.targetRoot || !this.targetDoc || !this.sourceDoc) return null;

    const sourceNote = children(this.sourceDoc.documentElement, this.spec.item).find(
      (el) => attr(el, 'id') === sourceId,
    );
    if (!sourceNote) {
      this.warnings.push(`A ${this.kind} reference pointed at missing content and was dropped.`);
      return null;
    }

    const copy = importNode(this.targetDoc, sourceNote);
    const newId = String(this.nextId++);
    copy.setAttributeNS(NS.w, 'w:id', newId);
    this.remapStyles(copy);
    await this.migrator?.migrate(copy);
    this.targetRoot.appendChild(copy);

    this.idMap.set(sourceId, newId);
    this.copied++;
    this.dirty = true;
    return newId;
  }

  /** Point paragraph and character styles at the reference's equivalents. */
  private remapStyles(note: Element): void {
    const textStyle = this.findReferenceStyle(this.spec.textStyleName);
    for (const pStyle of descendants(note, 'pStyle')) {
      const id = attr(pStyle, 'val');
      if (id && this.referenceStyles.has(id)) continue;
      if (textStyle) pStyle.setAttributeNS(NS.w, 'w:val', textStyle);
      else pStyle.parentNode?.removeChild(pStyle);
    }
    const refStyle = this.findReferenceStyle(this.spec.refStyleName);
    for (const rStyle of descendants(note, 'rStyle')) {
      const id = attr(rStyle, 'val');
      if (id && this.referenceStyles.has(id)) continue;
      if (refStyle) rStyle.setAttributeNS(NS.w, 'w:val', refStyle);
      else rStyle.parentNode?.removeChild(rStyle);
    }
  }

  private findReferenceStyle(pattern: RegExp): string | null {
    return this.referenceStyles.findByPattern(pattern)[0] ?? null;
  }

  /** Write the merged notes part and its relationships. */
  save(): void {
    if (!this.dirty || !this.targetDoc || !this.targetPartPath) return;
    this.target.writeXml(this.targetPartPath, this.targetDoc);
    if (this.targetRels) this.target.saveRels(this.targetPartPath, this.targetRels);
    for (const w of this.migrator?.warnings ?? []) this.warnings.push(w);
  }
}
