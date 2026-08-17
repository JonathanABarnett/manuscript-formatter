import { NS, RELTYPE } from '../ooxml/ns.js';
import { DocxPackage, Relationships } from '../ooxml/package.js';
import { attr, child, children, importNode, parseXml } from '../ooxml/xml.js';

const NUMBERING_CONTENT_TYPE =
  'application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml';

const EMPTY_NUMBERING =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' +
  `<w:numbering xmlns:w="${NS.w}"/>`;

/**
 * Merges the manuscript's list definitions into the output package so numbered
 * and bulleted lists keep their sequence. Ids are offset past whatever the
 * reference already defines, so nothing collides.
 */
export class NumberingMerger {
  private targetDoc: Document | null = null;
  private targetRoot: Element | null = null;
  private targetPartPath = '';
  private nextAbstractId = 0;
  private nextNumId = 1;
  private readonly abstractMap = new Map<string, string>();
  private readonly numMap = new Map<string, string>();
  private dirty = false;
  private ready = false;

  readonly warnings: string[] = [];

  constructor(
    private readonly sourceNumberingDoc: Document | null,
    private readonly target: DocxPackage,
    private readonly targetRels: Relationships,
  ) {}

  /** Locate or create the output numbering part. Safe to call repeatedly. */
  private async prepare(): Promise<boolean> {
    if (this.ready) return this.targetRoot !== null;
    this.ready = true;
    if (!this.sourceNumberingDoc) return false;

    const docDir = this.target.documentDir;
    let target = this.targetRels.firstTargetOfType(RELTYPE.numbering);
    if (!target) {
      target = 'numbering.xml';
      this.targetRels.add(RELTYPE.numbering, target);
      this.target.writeText(`${docDir}numbering.xml`, EMPTY_NUMBERING);
      await this.target.ensureContentType({
        partName: `${docDir}numbering.xml`,
        partType: NUMBERING_CONTENT_TYPE,
      });
      this.dirty = true;
    }
    this.targetPartPath = this.target.resolveTarget(this.target.documentPath, target);
    const doc =
      (await this.target.readXml(this.targetPartPath)) ?? parseXml(EMPTY_NUMBERING);
    this.targetDoc = doc;
    this.targetRoot = doc.documentElement;
    if (!this.targetRoot) return false;

    for (const el of children(this.targetRoot, 'abstractNum')) {
      const id = Number(attr(el, 'abstractNumId') ?? '-1');
      if (Number.isFinite(id) && id >= this.nextAbstractId) this.nextAbstractId = id + 1;
    }
    for (const el of children(this.targetRoot, 'num')) {
      const id = Number(attr(el, 'numId') ?? '0');
      if (Number.isFinite(id) && id >= this.nextNumId) this.nextNumId = id + 1;
    }
    return true;
  }

  /**
   * Map a manuscript `numId` onto one valid in the output, copying the list
   * definition across on first use. Returns null when it cannot be resolved.
   */
  async mapNumId(sourceNumId: string): Promise<string | null> {
    const cached = this.numMap.get(sourceNumId);
    if (cached) return cached;
    if (!(await this.prepare())) return null;

    const sourceRoot = this.sourceNumberingDoc?.documentElement;
    if (!sourceRoot || !this.targetRoot || !this.targetDoc) return null;

    const sourceNum = children(sourceRoot, 'num').find(
      (el) => attr(el, 'numId') === sourceNumId,
    );
    if (!sourceNum) {
      this.warnings.push(`List numbering ${sourceNumId} was not found and the list lost its numbers.`);
      return null;
    }
    const sourceAbstractId = attr(child(sourceNum, 'abstractNumId'), 'val');
    if (!sourceAbstractId) return null;

    const newAbstractId = await this.copyAbstract(sourceRoot, sourceAbstractId);
    if (newAbstractId === null) return null;

    const newNum = importNode(this.targetDoc, sourceNum);
    const newNumId = String(this.nextNumId++);
    newNum.setAttributeNS(NS.w, 'w:numId', newNumId);
    const abstractRef = child(newNum, 'abstractNumId');
    abstractRef?.setAttributeNS(NS.w, 'w:val', newAbstractId);
    this.targetRoot.appendChild(newNum);

    this.numMap.set(sourceNumId, newNumId);
    this.dirty = true;
    return newNumId;
  }

  private async copyAbstract(sourceRoot: Element, sourceAbstractId: string): Promise<string | null> {
    const cached = this.abstractMap.get(sourceAbstractId);
    if (cached) return cached;
    if (!this.targetRoot || !this.targetDoc) return null;

    const sourceAbstract = children(sourceRoot, 'abstractNum').find(
      (el) => attr(el, 'abstractNumId') === sourceAbstractId,
    );
    if (!sourceAbstract) return null;

    const copy = importNode(this.targetDoc, sourceAbstract);
    const newId = String(this.nextAbstractId++);
    copy.setAttributeNS(NS.w, 'w:abstractNumId', newId);
    // Drop the identity fields so Word does not treat this as the same list
    // definition it may already know from the reference document.
    for (const name of ['nsid', 'tmpl'] as const) {
      const el = child(copy, name);
      if (el) copy.removeChild(el);
    }
    // `numStyleLink`/`styleLink` point at style ids that may not exist here.
    for (const name of ['numStyleLink', 'styleLink'] as const) {
      const el = child(copy, name);
      if (el) copy.removeChild(el);
    }

    // abstractNum elements must precede num elements in the part.
    const firstNum = children(this.targetRoot, 'num')[0] ?? null;
    if (firstNum) this.targetRoot.insertBefore(copy, firstNum);
    else this.targetRoot.appendChild(copy);

    this.abstractMap.set(sourceAbstractId, newId);
    this.dirty = true;
    return newId;
  }

  /** Write the merged numbering part back, if anything changed. */
  save(): void {
    if (!this.dirty || !this.targetDoc || !this.targetPartPath) return;
    this.target.writeXml(this.targetPartPath, this.targetDoc);
  }
}
