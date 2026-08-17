import { NS, RELTYPE } from '../ooxml/ns.js';
import { DocxPackage, Relationships, relativeTarget } from '../ooxml/package.js';
import { attr, descendants } from '../ooxml/xml.js';

const IMAGE_MIME: Record<string, string> = {
  png: 'image/png',
  jpg: 'image/jpeg',
  jpeg: 'image/jpeg',
  gif: 'image/gif',
  bmp: 'image/bmp',
  tif: 'image/tiff',
  tiff: 'image/tiff',
  emf: 'image/x-emf',
  wmf: 'image/x-wmf',
  svg: 'image/svg+xml',
  webp: 'image/webp',
};

/** Attributes that hold a relationship id, by owning element. */
const REL_ATTRIBUTES: Array<{ localName: string; ns: string; attrName: string }> = [
  { localName: 'blip', ns: NS.a, attrName: 'embed' },
  { localName: 'blip', ns: NS.a, attrName: 'link' },
  { localName: 'hyperlink', ns: NS.w, attrName: 'id' },
  { localName: 'imagedata', ns: NS.v, attrName: 'id' },
  { localName: 'imagedata', ns: NS.v, attrName: 'href' },
];

/**
 * Copies media and hyperlink relationships from the manuscript package into the
 * output package, rewriting `r:id` references on nodes as they are migrated.
 * One migrator serves one target part (the document, or the footnotes part).
 */
export class ResourceMigrator {
  private readonly idMap = new Map<string, string>();
  private readonly copiedMedia = new Map<string, string>();
  private mediaCounter = 1;

  /** Media parts copied into the output, for reporting. */
  imagesCopied = 0;
  readonly warnings: string[] = [];

  constructor(
    private readonly source: DocxPackage,
    private readonly sourcePartPath: string,
    private readonly sourceRels: Relationships,
    private readonly target: DocxPackage,
    private readonly targetPartPath: string,
    private readonly targetRels: Relationships,
  ) {}

  /** Rewrite every relationship reference inside an already-imported node. */
  async migrate(node: Element): Promise<void> {
    for (const spec of REL_ATTRIBUTES) {
      const matches =
        node.localName === spec.localName && node.namespaceURI === spec.ns
          ? [node, ...descendants(node, spec.localName, spec.ns)]
          : descendants(node, spec.localName, spec.ns);
      for (const el of matches) {
        const oldId = attr(el, spec.attrName, NS.r);
        if (!oldId) continue;
        const newId = await this.remap(oldId);
        if (newId) el.setAttributeNS(NS.r, `r:${spec.attrName}`, newId);
        else el.removeAttributeNS(NS.r, spec.attrName);
      }
    }
  }

  /** Map a source relationship id to one valid in the target package. */
  async remap(sourceId: string): Promise<string | null> {
    const cached = this.idMap.get(sourceId);
    if (cached) return cached;

    const rel = this.sourceRels.byId(sourceId);
    if (!rel) {
      this.warnings.push(`A reference to a missing resource (${sourceId}) was dropped.`);
      return null;
    }

    let newId: string;
    if (rel.targetMode === 'External') {
      newId = this.targetRels.add(rel.type, rel.target, 'External');
    } else if (rel.type === RELTYPE.image) {
      const copied = await this.copyMediaPart(rel.target);
      if (!copied) {
        this.warnings.push(`An image could not be copied and was dropped (${rel.target}).`);
        return null;
      }
      newId = this.targetRels.add(rel.type, relativeTarget(copied, targetDir(this.targetPartPath)));
    } else if (rel.type === RELTYPE.hyperlink) {
      newId = this.targetRels.add(rel.type, rel.target, rel.targetMode ?? 'External');
    } else {
      // Anything else (embedded objects, charts) cannot be carried safely.
      this.warnings.push(
        `An embedded object was dropped because it cannot be transferred (${shortType(rel.type)}).`,
      );
      return null;
    }

    this.idMap.set(sourceId, newId);
    return newId;
  }

  /** Copy a binary part from the manuscript into the output under a fresh name. */
  private async copyMediaPart(target: string): Promise<string | null> {
    const existing = this.copiedMedia.get(target);
    if (existing) return existing;

    const sourcePath = this.source.resolveTarget(this.sourcePartPath, target);
    const data = await this.source.readBinary(sourcePath);
    if (!data) return null;

    const ext = (sourcePath.split('.').pop() ?? 'png').toLowerCase();
    let outPath = `${targetDir(this.targetPartPath)}media/manuscript${this.mediaCounter++}.${ext}`;
    while (this.target.has(outPath)) {
      outPath = `${targetDir(this.targetPartPath)}media/manuscript${this.mediaCounter++}.${ext}`;
    }

    this.target.writeBinary(outPath, data);
    await this.target.ensureContentType({
      extension: ext,
      extensionType: IMAGE_MIME[ext] ?? 'application/octet-stream',
    });
    this.copiedMedia.set(target, outPath);
    this.imagesCopied++;
    return outPath;
  }

  /** Persist the target relationships part. Call once migration is finished. */
  save(): void {
    this.target.saveRels(this.targetPartPath, this.targetRels);
  }
}

function targetDir(partPath: string): string {
  const idx = partPath.lastIndexOf('/');
  return idx === -1 ? '' : partPath.slice(0, idx + 1);
}

function shortType(type: string): string {
  return type.split('/').pop() ?? type;
}
