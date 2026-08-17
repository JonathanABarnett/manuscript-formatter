import JSZip from 'jszip';
import { NS, RELTYPE } from './ns.js';
import { children, parseXml, serializeXml } from './xml.js';

export interface Relationship {
  id: string;
  type: string;
  target: string;
  targetMode?: string;
}

/** The relationships of a single part (`_rels/<part>.rels`). */
export class Relationships {
  private readonly items = new Map<string, Relationship>();
  private nextId = 1;

  constructor(xml?: string) {
    if (!xml) return;
    const doc = parseXml(xml);
    const root = doc.documentElement;
    if (!root) return;
    for (const el of children(root, 'Relationship', NS.pkgRel)) {
      const id = el.getAttribute('Id');
      const type = el.getAttribute('Type');
      const target = el.getAttribute('Target');
      if (!id || !type || target === null) continue;
      const targetMode = el.getAttribute('TargetMode') || undefined;
      this.items.set(id, { id, type, target, targetMode });
      const numeric = Number(id.replace(/^rId/, ''));
      if (Number.isFinite(numeric) && numeric >= this.nextId) this.nextId = numeric + 1;
    }
  }

  all(): Relationship[] {
    return [...this.items.values()];
  }

  byId(id: string): Relationship | undefined {
    return this.items.get(id);
  }

  byType(type: string): Relationship[] {
    return this.all().filter((r) => r.type === type);
  }

  firstTargetOfType(type: string): string | undefined {
    return this.byType(type)[0]?.target;
  }

  /**
   * Add a relationship and return its new id. Internal (non-external)
   * relationships to an identical target are reused so repeated images or
   * hyperlinks do not multiply entries.
   */
  add(type: string, target: string, targetMode?: string): string {
    if (!targetMode) {
      const existing = this.all().find(
        (r) => r.type === type && r.target === target && !r.targetMode,
      );
      if (existing) return existing.id;
    }
    const id = `rId${this.nextId++}`;
    this.items.set(id, { id, type, target, targetMode });
    return id;
  }

  remove(id: string): void {
    this.items.delete(id);
  }

  toXml(): string {
    const escaped = (s: string): string =>
      s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
    const lines = this.all().map((r) => {
      const mode = r.targetMode ? ` TargetMode="${escaped(r.targetMode)}"` : '';
      return `<Relationship Id="${escaped(r.id)}" Type="${escaped(r.type)}" Target="${escaped(
        r.target,
      )}"${mode}/>`;
    });
    return (
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' +
      `<Relationships xmlns="${NS.pkgRel}">${lines.join('')}</Relationships>`
    );
  }
}

/**
 * A .docx read into memory. Wraps the zip container and exposes the pieces the
 * formatter needs: parts by name, per-part relationships, and [Content_Types].
 */
export class DocxPackage {
  private constructor(
    private readonly zip: JSZip,
    /** Path of the main document part, e.g. `word/document.xml`. */
    readonly documentPath: string,
  ) {}

  static async fromBuffer(buffer: Uint8Array, label = 'document'): Promise<DocxPackage> {
    let zip: JSZip;
    try {
      zip = await JSZip.loadAsync(buffer);
    } catch {
      throw new Error(
        `"${label}" is not a valid .docx file. If it is a .doc, .rtf or .odt, ` +
          'open it in Word and save as .docx first.',
      );
    }
    if (!zip.file('[Content_Types].xml')) {
      throw new Error(`"${label}" is a zip archive but not a Word document.`);
    }
    const rootRels = await readText(zip, '_rels/.rels');
    const documentPath = rootRels
      ? normalizePartPath(
          new Relationships(rootRels).firstTargetOfType(RELTYPE.officeDocument) ??
            'word/document.xml',
          '',
        )
      : 'word/document.xml';
    if (!zip.file(documentPath)) {
      throw new Error(`"${label}" has no main document part (expected ${documentPath}).`);
    }
    return new DocxPackage(zip, documentPath);
  }

  /** Directory holding the main document, e.g. `word/`. */
  get documentDir(): string {
    const idx = this.documentPath.lastIndexOf('/');
    return idx === -1 ? '' : this.documentPath.slice(0, idx + 1);
  }

  has(partPath: string): boolean {
    return this.zip.file(partPath) !== null;
  }

  listParts(): string[] {
    const out: string[] = [];
    this.zip.forEach((path, entry) => {
      if (!entry.dir) out.push(path);
    });
    return out;
  }

  async readText(partPath: string): Promise<string | null> {
    return readText(this.zip, partPath);
  }

  async readXml(partPath: string): Promise<Document | null> {
    const text = await this.readText(partPath);
    if (text === null) return null;
    try {
      return parseXml(text);
    } catch (err) {
      throw new Error(`Cannot parse ${partPath}: ${(err as Error).message}`);
    }
  }

  async readBinary(partPath: string): Promise<Uint8Array | null> {
    const file = this.zip.file(partPath);
    if (!file) return null;
    return file.async('uint8array');
  }

  writeText(partPath: string, content: string): void {
    this.zip.file(partPath, content);
  }

  writeXml(partPath: string, doc: Document | Element): void {
    this.zip.file(partPath, serializeXml(doc));
  }

  writeBinary(partPath: string, content: Uint8Array): void {
    this.zip.file(partPath, content);
  }

  removePart(partPath: string): void {
    this.zip.remove(partPath);
  }

  /** Path of the `.rels` part belonging to `partPath`. */
  static relsPathFor(partPath: string): string {
    const idx = partPath.lastIndexOf('/');
    const dir = idx === -1 ? '' : partPath.slice(0, idx + 1);
    const name = idx === -1 ? partPath : partPath.slice(idx + 1);
    return `${dir}_rels/${name}.rels`;
  }

  async relsFor(partPath: string): Promise<Relationships> {
    const xml = await this.readText(DocxPackage.relsPathFor(partPath));
    return new Relationships(xml ?? undefined);
  }

  saveRels(partPath: string, rels: Relationships): void {
    this.writeText(DocxPackage.relsPathFor(partPath), rels.toXml());
  }

  /** Resolve a relationship target (relative to the owning part) to a part path. */
  resolveTarget(ownerPartPath: string, target: string): string {
    const idx = ownerPartPath.lastIndexOf('/');
    const dir = idx === -1 ? '' : ownerPartPath.slice(0, idx + 1);
    return normalizePartPath(target, dir);
  }

  /**
   * Ensure `[Content_Types].xml` declares an extension default (for binaries
   * such as .png) and/or a part override (for XML parts such as footnotes).
   */
  async ensureContentType(options: {
    extension?: string;
    extensionType?: string;
    partName?: string;
    partType?: string;
  }): Promise<void> {
    const path = '[Content_Types].xml';
    const text = await this.readText(path);
    if (!text) return;
    const doc = parseXml(text);
    const root = doc.documentElement;
    if (!root) return;
    let changed = false;

    if (options.extension && options.extensionType) {
      const ext = options.extension.toLowerCase().replace(/^\./, '');
      const present = children(root, 'Default', NS.contentTypes).some(
        (el) => (el.getAttribute('Extension') ?? '').toLowerCase() === ext,
      );
      if (!present) {
        const el = doc.createElementNS(NS.contentTypes, 'Default');
        el.setAttribute('Extension', ext);
        el.setAttribute('ContentType', options.extensionType);
        root.insertBefore(el, root.firstChild);
        changed = true;
      }
    }

    if (options.partName && options.partType) {
      const partName = options.partName.startsWith('/')
        ? options.partName
        : `/${options.partName}`;
      const present = children(root, 'Override', NS.contentTypes).some(
        (el) => el.getAttribute('PartName') === partName,
      );
      if (!present) {
        const el = doc.createElementNS(NS.contentTypes, 'Override');
        el.setAttribute('PartName', partName);
        el.setAttribute('ContentType', options.partType);
        root.appendChild(el);
        changed = true;
      }
    }

    if (changed) this.writeXml(path, doc);
  }

  /** The package as bytes, ready to write to disk or hand to a download. */
  async toBuffer(): Promise<Uint8Array> {
    return this.zip.generateAsync({
      type: 'uint8array',
      compression: 'DEFLATE',
      compressionOptions: { level: 6 },
      mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    });
  }

  /** Deep copy, so the reference package can be reused as a base repeatedly. */
  async clone(): Promise<DocxPackage> {
    return DocxPackage.fromBuffer(await this.toBuffer(), 'clone');
  }
}

async function readText(zip: JSZip, partPath: string): Promise<string | null> {
  const file = zip.file(partPath);
  if (!file) return null;
  return file.async('string');
}

/** Collapse `../` and leading `/` in a relationship target. */
export function normalizePartPath(target: string, baseDir: string): string {
  if (target.startsWith('/')) return target.slice(1);
  const segments = (baseDir + target).split('/');
  const out: string[] = [];
  for (const segment of segments) {
    if (segment === '.' || segment === '') continue;
    if (segment === '..') out.pop();
    else out.push(segment);
  }
  return out.join('/');
}

/** Relationship target path expressed relative to the owning part's folder. */
export function relativeTarget(partPath: string, ownerDir: string): string {
  if (ownerDir && partPath.startsWith(ownerDir)) return partPath.slice(ownerDir.length);
  return `/${partPath}`;
}
