import JSZip from 'jszip';
import { DOMParser, XMLSerializer } from '@xmldom/xmldom';

export const NS = {
  w: 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
  a: 'http://schemas.openxmlformats.org/drawingml/2006/main',
  p: 'http://schemas.openxmlformats.org/presentationml/2006/main',
  r: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
  rel: 'http://schemas.openxmlformats.org/package/2006/relationships',
};

// xmldom returns '' for missing attributes; the DOM spec (and the `?? default` fallbacks here) expect null.
{
  const proto = Object.getPrototypeOf(new DOMParser().parseFromString('<x/>', 'text/xml').documentElement);
  const getAttribute = proto.getAttribute;
  const getAttributeNS = proto.getAttributeNS;
  proto.getAttribute = function (this: Element, name: string) {
    return this.hasAttribute(name) ? getAttribute.call(this, name) : null;
  };
  proto.getAttributeNS = function (this: Element, ns: string, name: string) {
    return this.hasAttributeNS(ns, name) ? getAttributeNS.call(this, ns, name) : null;
  };
}

export async function openZip(bytes: Uint8Array) {
  return JSZip.loadAsync(bytes);
}

export async function readXml(zip: JSZip, path: string): Promise<Document | undefined> {
  const file = zip.file(path);
  if (!file) return undefined;
  return new DOMParser().parseFromString(await file.async('string'), 'text/xml') as unknown as Document;
}

export function writeXml(zip: JSZip, path: string, doc: Document) {
  zip.file(path, new XMLSerializer().serializeToString(doc as any));
}

export async function zipBytes(zip: JSZip): Promise<Uint8Array> {
  return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
}

export function kids(el: Node | null | undefined, ns?: string, local?: string): Element[] {
  const out: Element[] = [];
  if (!el) return out;
  for (let n = el.firstChild; n; n = n.nextSibling) {
    if (n.nodeType !== 1) continue;
    const e = n as Element;
    if (ns && e.namespaceURI !== ns) continue;
    if (local && e.localName !== local) continue;
    out.push(e);
  }
  return out;
}

export const kid = (el: Node | null | undefined, ns: string, local: string): Element | undefined => kids(el, ns, local)[0];

/** Follow a path of direct children in one namespace, e.g. path(el, NS.a, 'xfrm', 'off'). */
export function path(el: Node | null | undefined, ns: string, ...locals: string[]): Element | undefined {
  let cur: Element | undefined = el as Element | undefined;
  for (const l of locals) {
    cur = kid(cur, ns, l);
    if (!cur) return undefined;
  }
  return cur;
}

export function all(el: Node | null | undefined, ns: string, local: string): Element[] {
  if (!el) return [];
  return Array.from((el as Element).getElementsByTagNameNS(ns, local));
}

export const is = (n: Node | null | undefined, ns: string, local: string) =>
  !!n && n.nodeType === 1 && (n as Element).namespaceURI === ns && (n as Element).localName === local;

export function wattr(el: Element | undefined, name: string): string | null {
  if (!el) return null;
  return el.getAttributeNS(NS.w, name) || el.getAttribute('w:' + name) || null;
}

export interface Rel {
  target: string;
  type: string;
  external: boolean;
}

/** Parse the relationships of a part, resolving targets to zip paths. */
export async function readRels(zip: JSZip, partPath: string): Promise<Map<string, Rel>> {
  const dir = partPath.includes('/') ? partPath.slice(0, partPath.lastIndexOf('/')) : '';
  const name = partPath.slice(partPath.lastIndexOf('/') + 1);
  const doc = await readXml(zip, `${dir ? dir + '/' : ''}_rels/${name}.rels`);
  const map = new Map<string, Rel>();
  if (!doc) return map;
  for (const rel of all(doc, NS.rel, 'Relationship')) {
    const external = rel.getAttribute('TargetMode') === 'External';
    const target = rel.getAttribute('Target') ?? '';
    map.set(rel.getAttribute('Id') ?? '', {
      target: external ? target : resolvePath(dir, target),
      type: rel.getAttribute('Type') ?? '',
      external,
    });
  }
  return map;
}

function resolvePath(dir: string, target: string): string {
  if (target.startsWith('/')) return target.slice(1);
  const parts = dir ? dir.split('/') : [];
  for (const seg of target.split('/')) {
    if (seg === '..') parts.pop();
    else if (seg !== '.' && seg) parts.push(seg);
  }
  return parts.join('/');
}

const MIME: Record<string, string> = {
  png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg', gif: 'image/gif', bmp: 'image/bmp',
  svg: 'image/svg+xml', webp: 'image/webp', tif: 'image/tiff', tiff: 'image/tiff',
};

/** Load an image part as a data URI; returns undefined for formats a browser can't show (emf/wmf). */
export async function imageDataUri(zip: JSZip, p: string): Promise<string | undefined> {
  const mime = MIME[p.split('.').pop()?.toLowerCase() ?? ''];
  const file = zip.file(p);
  if (!mime || !file) return undefined;
  return `data:${mime};base64,${await file.async('base64')}`;
}
