import type JSZip from 'jszip';
import type { FormatHandler } from '../provider';
import { NS, all, imageDataUri, is, kid, kids, openZip, path, readRels, readXml, writeXml, zipBytes, Rel } from './ooxml';

const A = NS.a;
const P = NS.p;
const EMU_PER_PX = 9525;

// ---------- Model shared with the webview ----------

export interface PRun {
  /** Run index within its paragraph; absent for read-only text. */
  r?: number;
  text: string;
  b?: boolean;
  i?: boolean;
  u?: boolean;
  sz: number; // px
  color?: string;
  font?: string;
}
export interface PPara {
  /** "slide:paragraph" id; absent for read-only text (layout/master decorations). */
  id?: string;
  align: string;
  bullet?: string;
  level: number;
  items: (PRun | { br: true })[];
  /** Size of an empty paragraph's text (px). */
  sz: number;
}
interface Box {
  x: number;
  y: number;
  w: number;
  h: number;
  rot?: number;
}
export type PElement =
  | (Box & { k: 'text'; fill?: string; anchor: string; paras: PPara[]; vert?: boolean })
  | (Box & { k: 'img'; src: string })
  | (Box & { k: 'table'; cols: number[]; rows: { h: number; cells: { span: number; rowSpan: number; fill?: string; paras: PPara[] }[] }[] })
  | (Box & { k: 'placeholder'; label: string });
export interface PSlide {
  bg: string;
  bgImage?: string;
  hidden?: boolean;
  elements: PElement[];
}
export interface PptxModel {
  width: number;
  height: number;
  slides: PSlide[];
}
export interface PptxPayload {
  paras: { id: string; runs: Record<number, string>; extra: string | null }[];
}

// ---------- Colors ----------

interface Theme {
  colors: Map<string, string>;
  major?: string;
  minor?: string;
}

function hexToHsl(hex: string): [number, number, number] {
  const r = parseInt(hex.slice(0, 2), 16) / 255, g = parseInt(hex.slice(2, 4), 16) / 255, b = parseInt(hex.slice(4, 6), 16) / 255;
  const max = Math.max(r, g, b), min = Math.min(r, g, b);
  const l = (max + min) / 2;
  if (max === min) return [0, 0, l];
  const d = max - min;
  const s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
  const h = max === r ? (g - b) / d + (g < b ? 6 : 0) : max === g ? (b - r) / d + 2 : (r - g) / d + 4;
  return [h / 6, s, l];
}

function hslToHex(h: number, s: number, l: number): string {
  const f = (p: number, q: number, t: number) => {
    if (t < 0) t += 1;
    if (t > 1) t -= 1;
    if (t < 1 / 6) return p + (q - p) * 6 * t;
    if (t < 1 / 2) return q;
    if (t < 2 / 3) return p + (q - p) * (2 / 3 - t) * 6;
    return p;
  };
  let r = l, g = l, b = l;
  if (s) {
    const q = l < 0.5 ? l * (1 + s) : l + s - l * s;
    const p = 2 * l - q;
    r = f(p, q, h + 1 / 3);
    g = f(p, q, h);
    b = f(p, q, h - 1 / 3);
  }
  return [r, g, b].map((x) => Math.round(Math.min(1, Math.max(0, x)) * 255).toString(16).padStart(2, '0')).join('');
}

class Colors {
  constructor(private theme: Theme, private clrMap: Record<string, string>) {}

  /** Resolve a color element (srgbClr, schemeClr, …) to "#rrggbb". */
  ofColor(c: Element | undefined): string | undefined {
    if (!c) return undefined;
    let hex: string | undefined;
    switch (c.localName) {
      case 'srgbClr': hex = c.getAttribute('val') ?? undefined; break;
      case 'sysClr': hex = c.getAttribute('lastClr') ?? (c.getAttribute('val') === 'window' ? 'FFFFFF' : '000000'); break;
      case 'schemeClr': {
        let v = c.getAttribute('val') ?? '';
        v = this.clrMap[v] ?? v;
        hex = this.theme.colors.get(v);
        break;
      }
      case 'prstClr': hex = ({ black: '000000', white: 'FFFFFF', red: 'FF0000', blue: '0000FF', green: '008000' } as Record<string, string>)[c.getAttribute('val') ?? '']; break;
    }
    if (!hex) return undefined;
    let [h, s, l] = hexToHsl(hex);
    for (const m of kids(c, A)) {
      const v = Number(m.getAttribute('val') ?? 0) / 100000;
      if (m.localName === 'lumMod') l *= v;
      else if (m.localName === 'lumOff') l += v;
      else if (m.localName === 'tint') l = l * v + (1 - v);
      else if (m.localName === 'shade') l *= v;
    }
    return '#' + hslToHex(h, s, l);
  }

  /** Resolve a fill container (spPr, bgPr, tcPr, …) to a CSS color; null means explicitly no fill. */
  ofFill(container: Element | undefined): string | null | undefined {
    if (!container) return undefined;
    if (kid(container, A, 'noFill')) return null;
    const solid = kid(container, A, 'solidFill');
    if (solid) return this.ofColor(kids(solid, A)[0]);
    const grad = kid(container, A, 'gradFill');
    if (grad) return this.ofColor(kids(all(grad, A, 'gs')[0], A)[0]);
    return undefined;
  }
}

async function readTheme(zip: JSZip, masterPath: string): Promise<Theme> {
  const rels = await readRels(zip, masterPath);
  const themeRel = [...rels.values()].find((r) => r.type.endsWith('/theme'));
  const doc = themeRel ? await readXml(zip, themeRel.target) : undefined;
  const colors = new Map<string, string>();
  const scheme = all(doc, A, 'clrScheme')[0];
  for (const c of kids(scheme, A)) {
    const v = kids(c, A)[0];
    const hex = v?.localName === 'srgbClr' ? v.getAttribute('val') : v?.getAttribute('lastClr');
    if (hex) colors.set(c.localName!, hex);
  }
  const font = (name: string) => kid(all(doc, A, name)[0], A, 'latin')?.getAttribute('typeface') ?? undefined;
  return { colors, major: font('majorFont'), minor: font('minorFont') };
}

// ---------- Parts ----------

interface Part {
  path: string;
  doc: Document;
  rels: Map<string, Rel>;
  tree: Element;
}

async function readPart(zip: JSZip, p: string): Promise<Part | undefined> {
  const doc = await readXml(zip, p);
  if (!doc) return undefined;
  const tree = all(doc, P, 'spTree')[0];
  return { path: p, doc, rels: await readRels(zip, p), tree };
}

const relOfType = (part: Part, type: string) => [...part.rels.values()].find((r) => r.type.endsWith('/' + type));

async function slidePaths(zip: JSZip): Promise<string[]> {
  const pres = await readXml(zip, 'ppt/presentation.xml');
  const rels = await readRels(zip, 'ppt/presentation.xml');
  return all(pres, P, 'sldId')
    .map((s) => rels.get(s.getAttributeNS(NS.r, 'id') || s.getAttribute('r:id') || '')?.target)
    .filter((t): t is string => !!t);
}

interface Ph {
  type: string;
  idx?: string;
}

function phOf(sp: Element): Ph | undefined {
  const nv = kids(sp).find((c) => c.localName?.startsWith('nv'));
  const ph = path(nv, P, 'nvPr', 'ph');
  if (!ph) return undefined;
  return { type: ph.getAttribute('type') ?? 'body', idx: ph.getAttribute('idx') ?? undefined };
}

const phFamily = (t: string) => (t === 'ctrTitle' ? 'title' : t === 'subTitle' || t === 'obj' ? 'body' : t);

function findPh(part: Part | undefined, ph: Ph): Element | undefined {
  if (!part) return undefined;
  const shapes = kids(part.tree).filter((c) => phOf(c));
  if (ph.idx !== undefined) {
    const byIdx = shapes.find((s) => phOf(s)!.idx === ph.idx);
    if (byIdx) return byIdx;
  }
  return shapes.find((s) => phOf(s)!.type === ph.type) ?? shapes.find((s) => phFamily(phOf(s)!.type) === phFamily(ph.type));
}

// ---------- Rendering ----------

interface Transform {
  (x: number, y: number, w: number, h: number): Box;
}
const identity: Transform = (x, y, w, h) => ({ x, y, w, h });

function readXfrm(xfrm: Element | undefined) {
  const off = kid(xfrm, A, 'off');
  const ext = kid(xfrm, A, 'ext');
  if (!off || !ext) return undefined;
  return {
    x: Number(off.getAttribute('x')), y: Number(off.getAttribute('y')),
    w: Number(ext.getAttribute('cx')), h: Number(ext.getAttribute('cy')),
    rot: Number(xfrm!.getAttribute('rot') ?? 0) / 60000,
  };
}

class SlideRenderer {
  private paraIndex = new Map<Element, number>();
  readonly elements: PElement[] = [];

  constructor(
    private zip: JSZip,
    private slideNo: number,
    private slide: Part,
    private layout: Part | undefined,
    private master: Part | undefined,
    private colors: Colors,
    private theme: Theme,
    private defaultTextStyle: Element | undefined,
  ) {
    all(slide.tree, A, 'p').forEach((p, i) => this.paraIndex.set(p, i));
  }

  async render() {
    const showMaster = (part: Part | undefined) => part?.doc.documentElement.getAttribute('showMasterSp') !== '0';
    // Decorations from the master and layout sit behind the slide's own shapes.
    if (this.master && showMaster(this.layout) && showMaster(this.slide)) await this.renderTree(this.master.tree, this.master, identity, true);
    if (this.layout && showMaster(this.slide)) await this.renderTree(this.layout.tree, this.layout, identity, true);
    await this.renderTree(this.slide.tree, this.slide, identity, false);
  }

  private async renderTree(tree: Element, part: Part, tf: Transform, decorative: boolean) {
    for (const el of kids(tree, P)) {
      if (decorative && phOf(el)) continue; // placeholders only show their slide content
      switch (el.localName) {
        case 'sp': this.renderShape(el, part, tf, decorative); break;
        case 'pic': await this.renderPic(el, part, tf); break;
        case 'grpSp': await this.renderGroup(el, part, tf, decorative); break;
        case 'graphicFrame': this.renderFrame(el, part, tf, decorative); break;
      }
    }
  }

  private box(el: Element, tf: Transform): Box | undefined {
    const spPr = kids(el).find((c) => c.localName === 'spPr' || c.localName === 'grpSpPr');
    let xfrm = readXfrm(kid(spPr, A, 'xfrm') ?? kid(el, P, 'xfrm'));
    const ph = phOf(el);
    if (!xfrm && ph) {
      const xfrmOf = (shape: Element | undefined) => readXfrm(kid(kid(shape, P, 'spPr'), A, 'xfrm'));
      xfrm = xfrmOf(findPh(this.layout, ph)) ?? xfrmOf(findPh(this.master, ph));
    }
    if (!xfrm) return undefined;
    const b = tf(xfrm.x, xfrm.y, xfrm.w, xfrm.h);
    return { x: b.x / EMU_PER_PX, y: b.y / EMU_PER_PX, w: b.w / EMU_PER_PX, h: b.h / EMU_PER_PX, rot: xfrm.rot || undefined };
  }

  private async renderGroup(el: Element, part: Part, tf: Transform, decorative: boolean) {
    const xfrm = kid(kid(el, P, 'grpSpPr'), A, 'xfrm');
    const g = readXfrm(xfrm);
    const chOff = kid(xfrm, A, 'chOff');
    const chExt = kid(xfrm, A, 'chExt');
    let inner = tf;
    if (g && chOff && chExt) {
      const cx = Number(chOff.getAttribute('x')), cy = Number(chOff.getAttribute('y'));
      const sx = g.w / (Number(chExt.getAttribute('cx')) || 1), sy = g.h / (Number(chExt.getAttribute('cy')) || 1);
      inner = (x, y, w, h) => tf(g.x + (x - cx) * sx, g.y + (y - cy) * sy, w * sx, h * sy);
    }
    await this.renderTree(el, part, inner, decorative);
  }

  private async renderPic(el: Element, part: Part, tf: Transform) {
    const box = this.box(el, tf);
    const blip = all(el, A, 'blip')[0];
    const rid = blip?.getAttributeNS(NS.r, 'embed') || blip?.getAttribute('r:embed');
    const rel = rid ? part.rels.get(rid) : undefined;
    if (!box || !rel || rel.external) return;
    const src = await imageDataUri(this.zip, rel.target);
    if (src) this.elements.push({ k: 'img', ...box, src });
  }

  private renderShape(sp: Element, part: Part, tf: Transform, decorative: boolean) {
    const box = this.box(sp, tf);
    if (!box) return;
    const ph = phOf(sp);
    const spPr = kid(sp, P, 'spPr');
    let fill = this.colors.ofFill(spPr);
    if (fill === undefined && !ph) {
      // Preset shapes without explicit fill use the theme style's fill.
      const ref = kid(kid(sp, P, 'style'), A, 'fillRef');
      if (ref && ref.getAttribute('idx') !== '0') fill = this.colors.ofColor(kids(ref, A)[0]);
    }
    const txBody = kid(sp, P, 'txBody');
    const layoutPh = ph ? findPh(this.layout, ph) : undefined;
    const masterPh = ph ? findPh(this.master, ph) : undefined;
    const bodyPrs = [kid(txBody, A, 'bodyPr'), kid(kid(layoutPh, P, 'txBody'), A, 'bodyPr'), kid(kid(masterPh, P, 'txBody'), A, 'bodyPr')];
    const anchor = bodyPrs.map((b) => b?.getAttribute('anchor')).find(Boolean) ?? 't';
    const vert = bodyPrs.map((b) => b?.getAttribute('vert')).find(Boolean);
    const scaleAttr = all(kid(txBody, A, 'bodyPr'), A, 'normAutofit')[0]?.getAttribute('fontScale');
    const scale = scaleAttr ? Number(scaleAttr) / 100000 : 1;
    const fontColor = this.colors.ofColor(kids(kid(kid(sp, P, 'style'), A, 'fontRef'), A)[0]);

    const styleKind = !ph ? 'other' : phFamily(ph.type) === 'title' ? 'title' : ['body'].includes(phFamily(ph.type)) ? 'body' : 'other';
    const txStyle = path(this.master?.doc.documentElement, P, 'txStyles', styleKind === 'title' ? 'titleStyle' : styleKind === 'body' ? 'bodyStyle' : 'otherStyle');
    const lstStyles = [kid(txBody, A, 'lstStyle'), kid(kid(layoutPh, P, 'txBody'), A, 'lstStyle'), kid(kid(masterPh, P, 'txBody'), A, 'lstStyle'), ph ? txStyle : this.defaultTextStyle];

    const paras = txBody ? this.paragraphs(txBody, lstStyles, scale, fontColor, decorative) : [];
    if (!paras.length && fill == null) return;
    this.elements.push({ k: 'text', ...box, fill: fill ?? undefined, anchor, paras, vert: vert === 'vert' || vert === 'eaVert' || undefined });
  }

  private paragraphs(txBody: Element, lstStyles: (Element | undefined)[], scale: number, fontColor: string | undefined, readOnly: boolean): PPara[] {
    const out: PPara[] = [];
    const autoNums = new Map<number, number>();
    for (const p of kids(txBody, A, 'p')) {
      const pPr = kid(p, A, 'pPr');
      const level = Number(pPr?.getAttribute('lvl') ?? 0);
      const lvlName = `lvl${level + 1}pPr`;
      const pChain = [pPr, ...lstStyles.map((s) => kid(s, A, lvlName))].filter((e): e is Element => !!e);
      const rChainBase = pChain.map((e) => kid(e, A, 'defRPr')).filter((e): e is Element => !!e);
      const pick = (attr: string) => pChain.map((e) => e.getAttribute(attr)).find((v) => v != null);
      const align = ({ l: 'left', ctr: 'center', r: 'right', just: 'justify', dist: 'justify' } as Record<string, string>)[pick('algn') ?? 'l'] ?? 'left';

      let bullet: string | undefined;
      for (const e of pChain) {
        if (kid(e, A, 'buNone')) break;
        const ch = kid(e, A, 'buChar');
        if (ch) { bullet = ch.getAttribute('char') ?? '•'; break; }
        if (kid(e, A, 'buAutoNum')) {
          const n = (autoNums.get(level) ?? 0) + 1;
          autoNums.set(level, n);
          bullet = `${n}.`;
          break;
        }
      }

      const resolve = (rPr: Element | undefined) => {
        const chain = [rPr, ...rChainBase].filter((e): e is Element => !!e);
        const attr = (n: string) => chain.map((e) => e.getAttribute(n)).find((v) => v != null);
        const sz = Number(attr('sz') ?? 1800) / 100;
        const color = chain.map((e) => this.colors.ofFill(e)).find((c) => c) ?? fontColor;
        const latin = chain.map((e) => kid(e, A, 'latin')?.getAttribute('typeface')).find(Boolean);
        const font = latin === '+mj-lt' ? this.theme.major : latin === '+mn-lt' || !latin ? this.theme.minor : latin;
        const flag = (n: string) => { const v = attr(n); return v == null ? undefined : v === '1' || v === 'true'; };
        const u = attr('u');
        return { sz: (sz * scale * 96) / 72, color: color ?? undefined, font, b: flag('b'), i: flag('i'), u: u ? u !== 'none' : undefined };
      };

      const items: PPara['items'] = [];
      let r = 0;
      for (const c of kids(p, A)) {
        if (c.localName === 'r' || c.localName === 'fld') {
          const run: PRun = { ...resolve(kid(c, A, 'rPr')), text: kid(c, A, 't')?.textContent ?? '' };
          if (!readOnly) run.r = r;
          items.push(run);
          r++;
        } else if (c.localName === 'br') items.push({ br: true });
      }
      if (readOnly && !items.length) continue;
      const para: PPara = { align, level, items, sz: resolve(kid(p, A, 'endParaRPr')).sz };
      if (bullet && items.length) para.bullet = bullet;
      if (!readOnly && this.paraIndex.has(p)) para.id = `${this.slideNo}:${this.paraIndex.get(p)}`;
      out.push(para);
    }
    // Drop text boxes that only contain empty, read-only paragraphs.
    return out.some((p) => p.id || p.items.length) ? out : [];
  }

  private renderFrame(el: Element, part: Part, tf: Transform, decorative: boolean) {
    const box = this.box(el, tf);
    if (!box) return;
    const tbl = all(el, A, 'tbl')[0];
    if (!tbl) {
      const uri = all(el, A, 'graphicData')[0]?.getAttribute('uri') ?? '';
      const label = uri.includes('chart') ? 'Chart' : uri.includes('diagram') ? 'SmartArt' : 'Object';
      this.elements.push({ k: 'placeholder', ...box, label });
      return;
    }
    const cols = all(kid(tbl, A, 'tblGrid'), A, 'gridCol').map((g) => Number(g.getAttribute('w')) / EMU_PER_PX);
    const rows = kids(tbl, A, 'tr').map((tr) => ({
      h: Number(tr.getAttribute('h') ?? 0) / EMU_PER_PX,
      cells: kids(tr, A, 'tc')
        .filter((tc) => tc.getAttribute('hMerge') !== '1' && tc.getAttribute('vMerge') !== '1')
        .map((tc) => {
          const txBody = kid(tc, A, 'txBody');
          return {
            span: Number(tc.getAttribute('gridSpan') ?? 1),
            rowSpan: Number(tc.getAttribute('rowSpan') ?? 1),
            fill: this.colors.ofFill(kid(tc, A, 'tcPr')) ?? undefined,
            paras: txBody ? this.paragraphs(txBody, [kid(txBody, A, 'lstStyle'), this.defaultTextStyle], 1, undefined, decorative) : [],
          };
        }),
    }));
    this.elements.push({ k: 'table', ...box, cols, rows });
  }
}

async function background(zip: JSZip, parts: (Part | undefined)[], colors: Colors): Promise<{ bg: string; bgImage?: string }> {
  for (const part of parts) {
    const bg = path(part?.doc.documentElement, P, 'cSld', 'bg');
    if (!bg) continue;
    const bgPr = kid(bg, P, 'bgPr');
    if (bgPr) {
      const blip = all(bgPr, A, 'blip')[0];
      const rid = blip?.getAttributeNS(NS.r, 'embed') || blip?.getAttribute('r:embed');
      const rel = rid ? part!.rels.get(rid) : undefined;
      if (rel && !rel.external) {
        const img = await imageDataUri(zip, rel.target);
        if (img) return { bg: '#ffffff', bgImage: img };
      }
      const c = colors.ofFill(bgPr);
      if (c) return { bg: c };
    }
    const ref = kid(bg, P, 'bgRef');
    if (ref) {
      const c = colors.ofColor(kids(ref, A)[0]);
      if (c) return { bg: c };
    }
  }
  return { bg: '#ffffff' };
}

// ---------- Handler ----------

async function load(bytes: Uint8Array): Promise<PptxModel> {
  const zip = await openZip(bytes);
  const pres = await readXml(zip, 'ppt/presentation.xml');
  if (!pres) throw new Error('Not a valid PowerPoint file (ppt/presentation.xml missing)');
  const sz = all(pres, P, 'sldSz')[0];
  const width = Number(sz?.getAttribute('cx') ?? 12192000) / EMU_PER_PX;
  const height = Number(sz?.getAttribute('cy') ?? 6858000) / EMU_PER_PX;
  const defaultTextStyle = all(pres, P, 'defaultTextStyle')[0];

  const slides: PSlide[] = [];
  const paths = await slidePaths(zip);
  for (let s = 0; s < paths.length; s++) {
    const slide = await readPart(zip, paths[s]);
    if (!slide) continue;
    const layoutRel = relOfType(slide, 'slideLayout');
    const layout = layoutRel ? await readPart(zip, layoutRel.target) : undefined;
    const masterRel = layout && relOfType(layout, 'slideMaster');
    const master = masterRel ? await readPart(zip, masterRel.target) : undefined;
    const theme = master ? await readTheme(zip, master.path) : { colors: new Map<string, string>() };
    const clrMapEl = master && all(master.doc, P, 'clrMap')[0];
    const clrMap: Record<string, string> = { tx1: 'dk1', bg1: 'lt1', tx2: 'dk2', bg2: 'lt2' };
    if (clrMapEl) for (const attr of Array.from(clrMapEl.attributes)) clrMap[attr.name] = attr.value;
    const colors = new Colors(theme, clrMap);

    const renderer = new SlideRenderer(zip, s, slide, layout, master, colors, theme, defaultTextStyle);
    await renderer.render();
    slides.push({
      ...(await background(zip, [slide, layout, master], colors)),
      hidden: slide.doc.documentElement.getAttribute('show') === '0' || undefined,
      elements: renderer.elements,
    });
  }
  return { width, height, slides };
}

async function save(base: Uint8Array, payload: PptxPayload): Promise<Uint8Array> {
  const zip = await openZip(base);
  const paths = await slidePaths(zip);
  const bySlide = new Map<number, PptxPayload['paras']>();
  for (const p of payload.paras) {
    const s = Number(p.id.split(':')[0]);
    bySlide.set(s, [...(bySlide.get(s) ?? []), p]);
  }
  for (const [s, edits] of bySlide) {
    const doc = paths[s] && (await readXml(zip, paths[s]));
    if (!doc) continue;
    const tree = all(doc, P, 'spTree')[0];
    const paras = all(tree, A, 'p');
    for (const edit of edits) {
      const p = paras[Number(edit.id.split(':')[1])];
      if (!p) continue;
      const runs = kids(p, A).filter((c) => c.localName === 'r' || c.localName === 'fld');
      const template = kid(runs[0], A, 'rPr') ?? kid(p, A, 'endParaRPr');
      runs.forEach((run, i) => {
        const text = edit.runs[i];
        if (text === undefined || text === '') {
          p.removeChild(run);
          return;
        }
        let t = kid(run, A, 't');
        if (!t) {
          t = doc.createElementNS(A, 'a:t');
          run.appendChild(t);
        }
        while (t.firstChild) t.removeChild(t.firstChild);
        t.appendChild(doc.createTextNode(text));
      });
      if (edit.extra) {
        const r = doc.createElementNS(A, 'a:r');
        if (template) {
          const rPr = doc.createElementNS(A, 'a:rPr');
          for (const attr of Array.from(template.attributes)) rPr.setAttribute(attr.name, attr.value);
          for (const c of kids(template)) rPr.appendChild(c.cloneNode(true));
          r.appendChild(rPr);
        }
        const t = doc.createElementNS(A, 'a:t');
        t.appendChild(doc.createTextNode(edit.extra));
        r.appendChild(t);
        p.insertBefore(r, kid(p, A, 'endParaRPr') ?? null);
      }
    }
    writeXml(zip, paths[s], doc);
  }
  return zipBytes(zip);
}

export const pptxHandler: FormatHandler = { script: 'pptx', load, save: (base, payload) => save(base, payload) };
