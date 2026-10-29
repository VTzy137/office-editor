import type JSZip from 'jszip';
import type { FormatHandler } from '../provider';
import { NS, all, imageDataUri, is, kid, kids, openZip, readRels, readXml, wattr, writeXml, zipBytes } from './ooxml';

const W = NS.w;
const WP = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing';
const XML_NS = 'http://www.w3.org/XML/1998/namespace';

// ---------- Model shared with the webview ----------

export interface RunProps {
  b?: boolean;
  i?: boolean;
  u?: boolean;
  strike?: boolean;
  sz?: number; // points
  color?: string;
  font?: string;
  cstyle?: string;
}
export interface DocRun extends RunProps {
  text: string;
  /** "paragraphId:runIndex" of the source run, used to keep its other properties on save. */
  ref: string;
}
export type Seg = (RunProps & { text: string }) | { img: string; w: number; h: number };
export interface DocPara {
  t: 'p';
  id: number;
  style: string;
  align?: string;
  label?: string;
  indent?: number;
  runs: DocRun[];
  /** Paragraphs with images, fields, equations… are shown read-only. */
  locked?: Seg[];
}
export interface DocTable {
  t: 'table';
  rows: { span: number; blocks: DocBlock[] }[][];
}
export type DocBlock = DocPara | DocTable;
export interface DocxModel {
  css: string;
  styles: { id: string; name: string }[];
  defaultStyle: string;
  blocks: DocBlock[];
}

export interface OutRun {
  text: string;
  ref: string | null;
  b: boolean | null; // null: no explicit value
  i: boolean | null;
  u: boolean;
  sz: number | null; // null: keep the source run's size
}
export interface OutPara {
  id: number | null;
  same?: boolean;
  locked?: boolean;
  style: string;
  runs: OutRun[];
}

// ---------- Helpers ----------

const onOff = (el: Element | undefined): boolean | undefined => {
  if (!el) return undefined;
  const v = wattr(el, 'val');
  return !(v === '0' || v === 'false' || v === 'off');
};

function readRunProps(rPr: Element | undefined, theme: ThemeFonts): RunProps {
  const p: RunProps = {};
  if (!rPr) return p;
  const b = onOff(kid(rPr, W, 'b'));
  if (b !== undefined) p.b = b;
  const i = onOff(kid(rPr, W, 'i'));
  if (i !== undefined) p.i = i;
  const u = kid(rPr, W, 'u');
  if (u) p.u = wattr(u, 'val') !== 'none';
  const strike = onOff(kid(rPr, W, 'strike'));
  if (strike !== undefined) p.strike = strike;
  const sz = wattr(kid(rPr, W, 'sz'), 'val');
  if (sz) p.sz = Number(sz) / 2;
  const color = wattr(kid(rPr, W, 'color'), 'val');
  if (color && color !== 'auto') p.color = '#' + color;
  const fonts = kid(rPr, W, 'rFonts');
  if (fonts) {
    const f = wattr(fonts, 'ascii') || wattr(fonts, 'hAnsi');
    const t = wattr(fonts, 'asciiTheme') || wattr(fonts, 'hAnsiTheme');
    if (f) p.font = f;
    else if (t) p.font = t.startsWith('major') ? theme.major : theme.minor;
  }
  const cs = wattr(kid(rPr, W, 'rStyle'), 'val');
  if (cs) p.cstyle = cs;
  return p;
}

interface ThemeFonts {
  major?: string;
  minor?: string;
}

async function readTheme(zip: JSZip): Promise<ThemeFonts> {
  const doc = await readXml(zip, 'word/theme/theme1.xml');
  const pick = (name: string) => {
    const f = all(doc, NS.a, name)[0];
    return f ? kid(f, NS.a, 'latin')?.getAttribute('typeface') || undefined : undefined;
  };
  return { major: pick('majorFont'), minor: pick('minorFont') };
}

const cssClass = (prefix: string, id: string) => prefix + id.replace(/[^A-Za-z0-9_-]/g, (c) => '_' + c.charCodeAt(0).toString(16));

function propsCss(p: RunProps & { align?: string }): string {
  const out: string[] = [];
  if (p.b !== undefined) out.push(`font-weight:${p.b ? 'bold' : 'normal'}`);
  if (p.i !== undefined) out.push(`font-style:${p.i ? 'italic' : 'normal'}`);
  if (p.u || p.strike) out.push(`text-decoration:${[p.u && 'underline', p.strike && 'line-through'].filter(Boolean).join(' ')}`);
  if (p.sz) out.push(`font-size:${p.sz}pt`);
  if (p.color) out.push(`color:${p.color}`);
  if (p.font) out.push(`font-family:'${p.font.replace(/'/g, '')}',sans-serif`);
  if (p.align) out.push(`text-align:${p.align}`);
  return out.join(';');
}

const ALIGN: Record<string, string> = { left: 'left', start: 'left', center: 'center', right: 'right', end: 'right', both: 'justify', distribute: 'justify' };

interface Styles {
  css: string;
  list: { id: string; name: string }[];
  defaultStyle: string;
  /** List numbering attached to paragraph styles (e.g. "List Bullet"). */
  numbering: Map<string, { numId?: string; ilvl?: string }>;
}

async function readStyles(zip: JSZip, theme: ThemeFonts): Promise<Styles> {
  const doc = await readXml(zip, 'word/styles.xml');
  const defaults = readRunProps(all(doc, W, 'rPrDefault')[0] && kid(all(doc, W, 'rPrDefault')[0], W, 'rPr'), theme);
  const raw = new Map<string, { type: string; name: string; basedOn?: string; props: RunProps & { align?: string }; hidden: boolean; numPr?: Element }>();
  let defaultStyle = 'Normal';
  for (const s of all(doc, W, 'style')) {
    const id = wattr(s, 'styleId');
    if (!id) continue;
    const type = wattr(s, 'type') ?? 'paragraph';
    const props: RunProps & { align?: string } = readRunProps(kid(s, W, 'rPr'), theme);
    const jc = wattr(kid(kid(s, W, 'pPr'), W, 'jc'), 'val');
    if (jc && ALIGN[jc]) props.align = ALIGN[jc];
    raw.set(id, {
      type,
      name: wattr(kid(s, W, 'name'), 'val') ?? id,
      basedOn: wattr(kid(s, W, 'basedOn'), 'val') ?? undefined,
      props,
      hidden: !!kid(s, W, 'semiHidden') || !!kid(s, W, 'hidden'),
      numPr: kid(kid(s, W, 'pPr'), W, 'numPr'),
    });
    if (type === 'paragraph' && (wattr(s, 'default') === '1' || wattr(s, 'default') === 'true')) defaultStyle = id;
  }
  const resolved = new Map<string, RunProps & { align?: string }>();
  const resolve = (id: string, depth = 0): RunProps & { align?: string } => {
    if (resolved.has(id)) return resolved.get(id)!;
    const s = raw.get(id);
    if (!s || depth > 20) return {};
    const merged = { ...(s.basedOn ? resolve(s.basedOn, depth + 1) : {}), ...s.props };
    resolved.set(id, merged);
    return merged;
  };
  const rules: string[] = [`.doc-page{${propsCss({ sz: 11, ...defaults })}}`];
  const list: { id: string; name: string }[] = [];
  for (const [id, s] of raw) {
    if (s.type === 'paragraph') {
      rules.push(`.${cssClass('ps-', id)}{${propsCss(resolve(id))}}`);
      if (!s.hidden || id === defaultStyle) list.push({ id, name: s.name });
    } else if (s.type === 'character') {
      rules.push(`.${cssClass('cs-', id)}{${propsCss(resolve(id))}}`);
    }
  }
  const numbering = new Map<string, { numId?: string; ilvl?: string }>();
  for (const id of raw.keys()) {
    const found: { numId?: string; ilvl?: string } = {};
    for (let cur = raw.get(id), depth = 0; cur && depth < 20; cur = cur.basedOn ? raw.get(cur.basedOn) : undefined, depth++) {
      found.numId ??= wattr(kid(cur.numPr, W, 'numId'), 'val') ?? undefined;
      found.ilvl ??= wattr(kid(cur.numPr, W, 'ilvl'), 'val') ?? undefined;
    }
    if (found.numId) numbering.set(id, found);
  }
  return { css: rules.join('\n'), list, defaultStyle, numbering };
}

interface Numbering {
  label(numId: string, ilvl: number): string;
}

async function readNumbering(zip: JSZip): Promise<Numbering> {
  const doc = await readXml(zip, 'word/numbering.xml');
  const abstract = new Map<string, Element>();
  for (const a of all(doc, W, 'abstractNum')) abstract.set(wattr(a, 'abstractNumId') ?? '', a);
  const numToAbstract = new Map<string, string>();
  for (const n of all(doc, W, 'num')) numToAbstract.set(wattr(n, 'numId') ?? '', wattr(kid(n, W, 'abstractNumId'), 'val') ?? '');
  const counters = new Map<string, number[]>();
  const fmt = (n: number, f: string) => {
    const roman = (x: number) => {
      const map: [number, string][] = [[1000, 'm'], [900, 'cm'], [500, 'd'], [400, 'cd'], [100, 'c'], [90, 'xc'], [50, 'l'], [40, 'xl'], [10, 'x'], [9, 'ix'], [5, 'v'], [4, 'iv'], [1, 'i']];
      let s = '';
      for (const [v, r] of map) while (x >= v) { s += r; x -= v; }
      return s;
    };
    const letter = (x: number) => String.fromCharCode(97 + ((x - 1) % 26));
    switch (f) {
      case 'lowerLetter': return letter(n);
      case 'upperLetter': return letter(n).toUpperCase();
      case 'lowerRoman': return roman(n);
      case 'upperRoman': return roman(n).toUpperCase();
      default: return String(n);
    }
  };
  return {
    label(numId, ilvl) {
      const a = abstract.get(numToAbstract.get(numId) ?? '');
      if (!a || numId === '0') return '';
      const lvls = kids(a, W, 'lvl');
      const lvlOf = (i: number) => lvls.find((l) => wattr(l, 'ilvl') === String(i));
      const lvl = lvlOf(ilvl);
      const numFmt = wattr(kid(lvl, W, 'numFmt'), 'val') ?? 'decimal';
      if (numFmt === 'bullet') return '•';
      if (numFmt === 'none') return '';
      const c = counters.get(numId) ?? [];
      for (let i = 0; i < ilvl; i++) if (!c[i]) c[i] = Number(wattr(kid(lvlOf(i), W, 'start'), 'val') ?? 1);
      c[ilvl] = c[ilvl] ? c[ilvl] + 1 : Number(wattr(kid(lvl, W, 'start'), 'val') ?? 1);
      c.length = ilvl + 1;
      counters.set(numId, c);
      const text = wattr(kid(lvl, W, 'lvlText'), 'val') ?? `%${ilvl + 1}.`;
      return text.replace(/%(\d)/g, (_, d) => {
        const i = Number(d) - 1;
        return fmt(c[i] ?? 1, wattr(kid(lvlOf(i), W, 'numFmt'), 'val') ?? 'decimal');
      });
    },
  };
}

// ---------- Paragraph traversal (must be identical for load and save) ----------

interface ParaInfo {
  el: Element;
  runs: Element[];
  simple: boolean;
}

const SKIP_IN_PARA = new Set(['pPr', 'bookmarkStart', 'bookmarkEnd', 'proofErr', 'permStart', 'permEnd', 'commentRangeStart', 'commentRangeEnd', 'del', 'moveFrom', 'moveFromRangeStart', 'moveFromRangeEnd', 'moveToRangeStart', 'moveToRangeEnd']);
const FLATTEN_IN_PARA = new Set(['hyperlink', 'ins', 'smartTag', 'customXml', 'moveTo', 'fldSimple', 'dir', 'bdo']);
const SIMPLE_RUN_CHILD = new Set(['rPr', 't', 'tab', 'br', 'lastRenderedPageBreak', 'softHyphen', 'noBreakHyphen']);

function paraInfo(p: Element): ParaInfo {
  const runs: Element[] = [];
  let simple = true;
  const walk = (el: Element) => {
    for (const c of kids(el)) {
      if (c.namespaceURI !== W) {
        simple = false; // equations, alternate content, …
        continue;
      }
      const n = c.localName;
      if (SKIP_IN_PARA.has(n)) continue;
      if (n === 'r') {
        for (const rc of kids(c)) {
          const ok = rc.namespaceURI === W && SIMPLE_RUN_CHILD.has(rc.localName) && !(rc.localName === 'br' && wattr(rc, 'type') && wattr(rc, 'type') !== 'textWrapping');
          if (!ok) simple = false;
        }
        runs.push(c);
      } else if (FLATTEN_IN_PARA.has(n)) walk(c);
      else if (n === 'sdt') {
        const content = kid(c, W, 'sdtContent');
        if (content) walk(content);
      } else simple = false;
    }
  };
  walk(p);
  return { el: p, runs, simple };
}

function collectParas(body: Element): ParaInfo[] {
  const out: ParaInfo[] = [];
  const walk = (el: Element) => {
    for (const c of kids(el, W)) {
      if (c.localName === 'p') out.push(paraInfo(c));
      else if (c.localName === 'tbl' || c.localName === 'tr' || c.localName === 'tc' || c.localName === 'sdtContent') walk(c);
      else if (c.localName === 'sdt') walk(c);
    }
  };
  walk(body);
  return out;
}

function runText(r: Element): string {
  let s = '';
  for (const c of kids(r, W)) {
    if (c.localName === 't') s += c.textContent ?? '';
    else if (c.localName === 'tab') s += '\t';
    else if (c.localName === 'br') s += '\n';
    else if (c.localName === 'noBreakHyphen') s += '‑';
  }
  return s;
}

// ---------- Load ----------

async function load(bytes: Uint8Array): Promise<DocxModel> {
  const zip = await openZip(bytes);
  const doc = await readXml(zip, 'word/document.xml');
  const body = doc && all(doc, W, 'body')[0];
  if (!body) throw new Error('Not a valid Word document (word/document.xml missing)');
  const theme = await readTheme(zip);
  const styles = await readStyles(zip, theme);
  const numbering = await readNumbering(zip);
  const rels = await readRels(zip, 'word/document.xml');
  const paras = collectParas(body);
  const index = new Map(paras.map((p, i) => [p.el, i]));

  const lockedSegs = async (p: Element): Promise<Seg[]> => {
    const segs: Seg[] = [];
    const walk = async (el: Element, props: RunProps) => {
      for (const c of kids(el)) {
        if (is(c, W, 'del') || is(c, W, 'pPr') || is(c, W, 'instrText') || is(c, W, 'rPr')) continue;
        if (is(c, W, 'r')) {
          await walk(c, readRunProps(kid(c, W, 'rPr'), theme));
        } else if (is(c, W, 't')) segs.push({ ...props, text: c.textContent ?? '' });
        else if (is(c, W, 'tab')) segs.push({ ...props, text: '\t' });
        else if (is(c, W, 'br')) segs.push({ ...props, text: '\n' });
        else if (is(c, W, 'drawing')) {
          const blip = all(c, NS.a, 'blip')[0];
          const id = blip?.getAttributeNS(NS.r, 'embed') || blip?.getAttribute('r:embed');
          const rel = id ? rels.get(id) : undefined;
          const src = rel && !rel.external ? await imageDataUri(zip, rel.target) : undefined;
          const ext = all(c, WP, 'extent')[0];
          const w = Number(ext?.getAttribute('cx') ?? 0) / 9525;
          const h = Number(ext?.getAttribute('cy') ?? 0) / 9525;
          if (src) segs.push({ img: src, w, h });
          else segs.push({ text: '[image]' });
        } else await walk(c, props);
      }
    };
    await walk(p, {});
    return segs;
  };

  const toPara = async (info: ParaInfo): Promise<DocPara> => {
    const id = index.get(info.el)!;
    const pPr = kid(info.el, W, 'pPr');
    const style = wattr(kid(pPr, W, 'pStyle'), 'val') ?? styles.defaultStyle;
    const jc = wattr(kid(pPr, W, 'jc'), 'val');
    const numPr = kid(pPr, W, 'numPr');
    const para: DocPara = { t: 'p', id, style, runs: [] };
    if (jc && ALIGN[jc]) para.align = ALIGN[jc];
    const styleNum = styles.numbering.get(style);
    const numId = wattr(kid(numPr, W, 'numId'), 'val') ?? styleNum?.numId;
    if (numId) {
      const ilvl = Number(wattr(kid(numPr, W, 'ilvl'), 'val') ?? styleNum?.ilvl ?? 0);
      const label = numbering.label(numId, ilvl);
      if (label) {
        para.label = label;
        para.indent = ilvl;
      }
    }
    if (!info.simple) para.locked = await lockedSegs(info.el);
    else
      para.runs = info.runs
        .map((r, ri) => ({ ...readRunProps(kid(r, W, 'rPr'), theme), text: runText(r), ref: `${id}:${ri}` }))
        .filter((r) => r.text !== '');
    return para;
  };

  const blocksOf = async (el: Element): Promise<DocBlock[]> => {
    const out: DocBlock[] = [];
    for (const c of kids(el, W)) {
      if (c.localName === 'p') out.push(await toPara(paras[index.get(c)!]));
      else if (c.localName === 'sdt') out.push(...(await blocksOf(kid(c, W, 'sdtContent') ?? c)));
      else if (c.localName === 'tbl') {
        const rows: DocTable['rows'] = [];
        for (const tr of kids(c, W, 'tr')) {
          const cells: DocTable['rows'][number] = [];
          for (const tc of kids(tr, W, 'tc')) {
            const span = Number(wattr(kid(kid(tc, W, 'tcPr'), W, 'gridSpan'), 'val') ?? 1);
            cells.push({ span, blocks: await blocksOf(tc) });
          }
          rows.push(cells);
        }
        out.push({ t: 'table', rows });
      }
    }
    return out;
  };

  return { css: styles.css, styles: styles.list, defaultStyle: styles.defaultStyle, blocks: await blocksOf(body) };
}

// ---------- Save ----------

// Schema order of w:rPr children; Word rejects files with out-of-order elements.
const RPR_ORDER = ['rStyle', 'rFonts', 'b', 'bCs', 'i', 'iCs', 'caps', 'smallCaps', 'strike', 'dstrike', 'outline', 'shadow', 'emboss', 'imprint', 'noProof', 'snapToGrid', 'vanish', 'webHidden', 'color', 'spacing', 'w', 'kern', 'position', 'sz', 'szCs', 'highlight', 'u', 'effect', 'bdr', 'shd', 'fitText', 'vertAlign', 'rtl', 'cs', 'em', 'lang', 'eastAsianLayout', 'specVanish', 'oMath'];

function setRPr(rPr: Element, name: string, val: string | null | undefined) {
  for (const c of kids(rPr, W, name)) rPr.removeChild(c);
  if (val === undefined) return;
  const el = rPr.ownerDocument!.createElementNS(W, 'w:' + name);
  if (val !== null) el.setAttributeNS(W, 'w:val', val);
  const order = RPR_ORDER.indexOf(name);
  const before = kids(rPr).find((c) => RPR_ORDER.indexOf(c.localName!) > order);
  rPr.insertBefore(el, before ?? null);
}

function buildRun(doc: Document, out: OutRun, base: Element | undefined): Element {
  const r = doc.createElementNS(W, 'w:r');
  const rPr = base ? (base.cloneNode(true) as Element) : doc.createElementNS(W, 'w:rPr');
  for (const c of kids(rPr, W)) if (c.localName === 'rPrChange' || c.localName === 'ins' || c.localName === 'del') rPr.removeChild(c);
  const hasCharStyle = !!kid(rPr, W, 'rStyle');
  const flag = (name: string, v: boolean | null) => {
    setRPr(rPr, name, undefined);
    setRPr(rPr, name + 'Cs', undefined);
    if (v === null && !hasCharStyle) return;
    setRPr(rPr, name, v ? null : '0');
  };
  flag('b', out.b);
  flag('i', out.i);
  setRPr(rPr, 'u', out.u ? 'single' : undefined);
  if (out.sz !== null) {
    setRPr(rPr, 'sz', String(Math.round(out.sz * 2)));
    setRPr(rPr, 'szCs', String(Math.round(out.sz * 2)));
  }
  if (kids(rPr).length) r.appendChild(rPr);
  for (const piece of out.text.split(/(\t|\n)/)) {
    if (piece === '') continue;
    if (piece === '\t') r.appendChild(doc.createElementNS(W, 'w:tab'));
    else if (piece === '\n') r.appendChild(doc.createElementNS(W, 'w:br'));
    else {
      const t = doc.createElementNS(W, 'w:t');
      t.setAttributeNS(XML_NS, 'xml:space', 'preserve');
      t.appendChild(doc.createTextNode(piece));
      r.appendChild(t);
    }
  }
  return r;
}

function setStyle(doc: Document, p: Element, style: string, defaultStyle: string) {
  let pPr = kid(p, W, 'pPr');
  const current = wattr(kid(pPr, W, 'pStyle'), 'val') ?? defaultStyle;
  if (current === style) return;
  if (!pPr) {
    pPr = doc.createElementNS(W, 'w:pPr');
    p.insertBefore(pPr, p.firstChild);
  }
  for (const c of kids(pPr, W, 'pStyle')) pPr.removeChild(c);
  if (style === defaultStyle) return;
  const ps = doc.createElementNS(W, 'w:pStyle');
  ps.setAttributeNS(W, 'w:val', style);
  pPr.insertBefore(ps, pPr.firstChild);
}

async function save(base: Uint8Array, payload: { paras: OutPara[] }): Promise<Uint8Array> {
  const zip = await openZip(base);
  const doc = (await readXml(zip, 'word/document.xml'))!;
  const body = all(doc, W, 'body')[0];
  const styles = await readStyles(zip, {});
  const paras = collectParas(body);

  const sourceRPr = (ref: string | null, fallbackPara: Element): Element | undefined => {
    if (ref) {
      const [pid, ri] = ref.split(':').map(Number);
      const run = paras[pid]?.runs[ri];
      if (run) return kid(run, W, 'rPr');
    }
    // A run typed into an empty paragraph takes the paragraph mark's formatting.
    const mark = kid(kid(fallbackPara, W, 'pPr'), W, 'rPr');
    if (!mark) return undefined;
    const clone = mark.cloneNode(true) as Element;
    for (const c of kids(clone, W)) if (c.localName === 'ins' || c.localName === 'del' || c.localName === 'rPrChange') clone.removeChild(c);
    return clone;
  };

  const fill = (p: Element, entry: OutPara) => {
    for (const c of kids(p)) if (!is(c, W, 'pPr')) p.removeChild(c);
    for (const n of Array.from(p.childNodes)) if (n.nodeType !== 1) p.removeChild(n);
    for (const run of entry.runs) if (run.text) p.appendChild(buildRun(doc, run, sourceRPr(run.ref, p)));
    setStyle(doc, p, entry.style, styles.defaultStyle);
  };

  const seen = new Set<number>();
  let prev: Element | undefined;
  for (const entry of payload.paras) {
    if (entry.id !== null && !seen.has(entry.id) && paras[entry.id]) {
      seen.add(entry.id);
      const p = paras[entry.id].el;
      if (!entry.same && !entry.locked && paras[entry.id].simple) fill(p, entry);
      prev = p;
      continue;
    }
    if (entry.locked) continue; // duplicated read-only blocks can't be recreated
    // New paragraph (e.g. Enter pressed): inherit paragraph properties from its source, without section breaks.
    const template = entry.id !== null ? paras[entry.id]?.el : prev;
    const p = doc.createElementNS(W, 'w:p');
    const pPr = kid(template, W, 'pPr');
    if (pPr) {
      const clone = pPr.cloneNode(true) as Element;
      for (const c of kids(clone, W)) if (c.localName === 'sectPr' || c.localName === 'pPrChange') clone.removeChild(c);
      p.appendChild(clone);
    }
    fill(p, entry);
    if (prev) prev.parentNode!.insertBefore(p, prev.nextSibling);
    else body.insertBefore(p, body.firstChild);
    prev = p;
  }

  paras.forEach((info, id) => {
    if (seen.has(id)) return;
    const parent = info.el.parentNode as Element;
    // Table cells must keep at least one paragraph.
    if (is(parent, W, 'tc') && kids(parent, W, 'p').length === 1) fill(info.el, { id, style: styles.defaultStyle, runs: [] });
    else parent.removeChild(info.el);
  });

  writeXml(zip, 'word/document.xml', doc);
  return zipBytes(zip);
}

export const docxHandler: FormatHandler = {
  script: 'docx',
  load,
  save: (base, payload) => save(base, payload),
};
