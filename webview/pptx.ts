import { start, markDirty, h, plainTextPaste } from './common';
import type { PElement, PPara, PRun, PptxModel, PptxPayload } from '../src/formats/pptx';

const app = document.getElementById('app')!;
let model: PptxModel;
let deck: HTMLDivElement;
const original = new Map<string, string>();

function renderRun(run: PRun): HTMLElement {
  const span = h('span', run.r !== undefined ? { 'data-r': run.r } : {}, run.text);
  span.style.fontSize = `${run.sz}px`;
  if (run.b) span.style.fontWeight = 'bold';
  if (run.i) span.style.fontStyle = 'italic';
  if (run.u) span.style.textDecoration = 'underline';
  if (run.color) span.style.color = run.color;
  if (run.font) span.style.fontFamily = `'${run.font}', 'Segoe UI', sans-serif`;
  return span;
}

function renderPara(p: PPara): HTMLElement {
  const el = h('div', { class: 'slide-para' });
  el.style.textAlign = p.align;
  el.style.fontSize = `${p.sz}px`;
  if (p.level) el.style.marginLeft = `${p.level * 0.6}em`;
  if (p.bullet) el.dataset.bullet = p.bullet;
  for (const item of p.items) el.appendChild('br' in item ? document.createElement('br') : renderRun(item));
  if (!p.items.length) el.appendChild(document.createElement('br'));
  if (p.id) {
    el.dataset.id = p.id;
    el.contentEditable = 'true';
    el.spellcheck = false;
    plainTextPaste(el);
    el.addEventListener('input', () => markDirty());
    el.addEventListener('keydown', (e) => {
      // Paragraph structure is kept as-is; only text inside runs is editable.
      if (e.key === 'Enter') e.preventDefault();
      if ((e.ctrlKey || e.metaKey) && ['b', 'i', 'u'].includes(e.key.toLowerCase())) {
        e.preventDefault();
        e.stopPropagation();
      }
    });
  }
  return el;
}

function place(el: HTMLElement, e: { x: number; y: number; w: number; h: number; rot?: number }) {
  el.style.left = `${e.x}px`;
  el.style.top = `${e.y}px`;
  el.style.width = `${e.w}px`;
  el.style.height = `${e.h}px`;
  if (e.rot) el.style.transform = `rotate(${e.rot}deg)`;
  return el;
}

function renderElement(e: PElement): HTMLElement {
  switch (e.k) {
    case 'img':
      return place(h('img', { class: 'slide-img', src: e.src, draggable: 'false' }), e);
    case 'placeholder':
      return place(h('div', { class: 'slide-object' }, `[${e.label}]`), e);
    case 'text': {
      const box = place(h('div', { class: `slide-text anchor-${e.anchor}` }), e);
      if (e.fill) box.style.background = e.fill;
      if (e.vert) box.style.writingMode = 'vertical-rl';
      box.append(...e.paras.map(renderPara));
      return box;
    }
    case 'table': {
      const wrap = place(h('div', { class: 'slide-table-wrap' }), e);
      const table = h('table', { class: 'slide-table' });
      const colgroup = h('colgroup', {});
      for (const w of e.cols) {
        const col = h('col', {});
        col.style.width = `${w}px`;
        colgroup.appendChild(col);
      }
      table.appendChild(colgroup);
      for (const row of e.rows) {
        const tr = table.insertRow();
        tr.style.height = `${row.h}px`;
        for (const cell of row.cells) {
          const td = tr.insertCell();
          td.colSpan = cell.span;
          td.rowSpan = cell.rowSpan;
          if (cell.fill) td.style.background = cell.fill;
          td.append(...cell.paras.map(renderPara));
        }
      }
      wrap.appendChild(table);
      return wrap;
    }
  }
}

function fit() {
  const available = deck.clientWidth - 48;
  const scale = Math.min(1.5, available / model.width);
  deck.style.setProperty('--scale', String(scale));
}

function serializePara(el: HTMLElement): { runs: Record<number, string>; extra: string | null } {
  const runs: Record<number, string> = {};
  for (const span of Array.from(el.querySelectorAll<HTMLElement>('[data-r]'))) runs[Number(span.dataset.r)] = '';
  let current: number | null = null;
  let prefix = '';
  const walker = document.createTreeWalker(el, NodeFilter.SHOW_TEXT);
  for (let n = walker.nextNode(); n; n = walker.nextNode()) {
    const text = (n.nodeValue ?? '').replace(/ /g, ' ');
    const owner = n.parentElement?.closest<HTMLElement>('[data-r]');
    if (owner && el.contains(owner)) current = Number(owner.dataset.r);
    // Text typed outside any run joins the run before it (or the first run).
    if (current === null) prefix += text;
    else runs[current] += text;
  }
  const keys = Object.keys(runs).map(Number).sort((a, b) => a - b);
  if (prefix && keys.length) {
    runs[keys[0]] = prefix + runs[keys[0]];
    prefix = '';
  }
  return { runs, extra: prefix || null };
}

start<PptxModel>({
  init(m, fileName) {
    model = m;
    deck = h('div', { class: 'deck' });
    m.slides.forEach((s, i) => {
      const slide = h('div', { class: 'slide' });
      slide.style.width = `${m.width}px`;
      slide.style.height = `${m.height}px`;
      slide.style.background = s.bg;
      if (s.bgImage) {
        slide.style.backgroundImage = `url("${s.bgImage}")`;
        slide.style.backgroundSize = 'cover';
      }
      slide.append(...s.elements.map(renderElement));
      const frame = h('div', { class: 'slide-frame' }, slide);
      frame.style.width = `calc(${m.width}px * var(--scale))`;
      frame.style.height = `calc(${m.height}px * var(--scale))`;
      deck.appendChild(h('div', { class: 'slide-wrap' + (s.hidden ? ' hidden-slide' : '') }, h('div', { class: 'slide-num' }, `Slide ${i + 1}${s.hidden ? ' (hidden)' : ''}`), frame));
    });
    if (!m.slides.length) deck.appendChild(h('div', { class: 'loading' }, 'This presentation has no slides.'));
    app.replaceChildren(
      h('div', { class: 'pptx-editor' }, h('div', { class: 'toolbar' }, h('span', { class: 'tb-title' }, fileName), h('span', { class: 'tb-hint' }, `${m.slides.length} slides · click any text to edit`)), deck),
    );
    original.clear();
    for (const p of Array.from(deck.querySelectorAll<HTMLElement>('[data-id]'))) original.set(p.dataset.id!, JSON.stringify(serializePara(p)));
    fit();
    new ResizeObserver(fit).observe(deck);
  },
  getContent(): PptxPayload {
    const paras: PptxPayload['paras'] = [];
    for (const p of Array.from(deck.querySelectorAll<HTMLElement>('[data-id]'))) {
      const s = serializePara(p);
      if (JSON.stringify(s) !== original.get(p.dataset.id!)) paras.push({ id: p.dataset.id!, ...s });
    }
    return { paras };
  },
});
