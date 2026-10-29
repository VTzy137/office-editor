import { start, markDirty, h, button, plainTextPaste } from './common';
import type { DocBlock, DocPara, DocxModel, OutPara, OutRun, RunProps, Seg } from '../src/formats/docx';

const app = document.getElementById('app')!;
let page: HTMLDivElement;
let styleSelect: HTMLSelectElement;
let sizeSelect: HTMLSelectElement;
let defaultStyle = 'Normal';
let savedRange: Range | null = null;
/** Serialized form of each paragraph as loaded, to detect untouched ones. */
const original = new Map<number, string>();

const cssClass = (prefix: string, id: string) => prefix + id.replace(/[^A-Za-z0-9_-]/g, (c) => '_' + c.charCodeAt(0).toString(16));

function applyProps(el: HTMLElement, p: RunProps) {
  if (p.b !== undefined) el.style.fontWeight = p.b ? 'bold' : 'normal';
  if (p.i !== undefined) el.style.fontStyle = p.i ? 'italic' : 'normal';
  const deco = [p.u && 'underline', p.strike && 'line-through'].filter(Boolean).join(' ');
  if (deco) el.style.textDecoration = deco;
  if (p.sz) el.style.fontSize = `${p.sz}pt`;
  if (p.color) el.style.color = p.color;
  if (p.font) el.style.fontFamily = `'${p.font}', sans-serif`;
  if (p.cstyle) el.classList.add(cssClass('cs-', p.cstyle));
}

function setParaStyle(p: HTMLElement, style: string) {
  for (const c of Array.from(p.classList)) if (c.startsWith('ps-')) p.classList.remove(c);
  p.classList.add(cssClass('ps-', style));
  p.dataset.style = style;
}

function renderSegs(segs: Seg[]): Node[] {
  return segs.map((s) => {
    if ('img' in s) {
      const img = h('img', { src: s.img, class: 'doc-img' });
      if (s.w) img.style.width = `${s.w}px`;
      return img;
    }
    const span = h('span', {}, s.text);
    applyProps(span, s);
    return span;
  });
}

function renderPara(p: DocPara): HTMLElement {
  const el = h('p', { 'data-pid': p.id });
  setParaStyle(el, p.style);
  if (p.align) el.style.textAlign = p.align;
  if (p.label) {
    el.dataset.label = p.label;
    el.style.paddingLeft = `${(p.indent ?? 0) * 24 + 28}px`;
  }
  if (p.locked) {
    el.contentEditable = 'false';
    el.dataset.locked = '';
    el.title = 'Contains images, fields or other objects — shown read-only';
    el.append(...renderSegs(p.locked));
  } else {
    for (const r of p.runs) {
      const span = h('span', { 'data-ref': r.ref }, r.text);
      applyProps(span, r);
      el.appendChild(span);
    }
  }
  if (!el.childNodes.length) el.appendChild(document.createElement('br'));
  return el;
}

function renderBlocks(blocks: DocBlock[]): Node[] {
  return blocks.map((b) => {
    if (b.t === 'p') return renderPara(b);
    const table = h('table', { class: 'doc-table' });
    for (const row of b.rows) {
      const tr = table.insertRow();
      for (const cell of row) {
        const td = tr.insertCell();
        if (cell.span > 1) td.colSpan = cell.span;
        td.append(...renderBlocks(cell.blocks));
      }
    }
    return table;
  });
}

// ---------- Serialization ----------

const isBold = (s: CSSStyleDeclaration) => Number(s.fontWeight) >= 600 || s.fontWeight === 'bold';
const isItalic = (s: CSSStyleDeclaration) => s.fontStyle === 'italic' || s.fontStyle === 'oblique';

function serializeRuns(p: HTMLElement): OutRun[] {
  const paraStyle = getComputedStyle(p);
  const paraBold = isBold(paraStyle);
  const paraItalic = isItalic(paraStyle);
  const runs: OutRun[] = [];
  const push = (text: string, owner: HTMLElement) => {
    const cs = getComputedStyle(owner);
    let underline = false;
    for (let e: HTMLElement | null = owner; e && e !== p; e = e.parentElement) {
      if (getComputedStyle(e).textDecorationLine.includes('underline')) underline = true;
    }
    const sized = owner.closest<HTMLElement>('[data-sz]');
    const ref = owner.closest<HTMLElement>('[data-ref]');
    const bold = isBold(cs);
    const italic = isItalic(cs);
    const run: OutRun = {
      text,
      ref: ref && p.contains(ref) ? ref.dataset.ref! : null,
      b: bold !== paraBold ? bold : null,
      i: italic !== paraItalic ? italic : null,
      u: underline,
      sz: sized && p.contains(sized) ? Number(sized.dataset.sz) : null,
    };
    const last = runs[runs.length - 1];
    if (last && last.ref === run.ref && last.b === run.b && last.i === run.i && last.u === run.u && last.sz === run.sz) last.text += text;
    else runs.push(run);
  };
  const walker = document.createTreeWalker(p, NodeFilter.SHOW_TEXT | NodeFilter.SHOW_ELEMENT);
  for (let n = walker.nextNode(); n; n = walker.nextNode()) {
    if (n.nodeType === Node.TEXT_NODE) {
      if (n.nodeValue) push(n.nodeValue.replace(/ /g, ' ').replace(/​/g, ''), n.parentElement!);
    } else if ((n as Element).tagName === 'BR') {
      // A trailing <br> is only the browser's placeholder for an empty line.
      const isLast = !nextContent(n, p);
      if (!isLast) push('\n', (n as Element).parentElement!);
    }
  }
  return runs.filter((r) => r.text !== '');
}

function nextContent(n: Node, root: Node): boolean {
  for (let cur: Node | null = n; cur && cur !== root; cur = cur.parentNode) {
    for (let s = cur.nextSibling; s; s = s.nextSibling) if (s.textContent || (s as Element).tagName === 'BR' || (s as Element).querySelector?.('br')) return true;
  }
  return false;
}

function serializePara(p: HTMLElement): OutPara {
  const id = p.dataset.pid !== undefined ? Number(p.dataset.pid) : null;
  const style = p.dataset.style ?? defaultStyle;
  if (p.dataset.locked !== undefined) return { id, locked: true, style, runs: [] };
  return { id, style, runs: serializeRuns(p) };
}

const BLOCK_TAGS = new Set(['P', 'DIV', 'H1', 'H2', 'H3', 'H4', 'H5', 'H6', 'LI', 'BLOCKQUOTE', 'PRE']);
const CONTAINER_TAGS = new Set(['TABLE', 'TBODY', 'THEAD', 'TFOOT', 'TR', 'TD', 'TH', 'UL', 'OL']);

/** Wrap text the browser left directly in a container into a paragraph so it isn't lost. */
function wrapStrays(container: HTMLElement) {
  let wrapper: HTMLElement | null = null;
  for (const n of Array.from(container.childNodes)) {
    const el = n as HTMLElement;
    const isBlock = n.nodeType === Node.ELEMENT_NODE && (BLOCK_TAGS.has(el.tagName) || CONTAINER_TAGS.has(el.tagName));
    if (isBlock || (n.nodeType === Node.TEXT_NODE && !n.nodeValue?.trim())) {
      wrapper = null;
      continue;
    }
    if (!wrapper) {
      wrapper = h('p', {});
      setParaStyle(wrapper, defaultStyle);
      container.insertBefore(wrapper, n);
    }
    wrapper.appendChild(n);
  }
}

function collect(container: HTMLElement, out: OutPara[]) {
  if (container === page || container.tagName === 'TD' || container.tagName === 'TH') wrapStrays(container);
  for (const child of Array.from(container.children) as HTMLElement[]) {
    if (child.dataset.pid !== undefined) out.push(serializePara(child));
    else if (CONTAINER_TAGS.has(child.tagName) || child.querySelector('[data-pid]')) collect(child, out);
    else if (BLOCK_TAGS.has(child.tagName)) out.push(serializePara(child));
  }
}

function serialize(): OutPara[] {
  const out: OutPara[] = [];
  collect(page, out);
  const seen = new Set<number>();
  for (const p of out) {
    if (p.id === null || seen.has(p.id)) continue;
    seen.add(p.id);
    if (!p.locked && original.get(p.id) === JSON.stringify([p.style, p.runs])) {
      p.same = true;
      p.runs = [];
    }
  }
  return out;
}

// ---------- Toolbar ----------

function restoreSelection() {
  page.focus({ preventScroll: true });
  if (savedRange) {
    const sel = getSelection()!;
    sel.removeAllRanges();
    sel.addRange(savedRange);
  }
}

function currentParas(): HTMLElement[] {
  const sel = getSelection();
  if (!sel?.rangeCount) return [];
  const range = sel.getRangeAt(0);
  const paraOf = (n: Node) => (n.nodeType === Node.ELEMENT_NODE ? (n as Element) : n.parentElement)?.closest<HTMLElement>('p, div:not(.doc-page), h1, h2, h3, li');
  const start = paraOf(range.startContainer);
  if (range.collapsed) return start && page.contains(start) ? [start] : [];
  return (Array.from(page.querySelectorAll('p, div, h1, h2, h3, li')) as HTMLElement[]).filter((p) => range.intersectsNode(p) && !p.querySelector('p'));
}

function command(cmd: string) {
  restoreSelection();
  document.execCommand(cmd);
  markDirty();
}

function applySize(pt: number) {
  restoreSelection();
  const sel = getSelection();
  if (!sel?.rangeCount || sel.getRangeAt(0).collapsed) return;
  document.execCommand('styleWithCSS', false, 'false');
  document.execCommand('fontSize', false, '7');
  for (const font of Array.from(page.querySelectorAll('font[size="7"]'))) {
    const span = h('span', { 'data-sz': pt });
    span.style.fontSize = `${pt}pt`;
    span.append(...Array.from(font.childNodes));
    font.replaceWith(span);
    for (const inner of Array.from(span.querySelectorAll<HTMLElement>('[data-sz], [style*="font-size"]'))) {
      delete inner.dataset.sz;
      inner.style.fontSize = '';
    }
  }
  markDirty();
}

function applyStyle(style: string) {
  restoreSelection();
  for (const p of currentParas()) {
    if (p.dataset.locked !== undefined) continue;
    setParaStyle(p, style);
  }
  markDirty();
}

function syncToolbar() {
  const sel = getSelection();
  // Ignore selection changes caused by toolbar controls taking focus.
  if (!sel?.rangeCount || !page.contains(sel.anchorNode) || !page.contains(document.activeElement)) return;
  savedRange = sel.getRangeAt(0).cloneRange();
  const para = currentParas()[0];
  if (para?.dataset.style) styleSelect.value = para.dataset.style;
  const anchor = sel.anchorNode?.nodeType === Node.TEXT_NODE ? sel.anchorNode.parentElement : (sel.anchorNode as HTMLElement);
  if (anchor) sizeSelect.value = String(Math.round((parseFloat(getComputedStyle(anchor).fontSize) * 72) / 96));
}

function buildToolbar(model: DocxModel, fileName: string): HTMLElement {
  styleSelect = h('select', { class: 'tb-select', title: 'Paragraph style' });
  for (const s of model.styles) styleSelect.appendChild(h('option', { value: s.id }, s.name));
  styleSelect.addEventListener('change', () => applyStyle(styleSelect.value));

  sizeSelect = h('select', { class: 'tb-select', title: 'Font size' });
  const sizes = [8, 9, 10, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72];
  for (const s of sizes) sizeSelect.appendChild(h('option', { value: s }, String(s)));
  sizeSelect.addEventListener('mousedown', () => syncToolbar());
  sizeSelect.addEventListener('change', () => applySize(Number(sizeSelect.value)));

  return h(
    'div',
    { class: 'toolbar' },
    h('span', { class: 'tb-title' }, fileName),
    h('span', { class: 'tb-sep' }),
    styleSelect,
    sizeSelect,
    button('B', 'Bold (Ctrl+B)', () => command('bold'), 'fmt-b'),
    button('I', 'Italic (Ctrl+I)', () => command('italic'), 'fmt-i'),
    button('U', 'Underline (Ctrl+U)', () => command('underline'), 'fmt-u'),
  );
}

start<DocxModel>({
  init(model, fileName) {
    defaultStyle = model.defaultStyle;
    let style = document.getElementById('doc-styles');
    if (!style) {
      style = h('style', { id: 'doc-styles' });
      document.head.appendChild(style);
    }
    style.textContent = model.css;

    page = h('div', { class: 'doc-page', contenteditable: 'true', spellcheck: 'false' });
    page.append(...renderBlocks(model.blocks));
    if (!page.childNodes.length) page.appendChild(renderPara({ t: 'p', id: -1, style: defaultStyle, runs: [] }));
    const toolbar = buildToolbar(model, fileName);
    app.replaceChildren(h('div', { class: 'doc-editor' }, toolbar, h('div', { class: 'doc-scroll' }, page)));

    document.execCommand('styleWithCSS', false, 'false');
    original.clear();
    for (const p of Array.from(page.querySelectorAll<HTMLElement>('[data-pid]'))) {
      const s = serializePara(p);
      if (s.id !== null) original.set(s.id, JSON.stringify([s.style, s.runs]));
    }

    plainTextPaste(page);
    page.addEventListener('input', () => markDirty());
    page.addEventListener('keydown', (e) => {
      // Handle formatting shortcuts here so VS Code doesn't run its own bindings (e.g. Ctrl+B toggles the sidebar).
      if ((e.ctrlKey || e.metaKey) && !e.altKey && !e.shiftKey) {
        const cmd = ({ b: 'bold', i: 'italic', u: 'underline' } as Record<string, string>)[e.key.toLowerCase()];
        if (cmd) {
          e.preventDefault();
          e.stopPropagation();
          document.execCommand(cmd);
          markDirty();
        }
      }
    });
    // selectionchange is async; also capture the selection synchronously so a quick toolbar click sees it.
    document.addEventListener('selectionchange', syncToolbar);
    for (const ev of ['mouseup', 'keyup', 'dblclick']) page.addEventListener(ev, syncToolbar);
  },
  getContent() {
    return { paras: serialize() };
  },
});
